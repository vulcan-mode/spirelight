// Scheduled sync: logs into voice.spirelight.ai/referral/people with a
// saved session cookie, scrapes the People table (same extraction logic
// as export-spirelight-people.js, adapted for headless/CI use), and POSTs
// the rows to the Apps Script backend (see apps-script/Code.gs's
// handlePeopleSync_), which overwrites the "People" tab.
//
// Required env vars:
//   SPIRELIGHT_COOKIE   - the full `Cookie:` request header value copied
//                         from DevTools (Network tab > the /referral/people
//                         document request > Request Headers > cookie).
//                         Needed because the auth cookie is httpOnly and
//                         can't be read from page JS or document.cookie.
//   PEOPLE_SYNC_URL     - the Apps Script Web app URL (same one used by
//                         referral-form/index.html's SCRIPT_URL).
//   PEOPLE_SYNC_SECRET  - shared secret, bootstrapped once into Script
//                         Properties (see Code.gs's file header comment).
//
// Cookies expire/rotate periodically -- when they do, this script's login
// check below fails loudly and the GitHub Actions run shows as failed
// (GitHub emails the repo owner), which is the signal to refresh
// SPIRELIGHT_COOKIE by hand.

import { chromium } from 'playwright';

const PEOPLE_URL = 'https://voice.spirelight.ai/referral/people';

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required env var: ${name}`);
  return value;
}

// Splits the trailing "name email" concatenation the page renders (two
// separately-styled lines joined by innerText with a single space) into
// a clean Name + Email pair. When no real name was ever entered, the
// page shows the email on both lines, so this collapses to Email again.
function splitNameEmail(raw) {
  const match = raw.match(/^(.*?)\s+(\S+@\S+)$/);
  if (!match) return { name: raw, email: '' };
  return { name: match[1].trim() || match[2], email: match[2] };
}

async function scrapePeopleTable(page) {
  const result = await page.evaluate(() => {
    const requiredHeaders = ['name', 'country, region', 'progress'];
    const tables = [...document.querySelectorAll('table')];
    // Some header cells render visually all-caps (via a wrapper that only
    // affects innerText, not textContent), so match case-insensitively on
    // textContent rather than innerText.
    const table = tables.find((candidate) => {
      const headers = [...candidate.querySelectorAll('th')].map((cell) =>
        cell.textContent.trim().toLowerCase()
      );
      return requiredHeaders.every((header) => headers.includes(header));
    });
    if (!table) return { error: 'table not found', tableCount: tables.length };

    const headers = [...table.querySelectorAll('th')].map(
      (cell, index) => cell.textContent.trim() || `column_${index + 1}`
    );
    const rows = [...table.querySelectorAll('tbody tr')];
    const records = [];
    const malformedRows = [];
    rows.forEach((row, rowIndex) => {
      const cells = [...row.querySelectorAll(':scope > td')];
      if (cells.length !== headers.length) {
        malformedRows.push({ row: rowIndex + 1, cells: cells.length });
        return;
      }
      const record = {};
      headers.forEach((header, cellIndex) => {
        record[header] = cells[cellIndex].innerText
          .replace(/([a-záéíóúüñ])([A-ZÁÉÍÓÚÜÑ])/g, '$1 $2')
          .replace(/\s+/g, ' ')
          .trim();
      });
      records.push(record);
    });

    const pageText = document.body.innerText;
    const rangeMatch = pageText.match(/showing\s+(\d+)\s*(?:-|–|to)\s*(\d+)\s+of\s+(\d+)/i);
    return {
      records,
      malformedRows,
      visibleRowCount: records.length,
      reportedTotalRowCount: rangeMatch ? Number(rangeMatch[3]) : null
    };
  });

  return result;
}

async function postToSheet(records, syncUrl, secret) {
  const body = new URLSearchParams({
    formType: 'people_sync',
    secret,
    records: JSON.stringify(records)
  });
  const response = await fetch(syncUrl, { method: 'POST', body });
  const text = await response.text();
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error(`Apps Script returned non-JSON response: ${text.slice(0, 300)}`);
  }
  if (parsed.result !== 'success') {
    throw new Error(`Apps Script sync failed: ${parsed.message || text}`);
  }
  return parsed;
}

async function main() {
  const cookie = requireEnv('SPIRELIGHT_COOKIE');
  const syncUrl = requireEnv('PEOPLE_SYNC_URL');
  const secret = requireEnv('PEOPLE_SYNC_SECRET');

  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.setExtraHTTPHeaders({ cookie });
    await page.goto(PEOPLE_URL, { waitUntil: 'domcontentloaded' });

    // Next.js RSC app -- domcontentloaded fires before client hydration paints
    // anything, so this has to actually wait (isVisible() alone only checks
    // the DOM's current state and returns instantly, which was the bug here:
    // it always ran before hydration finished and always reported "not
    // logged in" regardless of whether the cookie was actually valid).
    const loggedIn = await page
      .getByText('People you referred', { exact: false })
      .first()
      .waitFor({ state: 'visible', timeout: 20000 })
      .then(() => true)
      .catch(() => false);
    if (!loggedIn) {
      throw new Error(
        `Not logged in (page did not show the expected heading within 20s) -- landed on ${page.url()}. SPIRELIGHT_COOKIE has likely expired and needs refreshing.`
      );
    }

    await page.waitForSelector('table th', { timeout: 20000 });
    const scraped = await scrapePeopleTable(page);
    if (scraped.error) {
      throw new Error(`Scrape failed: ${scraped.error} (tableCount=${scraped.tableCount})`);
    }
    if (scraped.malformedRows.length) {
      console.warn('Skipped malformed rows:', scraped.malformedRows);
    }
    if (
      scraped.reportedTotalRowCount !== null &&
      scraped.reportedTotalRowCount !== scraped.visibleRowCount
    ) {
      console.warn(
        `Page reports ${scraped.reportedTotalRowCount} total rows; only captured ${scraped.visibleRowCount} (pagination/load-more may need handling).`
      );
    }

    const records = scraped.records.map((record) => {
      const { name, email } = splitNameEmail(record['Name'] || '');
      return { ...record, Name: name, Email: email };
    });

    const result = await postToSheet(records, syncUrl, secret);
    console.log(`Synced ${result.count} record(s) to the People tab.`);
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
