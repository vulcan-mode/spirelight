// Run this in the browser DevTools Console while signed in at
// https://voice.spirelight.ai/referral/people
// It downloads the currently rendered People table as a local JSON file.
// It does not send the records anywhere.
(() => {
  const requiredHeaders = ["name", "country, region", "progress"];
  const tables = [...document.querySelectorAll("table")];
  // Some header cells render visually all-caps (via a wrapper that only
  // affects innerText, not textContent), so match case-insensitively on
  // textContent rather than innerText.
  const table = tables.find((candidate) => {
    const headers = [...candidate.querySelectorAll("th")].map((cell) =>
      cell.textContent.trim().toLowerCase()
    );
    return requiredHeaders.every((header) => headers.includes(header));
  });

  if (!table) {
    console.error("Export stopped: could not find the expected People table.");
    return;
  }

  const headers = [...table.querySelectorAll("th")].map((cell, index) =>
    cell.textContent.trim() || `column_${index + 1}`
  );
  const rows = [...table.querySelectorAll("tbody tr")];
  if (rows.length === 0) {
    console.error("Export stopped: the People table has no visible rows.");
    return;
  }

  const records = [];
  const malformedRows = [];
  rows.forEach((row, rowIndex) => {
    const cells = [...row.querySelectorAll(":scope > td")];
    if (cells.length !== headers.length) {
      malformedRows.push({ row: rowIndex + 1, cells: cells.length });
      return;
    }

    const record = {};
    headers.forEach((header, cellIndex) => {
      record[header] = cells[cellIndex].innerText
        // The site renders some adjacent labels in separate elements without
        // whitespace (for example country + region or progress + substatus).
        .replace(/([a-záéíóúüñ])([A-ZÁÉÍÓÚÜÑ])/g, "$1 $2")
        .replace(/\s+/g, " ")
        .trim();
    });
    records.push(record);
  });

  if (records.length === 0) {
    console.error("Export stopped: no rows matched the table headers.");
    return;
  }

  const pageText = document.body.innerText;
  const rangeMatch = pageText.match(/showing\s+(\d+)\s*(?:-|–|to)\s*(\d+)\s+of\s+(\d+)/i);
  const paginationControls = [...document.querySelectorAll("button, a")]
    .map((element) => ({
      label: [element.getAttribute("aria-label"), element.title, element.innerText]
        .filter(Boolean)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
      disabled: element.matches(":disabled,[aria-disabled='true']")
    }))
    .filter(({ label }) => /\b(next|previous|load more)\b|page\s+\d+/i.test(label));

  const totalRows = rangeMatch ? Number(rangeMatch[3]) : null;
  const warnings = [];
  if (malformedRows.length) {
    warnings.push(`${malformedRows.length} row(s) had a different number of cells and were skipped.`);
  }
  if (totalRows !== null && totalRows > records.length) {
    warnings.push(`The page reports ${totalRows} total rows; this export contains ${records.length}.`);
  }
  if (paginationControls.some((control) => !control.disabled)) {
    warnings.push("Pagination controls are present; verify that this export includes every page.");
  }

  const exportedAt = new Date().toISOString();
  const payload = {
    source: location.origin + location.pathname,
    exportedAt,
    pageTitle: document.title,
    columns: headers,
    visibleRowCount: records.length,
    reportedTotalRowCount: totalRows,
    paginationControls,
    warnings,
    records
  };

  const datePart = exportedAt.replace(/[:.]/g, "-");
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json"
  });
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = `spirelight-people-${datePart}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(downloadUrl), 60_000);

  console.info(`Exported ${records.length} row(s) to ${link.download}.`);
  if (warnings.length) console.warn("Review export warnings:", warnings);
  if (malformedRows.length) console.warn("Skipped malformed rows:", malformedRows);
})();
