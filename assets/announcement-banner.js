// Shared site-wide announcement banner. One include, one place to fix
// things -- see the "getting complicated" conversation: this replaces
// what would otherwise be the same banner logic copy-pasted into every
// page's own <script> block.
//
// Controlled entirely from the "Avisos" tab in the Sheet (Activo,
// Mensaje, Tipo, Enlace columns) -- no code change needed to post or
// clear an announcement, just edit the sheet.
(function () {
  var WORKER_URL = "https://spirelight-leads.domlebo1.workers.dev";
  var STYLE_ID = 'announcement-banner-style';
  var TYPE_COLORS = {
    info: { bg: '#eaf7ff', text: '#0d3a66', border: '#cfeafe' },
    aviso: { bg: '#fff4d6', text: '#8a5a00', border: '#f3d98a' },
    urgente: { bg: '#fde8e6', text: '#a3231a', border: '#f3b3ac' }
  };

  function injectStyle() {
    if (document.getElementById(STYLE_ID)) return;
    var style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent =
      '.announcement-banner{width:100%;box-sizing:border-box;padding:12px 20px;' +
      'text-align:center;font-family:"Manrope",system-ui,sans-serif;font-weight:700;' +
      'font-size:14px;line-height:1.4;}' +
      '.announcement-banner a{color:inherit;text-decoration:underline;font-weight:800;' +
      'margin-left:6px;white-space:nowrap;}';
    document.head.appendChild(style);
  }

  function render(data) {
    injectStyle();
    var colors = TYPE_COLORS[data.type] || TYPE_COLORS.info;
    var banner = document.createElement('div');
    banner.className = 'announcement-banner';
    banner.style.background = colors.bg;
    banner.style.color = colors.text;
    banner.style.borderBottom = '1px solid ' + colors.border;

    var text = document.createElement('span');
    text.textContent = data.message;
    banner.appendChild(text);

    if (data.link) {
      var link = document.createElement('a');
      link.href = data.link;
      link.textContent = 'Más información →';
      banner.appendChild(link);
    }

    document.body.insertBefore(banner, document.body.firstChild);
  }

  fetch(WORKER_URL + '?announcement=1')
    .then(function (r) { return r.json(); })
    .then(function (data) {
      if (data && data.active && data.message) render(data);
    })
    .catch(function () { /* a missing/failed banner is not worth surfacing as an error */ });
})();
