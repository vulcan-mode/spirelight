// Shared across every page with the real nav (homepage, registro,
// metodo-pago, referral-form, privacidad, como-funciona, gracias): if
// this device already has a cached identity (spirelight_referrer), the
// "Regístrate" link makes no sense anymore -- swap it for "Mis datos"
// pointing at gracias/, prefilled with the cached phone so they never
// have to type it again just to look at or fix their own info.
(function () {
  // Two copies of this link can exist now -- one always-inline at
  // medium+ widths, one inside the hamburger drawer -- so every match
  // needs the swap, not just the first (getElementById would silently
  // miss the second one and leave it saying the wrong thing).
  var links = document.querySelectorAll('.nav-registro-link');
  if (!links.length) return;

  var identity;
  try {
    var raw = localStorage.getItem('spirelight_referrer');
    identity = raw ? JSON.parse(raw) : null;
  } catch (e) {
    identity = null;
  }

  if (identity && identity.phone) {
    // The homepage lives at the site root (one level up from every
    // other page this script runs on), so "gracias/" needs no "../"
    // there -- computed from actual path depth instead of assuming
    // every page is one level deep, which broke the moment this
    // script started running on index.html too.
    var depth = window.location.pathname.replace(/^\/|\/$/g, '').split('/').filter(Boolean).length;
    var prefix = depth === 0 ? '' : '../';
    links.forEach(function (link) {
      link.textContent = 'Mis datos';
      link.href = prefix + 'gracias/?phone=' + encodeURIComponent(identity.phone);
    });
  }
})();
