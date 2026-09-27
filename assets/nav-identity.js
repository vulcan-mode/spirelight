// Shared across every page with the real nav (referral-form, privacidad,
// como-funciona, gracias): if this device already has a cached identity
// (spirelight_referrer), the "Regístrate" link makes no sense anymore --
// swap it for "Mis datos" pointing at gracias/, prefilled with
// the cached phone so they never have to type it again just to look at
// or fix their own info.
(function () {
  var link = document.getElementById('navRegistroLink');
  if (!link) return;

  var identity;
  try {
    var raw = localStorage.getItem('spirelight_referrer');
    identity = raw ? JSON.parse(raw) : null;
  } catch (e) {
    identity = null;
  }

  if (identity && identity.phone) {
    link.textContent = 'Mis datos';
    link.href = '../gracias/?phone=' + encodeURIComponent(identity.phone);
  }
})();
