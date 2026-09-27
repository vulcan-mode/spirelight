// Shared lightweight toast -- for transient status messages ("revisa tu
// correo", "guardamos tu método de pago") that shouldn't take over the
// whole screen and strand the visitor with no way back to what they
// were doing. Self-contained (injects its own styles) so any page can
// just <script src="../assets/toast.js"> and call showToast(text).
(function () {
  var STYLE_ID = 'spirelight-toast-styles';
  if (!document.getElementById(STYLE_ID)) {
    var style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = [
      '.spirelight-toast {',
      '  position: fixed;',
      '  left: 50%;',
      '  bottom: 28px;',
      '  transform: translate(-50%, 16px);',
      '  max-width: min(420px, calc(100vw - 32px));',
      '  background: #0d3a66;',
      '  color: #fff;',
      '  font-family: Manrope, system-ui, sans-serif;',
      '  font-size: 14px;',
      '  font-weight: 600;',
      '  line-height: 1.4;',
      '  padding: 14px 18px;',
      '  border-radius: 14px;',
      '  box-shadow: 0 12px 32px rgba(13, 58, 102, 0.35);',
      '  opacity: 0;',
      '  pointer-events: none;',
      '  transition: opacity 0.25s ease, transform 0.25s ease;',
      '  z-index: 1000;',
      '  text-align: center;',
      '}',
      '.spirelight-toast.visible {',
      '  opacity: 1;',
      '  transform: translate(-50%, 0);',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  window.showToast = function (message, durationMs) {
    var el = document.createElement('div');
    el.className = 'spirelight-toast';
    el.textContent = message;
    document.body.appendChild(el);
    // Force a reflow so the transition actually runs instead of
    // starting already in its end state.
    void el.offsetWidth;
    el.classList.add('visible');

    setTimeout(function () {
      el.classList.remove('visible');
      setTimeout(function () { el.remove(); }, 300);
    }, durationMs || 4000);
  };
})();
