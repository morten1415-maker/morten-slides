/* =====================================================
   Morten DS — Slide-framework runtime
   Tastaturnavigation + fit-to-viewport-skalering.
   Ingen dependencies. Inkludér efter slides i HTML.
   ===================================================== */
(function () {
  const stage = document.querySelector('.deck__stage');
  const slides = Array.from(document.querySelectorAll('.slide'));
  if (!stage || slides.length === 0) return;

  let index = 0;

  // HUD (kun progress-bar — intet side-nummer)
  const hud = document.createElement('div');
  hud.className = 'deck__hud';
  hud.innerHTML = '<span class="bar"><i></i></span>';
  document.body.appendChild(hud);
  const barEl = hud.querySelector('.bar > i');

  function pad(n) { return String(n).padStart(2, '0'); }

  function render() {
    slides.forEach((s, i) => s.classList.toggle('is-active', i === index));
    barEl.style.width = ((index + 1) / slides.length * 100) + '%';
    location.hash = '#' + (index + 1);
  }

  function go(i) {
    index = Math.max(0, Math.min(slides.length - 1, i));
    render();
  }
  function next() { go(index + 1); }
  function prev() { go(index - 1); }

  // Skaler canvas så 1280×720 passer i viewport (contain)
  function fit() {
    const margin = 48;
    const sw = (window.innerWidth - margin) / stage.offsetWidth;
    const sh = (window.innerHeight - margin) / stage.offsetHeight;
    stage.style.transform = 'scale(' + Math.min(sw, sh) + ')';
  }

  window.addEventListener('resize', fit);

  document.addEventListener('keydown', (e) => {
    switch (e.key) {
      case 'ArrowRight':
      case 'ArrowDown':
      case 'PageDown':
      case ' ':
        e.preventDefault(); next(); break;
      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
        e.preventDefault(); prev(); break;
      case 'Home': e.preventDefault(); go(0); break;
      case 'End': e.preventDefault(); go(slides.length - 1); break;
    }
  });

  // Klik på højre/venstre halvdel = frem/tilbage
  stage.addEventListener('click', (e) => {
    const mid = window.innerWidth / 2;
    (e.clientX < mid ? prev : next)();
  });

  // Start på hash hvis sat (#3)
  const fromHash = parseInt((location.hash || '').slice(1), 10);
  if (fromHash >= 1 && fromHash <= slides.length) index = fromHash - 1;

  fit();
  render();
})();
