(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const fine = matchMedia('(hover:hover) and (pointer:fine)').matches;
  const G = window.gsap;
  const ST = window.ScrollTrigger;
  const hasMotion = !!(G && ST) && !reduce;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const clamp = (v, a = 0, b = 1) => Math.min(b, Math.max(a, v));
  const isMobile = () => innerWidth <= 760;

  /* ── Hero words: same width for all three, reached with the wdth axis ── */
  const flag = $('.flag');
  const bands = $$('[data-band]');
  const words = $$('.word');
  const letters = [];

  words.forEach(w => {
    const text = w.textContent;
    w.textContent = '';
    [...text].forEach(c => {
      const s = document.createElement('span');
      s.className = 'l';
      s.textContent = c;
      w.append(s);
      letters.push({ el: s, word: w, cur: 0, target: 0, x: 0, y: 0 });
    });
  });

  function portraitBox(W, H) {
    const aspect = 928 / 1063;
    if (isMobile()) {
      const blueH = bands[2].offsetHeight;
      const h = Math.min((W * 0.86) / aspect, blueH * 0.66);
      return { x: W - h * aspect, y: H - h, w: h * aspect, h };
    }
    const h = Math.min(H, (W * 0.52) / aspect);
    return { x: W - h * aspect, y: H - h, w: h * aspect, h };
  }

  function fitWords() {
    const W = flag.clientWidth;
    const H = flag.clientHeight;
    const box = portraitBox(W, H);
    const gut = parseFloat(getComputedStyle(bands[0]).paddingLeft);
    const target = isMobile() ? W - gut * 2 : box.x - gut - W * 0.012;
    const measure = (w, v) => { w.style.setProperty('--w', v); return w.scrollWidth; };

    let size = Math.min(...bands.map(b => b.offsetHeight)) * (isMobile() ? 0.86 : 0.8);
    words.forEach(w => { w.style.fontSize = size + 'px'; });
    const worst = Math.max(...words.map(w => measure(w, 50)));
    if (worst > target) {
      size *= target / worst;
      words.forEach(w => { w.style.fontSize = size + 'px'; });
    }
    words.forEach(w => {
      let lo = 50, hi = 150;
      for (let i = 0; i < 12; i++) {
        const mid = (lo + hi) / 2;
        if (measure(w, mid) > target) hi = mid; else lo = mid;
      }
      w.style.setProperty('--w', lo.toFixed(1));
      w.dataset.base = lo;
    });
    const fr = flag.getBoundingClientRect();
    letters.forEach(l => {
      const r = l.el.getBoundingClientRect();
      l.x = r.left - fr.left + r.width / 2;
      l.y = r.top - fr.top + r.height / 2;
    });
  }

  /* ── Packets: the portrait assembled on Manhattan routes ── */
  const canvas = $('#packets');
  const ctx = canvas.getContext('2d');
  const img = new Image();
  img.src = 'seanpp.jpeg';

  let P = null, W = 0, H = 0, dpr = 1, baked = null;
  let t0 = -1e9, introEnd = 0, scrollP = 0, loopOn = false;
  let px = -1e4, py = -1e4, lastMove = -1e4, energy = 0;

  function build() {
    W = flag.clientWidth;
    H = flag.clientHeight;
    dpr = Math.min(devicePixelRatio || 1, 1.75);
    canvas.width = W * dpr;
    canvas.height = H * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    const box = portraitBox(W, H);
    const step = Math.max(isMobile() ? 3.5 : 4.5, box.w / 150);
    const cols = Math.round(box.w / step), rows = Math.round(box.h / step);
    const off = document.createElement('canvas');
    off.width = cols; off.height = rows;
    const oc = off.getContext('2d', { willReadFrequently: true });
    oc.drawImage(img, 0, 0, cols, rows);
    const data = oc.getImageData(0, 0, cols, rows).data;

    const list = [];
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const i = (r * cols + c) * 4;
      const lum = data[i] * 0.3 + data[i + 1] * 0.59 + data[i + 2] * 0.11;
      let d = 1 - clamp((lum - 12) / 200);
      d = Math.pow(d, 1.5);
      if (d < 0.05) continue;
      list.push([box.x + c * step, box.y + r * step, step * (0.16 + 0.92 * d)]);
    }

    const n = list.length;
    P = { n, tx: new Float32Array(n), ty: new Float32Array(n), s: new Float32Array(n), sx: new Float32Array(n), sy: new Float32Array(n), de: new Float32Array(n), du: new Float32Array(n), ox: new Float32Array(n), oy: new Float32Array(n), vx: new Float32Array(n), vy: new Float32Array(n), dr: new Float32Array(n) };
    const lanes = Array.from({ length: 44 }, () => Math.random() * H);
    let maxEnd = 0;
    list.forEach(([x, y, s], i) => {
      P.tx[i] = x; P.ty[i] = y; P.s[i] = s;
      P.sx[i] = Math.random() < 0.78 ? -30 - Math.random() * 260 : W + 30 + Math.random() * 200;
      P.sy[i] = lanes[(Math.random() * lanes.length) | 0];
      P.de[i] = Math.random() * 900 + (x / W) * 260;
      P.du[i] = 1100 + Math.random() * 800;
      P.dr[i] = (0.4 + Math.random()) * (Math.random() < 0.5 ? -1 : 1);
      maxEnd = Math.max(maxEnd, P.de[i] + P.du[i]);
    });
    P.maxEnd = maxEnd;
    baked = null;
  }

  const ease = t => (t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2);

  function drawParticles(now) {
    const intro = now - t0 < P.maxEnd;
    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = '#0B1220';
    ctx.beginPath();
    const fade = 1 - clamp(scrollP * 1.15);
    const near = now - lastMove < 260;
    const R = 120, R2 = R * R;
    energy = 0;
    for (let i = 0; i < P.n; i++) {
      let x = P.tx[i], y = P.ty[i], s = P.s[i];
      if (intro) {
        const t = clamp((now - t0 - P.de[i]) / P.du[i]);
        const a = ease(clamp(t * 2)), b = ease(clamp(t * 2 - 1));
        x = P.sx[i] + (x - P.sx[i]) * a;
        y = P.sy[i] + (y - P.sy[i]) * b;
        s = s * (0.45 + 0.55 * b);
      }
      if (P.ox[i] || P.oy[i] || near) {
        if (near) {
          const dx = x + P.ox[i] - px, dy = y + P.oy[i] - py, d2 = dx * dx + dy * dy;
          if (d2 < R2) {
            const d = Math.sqrt(d2) || 1, f = (1 - d / R) * 3.4;
            P.vx[i] += (dx / d) * f; P.vy[i] += (dy / d) * f;
          }
        }
        P.vx[i] = (P.vx[i] - P.ox[i] * 0.07) * 0.84;
        P.vy[i] = (P.vy[i] - P.oy[i] * 0.07) * 0.84;
        P.ox[i] += P.vx[i]; P.oy[i] += P.vy[i];
        const e = Math.abs(P.ox[i]) + Math.abs(P.oy[i]) + Math.abs(P.vx[i]) + Math.abs(P.vy[i]);
        if (e < 0.08) { P.ox[i] = P.oy[i] = P.vx[i] = P.vy[i] = 0; } else energy += e;
        x += P.ox[i]; y += P.oy[i];
      }
      if (scrollP > 0) {
        x += P.dr[i] * scrollP * scrollP * W * 0.55;
        y += (P.dr[i] * 0.3) * scrollP * H;
        s *= fade;
      }
      if (s > 0.15) ctx.rect(x, y, s, s);
    }
    ctx.fill();
    return intro;
  }

  /* Letters widen under the pointer */
  function updateLetters() {
    let moving = false;
    letters.forEach(l => {
      const near = Math.abs(l.y - py) < 150 && performance.now() - lastMove < 1200;
      const d = near ? clamp(1 - Math.abs(l.x - px) / 190) : 0;
      l.target = d * d * 34;
      const nxt = l.cur + (l.target - l.cur) * 0.16;
      if (Math.abs(nxt - l.cur) > 0.04) {
        moving = true;
        l.cur = nxt;
        l.el.style.setProperty('--lw', (+l.word.dataset.base + l.cur).toFixed(1));
      } else if (l.cur !== l.target && l.target === 0 && l.cur !== 0) {
        l.cur = 0;
        l.el.style.removeProperty('--lw');
      }
    });
    return moving;
  }

  function tick(now) {
    if (!P) { loopOn = false; return; }
    const intro = drawParticles(now);
    const letting = fine && !reduce ? updateLetters() : false;
    if (intro || energy > 0.5 || letting || now - lastMove < 300) {
      requestAnimationFrame(tick);
    } else {
      loopOn = false;
    }
  }
  function wake() {
    if (!loopOn && P) { loopOn = true; requestAnimationFrame(tick); }
  }

  if (fine && !reduce) {
    $('.hero').addEventListener('pointermove', e => {
      const r = flag.getBoundingClientRect();
      px = e.clientX - r.left; py = e.clientY - r.top;
      lastMove = performance.now();
      wake();
    }, { passive: true });
  }

  /* ── Start ── */
  let started = false;
  async function start() {
    if (started) return;
    started = true;
    try { await Promise.all([document.fonts.load('900 100px Anybody'), img.decode()]); } catch (e) { /* draw what loaded */ }
    fitWords();
    build();

    if (!hasMotion) {
      t0 = -1e9;
      drawParticles(performance.now());
      return;
    }

    G.registerPlugin(ST);
    const tl = G.timeline({ defaults: { ease: 'expo.out' } });
    tl.fromTo('.band .bg', { scaleX: 0 }, { scaleX: 1, duration: 1, ease: 'expo.inOut', stagger: 0.11 }, 0)
      .fromTo('.word .l', { yPercent: 112 }, { yPercent: 0, duration: 1, stagger: { each: 0.018, from: 'start' } }, 0.55)
      .fromTo('.tag', { autoAlpha: 0, y: 10 }, { autoAlpha: 1, y: 0, duration: 0.7, stagger: 0.08 }, 0.95)
      .add(() => { t0 = performance.now(); wake(); }, 0.75);

    initScroll();
  }

  /* ── Scroll choreography ── */
  function initScroll() {
    if (fine && window.Lenis) {
      const lenis = new Lenis({ lerp: 0.11 });
      lenis.on('scroll', ST.update);
      G.ticker.add(t => lenis.raf(t * 1000));
      G.ticker.lagSmoothing(0);
      document.addEventListener('click', e => {
        const a = e.target.closest('a[href^="#"]');
        if (!a) return;
        const t = a.getAttribute('href') === '#haut' ? 0 : $(a.getAttribute('href'));
        if (t === null) return;
        e.preventDefault();
        lenis.scrollTo(t, { duration: 1.2 });
      });
    }

    const mm = G.matchMedia();

    mm.add('(min-width:761px)', () => {
      G.timeline({
        scrollTrigger: {
          trigger: '.hero', start: 'top top', end: '+=120%', pin: true, scrub: true, anticipatePin: 1,
          onToggle: s => flag.classList.toggle('scrub', s.isActive),
          onUpdate: s => { scrollP = clamp(s.progress / 0.72); wake(); }
        },
        defaults: { ease: 'none' }
      })
        .to('.band.v', { '--k': 0, duration: 1 }, 0)
        .to('.band.b', { '--k': 0, duration: 1 }, 0)
        .to('.band .word, .band .tag', { opacity: 0, duration: 0.3 }, 0);

      $$('.case').forEach((c, i, all) => {
        const inner = $('.case-inner', c);
        const shot = $('.case-shot img', c);
        G.fromTo(shot, { yPercent: 7 }, { yPercent: -3, ease: 'none', scrollTrigger: { trigger: c, start: 'top bottom', end: 'top top', scrub: true } });
        G.fromTo($('.case-name', c), { xPercent: -3 }, { xPercent: 0, ease: 'none', scrollTrigger: { trigger: c, start: 'top bottom', end: 'top top', scrub: true } });
        if (i < all.length - 1) {
          G.to(inner, { scale: 0.94, opacity: 0.35, ease: 'none', scrollTrigger: { trigger: all[i + 1], start: 'top bottom', end: 'top top', scrub: true } });
        }
      });
    });

    /* Words ink in as you read */
    $$('[data-words]').forEach(el => {
      const parts = el.textContent.split(/(\s+)/);
      el.setAttribute('aria-label', el.textContent);
      el.textContent = '';
      parts.forEach(p => {
        if (/^\s+$/.test(p)) { el.append(' '); return; }
        const s = document.createElement('span');
        s.className = 'w'; s.setAttribute('aria-hidden', 'true'); s.textContent = p;
        el.append(s);
      });
      G.fromTo($$('.w', el), { opacity: 0.14 }, { opacity: 1, stagger: 0.1, ease: 'none', scrollTrigger: { trigger: el, start: 'top 78%', end: 'bottom 62%', scrub: 0.4 } });
    });

    /* One reveal for everything else */
    G.utils.toArray('.trade, .steps li, .timeline li, .rows li').forEach(el => {
      G.from(el, { y: 26, autoAlpha: 0, duration: 0.8, ease: 'expo.out', scrollTrigger: { trigger: el, start: 'top 92%', once: true } });
    });
    G.utils.toArray('.more-title, .method h2, .path h2, .contact-title').forEach(el => {
      G.from(el, { yPercent: 30, autoAlpha: 0, duration: 1, ease: 'expo.out', scrollTrigger: { trigger: el, start: 'top 90%', once: true } });
    });

    if (fine) initPeek();
    document.fonts.ready.then(() => ST.refresh());
  }

  /* ── Hover preview that scrolls through the site ── */
  function initPeek() {
    const peek = $('#peek');
    const pImg = $('img', peek);
    const qx = G.quickTo(peek, 'x', { duration: 0.45, ease: 'power3' });
    const qy = G.quickTo(peek, 'y', { duration: 0.45, ease: 'power3' });
    let pan = null;
    const list = $('#rows');
    addEventListener('pointermove', e => {
      const w = peek.offsetWidth, h = peek.offsetHeight;
      qx(Math.min(e.clientX + 24, innerWidth - w - 16));
      qy(clamp(e.clientY - h / 2, 56, innerHeight - h - 16));
    }, { passive: true });
    $$('.row', list).forEach(row => {
      row.addEventListener('pointerenter', () => {
        pImg.src = row.dataset.peek;
        peek.classList.add('on');
        const run = () => {
          if (pan) pan.kill();
          G.set(pImg, { y: 0 });
          const dist = pImg.offsetHeight - peek.offsetHeight;
          if (dist > 8) pan = G.to(pImg, { y: -dist, duration: dist / 110, ease: 'none', delay: 0.35, repeat: -1, yoyo: true, repeatDelay: 0.6 });
        };
        if (pImg.complete) run(); else pImg.onload = run;
      });
      row.addEventListener('pointerleave', () => {
        peek.classList.remove('on');
        if (pan) pan.kill();
      });
    });
  }

  /* ── Resize ── */
  let lastW = innerWidth, rt;
  addEventListener('resize', () => {
    if (innerWidth === lastW) return;
    lastW = innerWidth;
    clearTimeout(rt);
    rt = setTimeout(() => {
      fitWords();
      build();
      t0 = -1e9;
      drawParticles(performance.now());
      if (ST && G) ST.refresh();
    }, 180);
  });

  /* ── Contact form: opens the visitor's mail app ── */
  const form = $('#form');
  form.addEventListener('submit', e => {
    e.preventDefault();
    const msg = $('#form-msg');
    const d = new FormData(form);
    const name = d.get('name').trim(), email = d.get('email').trim(), text = d.get('message').trim();
    if (!name || !email || !text) {
      msg.textContent = 'Indiquez votre nom, votre e-mail et votre projet.';
      return;
    }
    const subject = encodeURIComponent('Projet : message de ' + name);
    const body = encodeURIComponent(text + '\n\n' + name + '\n' + email);
    msg.textContent = 'Votre messagerie s\'ouvre avec le message prêt à envoyer.';
    location.href = 'mailto:nsoleypro20@gmail.com?subject=' + subject + '&body=' + body;
  });

  start();
})();
