/**
 * D365 Demo Copilot — Visual Overlay Engine
 * 
 * Provides all visual demonstration capabilities:
 * - Spotlight: Dims page and highlights a target element with a glowing ring
 * - Captions: Movie-subtitle-style text at bottom of screen
 * - Business Value Cards: Callout cards with metrics and outcomes
 * - Progress: Step counter and progress bar
 * - Click Ripple: Visual feedback when agent clicks an element
 * - Pause Overlay: Full-screen pause indicator
 * - Title Slide: Opening/closing presentation slide
 * - Tooltips: Small annotations on elements
 * 
 * All methods are exposed on window.DemoCopilot for Playwright to call.
 */

(function () {
  'use strict';

  // Prevent double initialization
  if (window.DemoCopilot) {
    // But verify DOM still exists — D365 may have wiped it
    if (document.getElementById('demo-copilot-root')) return;
    // DOM is gone but JS object remains — clean up and re-init
    console.log('[DemoCopilot] DOM was removed, re-initializing...');
    delete window.DemoCopilot;
  }

  // ---- Utility ----
  const qs = (sel, ctx = document) => ctx.querySelector(sel);
  const ce = (tag, cls, html) => {
    const el = document.createElement(tag);
    if (cls) el.className = cls;
    if (html) el.innerHTML = html;
    return el;
  };

  /**
   * Ensure the overlay DOM scaffold exists.
   * Called before any operation to self-heal if D365 removed our elements.
   */
  function ensureDOM() {
    if (!document.getElementById('demo-copilot-root')) {
      console.log('[DemoCopilot] Self-healing: rebuilding DOM scaffold');
      buildDOM();
    }
  }

  // ---- State ----
  const state = {
    totalSteps: 0,
    currentStep: 0,
    isPaused: false,
    isActive: false,
    spotlightRect: null,       // bounding rect of current spotlight target
    captionPosition: 'auto',   // 'auto', 'top', or 'bottom'
  };

  // ---- DOM Scaffold ----
  function buildDOM() {
    // Remove existing if any
    const existing = qs('#demo-copilot-root');
    if (existing) existing.remove();

    const root = ce('div');
    root.id = 'demo-copilot-root';

    // 1. Spotlight overlay (SVG with cutout)
    root.innerHTML = `
      <!-- Spotlight -->
      <div class="demo-spotlight-overlay" id="demo-spotlight">
        <svg xmlns="http://www.w3.org/2000/svg">
          <defs>
            <mask id="demo-spotlight-mask">
              <rect width="100%" height="100%" fill="white"/>
              <rect id="demo-spotlight-hole" rx="8" ry="8" fill="black"
                    x="0" y="0" width="0" height="0"/>
            </mask>
          </defs>
          <rect width="100%" height="100%" fill="rgba(0,0,0,0.55)"
                mask="url(#demo-spotlight-mask)"/>
        </svg>
      </div>

      <!-- Spotlight ring -->
      <div class="demo-spotlight-ring" id="demo-spotlight-ring"
           style="display:none;"></div>

      <!-- Caption bar -->
      <div class="demo-caption-bar" id="demo-caption-bar">
        <div class="demo-caption-container">
          <span class="demo-caption-phase tell" id="demo-caption-phase">TELL</span>
          <p class="demo-caption-text" id="demo-caption-text"></p>
        </div>
      </div>

      <!-- Business value backdrop -->
      <div class="demo-value-backdrop" id="demo-value-backdrop"></div>

      <!-- Business value card -->
      <div class="demo-value-card" id="demo-value-card">
        <div class="demo-value-card-header">
          <div class="demo-value-card-icon">
            <svg viewBox="0 0 24 24"><path d="M12 2L1 21h22L12 2zm0 4l7.53 13H4.47L12 6z" fill="none" stroke="white" stroke-width="2"/><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>
          </div>
          <h4 class="demo-value-card-title" id="demo-value-title">Business Value</h4>
        </div>
        <div class="demo-value-card-body">
          <p class="demo-value-card-text" id="demo-value-text"></p>
          <div class="demo-value-card-metric" id="demo-value-metric" style="display:none;">
            <span class="demo-value-card-metric-value" id="demo-value-metric-val"></span>
            <span class="demo-value-card-metric-label" id="demo-value-metric-label"></span>
          </div>
        </div>
      </div>

      <!-- Progress bar -->
      <div class="demo-progress-bar" id="demo-progress-bar" style="display:none;">
        <div class="demo-progress-fill" id="demo-progress-fill"></div>
      </div>

      <!-- Step indicator -->
      <div class="demo-step-indicator" id="demo-step-indicator">
        <span class="demo-step-indicator-label">Step</span>
        <span class="demo-step-indicator-current" id="demo-step-current">1 / 1</span>
        <div class="demo-step-dots" id="demo-step-dots"></div>
      </div>

      <!-- Agent status pill -->
      <div class="demo-status-pill" id="demo-status-pill">
        <div class="demo-status-dot" id="demo-status-dot"></div>
        <span class="demo-status-label" id="demo-status-label">Ready</span>
      </div>

      <!-- Pause overlay -->
      <div class="demo-pause-overlay" id="demo-pause-overlay">
        <div class="demo-pause-icon">
          <svg viewBox="0 0 24 24"><rect x="6" y="4" width="4" height="16" rx="1"/><rect x="14" y="4" width="4" height="16" rx="1"/></svg>
          <div class="demo-pause-text">Demo Paused</div>
          <div class="demo-pause-hint">Say "continue" or press Ctrl+Space to resume</div>
        </div>
      </div>

      <!-- Title slide -->
      <div class="demo-title-slide" id="demo-title-slide">
        <svg class="demo-title-logo" viewBox="0 0 24 24" fill="white">
          <path d="M3 3v18h18V3H3zm16 16H5V5h14v14zM7 7h4v4H7V7zm0 6h4v4H7v-4zm6-6h4v4h-4V7zm0 6h4v4h-4v-4z"/>
        </svg>
        <h1 class="demo-title-heading" id="demo-title-heading"></h1>
        <p class="demo-title-subheading" id="demo-title-subheading"></p>
        <p class="demo-title-meta" id="demo-title-meta"></p>
      </div>
    `;

    document.body.appendChild(root);
  }

  /**
   * Determine best caption position based on where the spotlight target is.
   * If target is in the bottom half of the viewport, show caption at top.
   * If target is in the top half, show caption at bottom.
   */
  function updateCaptionPosition() {
    const bar = qs('#demo-caption-bar');
    if (!bar) return;

    let pos = state.captionPosition;
    if (pos === 'auto' && state.spotlightRect) {
      const viewH = window.innerHeight;
      const targetMid = state.spotlightRect.top + state.spotlightRect.height / 2;
      pos = (targetMid > viewH * 0.5) ? 'top' : 'bottom';
    } else if (pos === 'auto') {
      pos = 'bottom';
    }

    bar.classList.remove('position-top', 'position-bottom');
    bar.classList.add(`position-${pos}`);
  }

  // ---- Spotlight ----
  async function spotlightOn(selector, padding = 12) {
    ensureDOM();
    let el = typeof selector === 'string' ? qs(selector) : selector;
    if (!el) {
      // Retry once after D365 async render delay
      await sleep(500);
      el = typeof selector === 'string' ? qs(selector) : selector;
      if (!el) {
        console.warn('[DemoCopilot] Spotlight target not found:', selector);
        spotlightOff();  // Clear any lingering dim
        return;
      }
    }

    const rect = el.getBoundingClientRect();
    state.spotlightRect = rect;  // Track for caption positioning
    const hole = qs('#demo-spotlight-hole');
    const ring = qs('#demo-spotlight-ring');
    const overlay = qs('#demo-spotlight');

    // Position the cutout hole
    hole.setAttribute('x', rect.left - padding);
    hole.setAttribute('y', rect.top - padding);
    hole.setAttribute('width', rect.width + padding * 2);
    hole.setAttribute('height', rect.height + padding * 2);

    // Position the glowing ring
    ring.style.display = 'block';
    ring.style.left = `${rect.left - padding}px`;
    ring.style.top = `${rect.top - padding}px`;
    ring.style.width = `${rect.width + padding * 2}px`;
    ring.style.height = `${rect.height + padding * 2}px`;

    overlay.classList.add('active');

    // Scroll element into view if needed
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function spotlightOff() {
    const overlay = qs('#demo-spotlight');
    const ring = qs('#demo-spotlight-ring');
    const hole = qs('#demo-spotlight-hole');

    if (overlay) overlay.classList.remove('active');
    if (ring) ring.style.display = 'none';
    if (hole) {
      hole.setAttribute('x', 0);
      hole.setAttribute('y', 0);
      hole.setAttribute('width', 0);
      hole.setAttribute('height', 0);
    }
    state.spotlightRect = null;
  }

  // ---- Captions ----
  function showCaption(text, phase = 'tell', position = 'auto') {
    ensureDOM();
    const bar = qs('#demo-caption-bar');
    const phaseEl = qs('#demo-caption-phase');
    const textEl = qs('#demo-caption-text');

    if (!bar || !phaseEl || !textEl) {
      console.warn('[DemoCopilot] Caption DOM elements not found after ensureDOM');
      return;
    }

    // Set phase badge
    phaseEl.className = `demo-caption-phase ${phase}`;
    const phaseLabels = { tell: 'TELL', show: 'SHOW', value: 'BUSINESS VALUE' };
    phaseEl.textContent = phaseLabels[phase] || phase.toUpperCase();

    // Set text (supports <span class="highlight"> for emphasis)
    textEl.innerHTML = text;

    // Position the caption bar (top or bottom) based on spotlight location
    state.captionPosition = position;
    updateCaptionPosition();

    bar.classList.add('visible');
  }

  function hideCaption() {
    const bar = qs('#demo-caption-bar');
    if (bar) bar.classList.remove('visible');
  }

  /**
   * Typewriter-style caption reveal
   */
  async function showCaptionAnimated(text, phase = 'tell', speed = 30, position = 'auto') {
    ensureDOM();
    const bar = qs('#demo-caption-bar');
    const phaseEl = qs('#demo-caption-phase');
    const textEl = qs('#demo-caption-text');

    if (!bar || !phaseEl || !textEl) {
      console.warn('[DemoCopilot] Caption DOM elements not found for animation');
      return;
    }

    phaseEl.className = `demo-caption-phase ${phase}`;
    const phaseLabels = { tell: 'TELL', show: 'SHOW', value: 'BUSINESS VALUE' };
    phaseEl.textContent = phaseLabels[phase] || phase.toUpperCase();

    // Position the caption bar
    state.captionPosition = position;
    updateCaptionPosition();

    textEl.innerHTML = '';
    bar.classList.add('visible');

    // Strip HTML for character-by-character reveal, then set full HTML at end
    const plainText = text.replace(/<[^>]+>/g, '');
    for (let i = 0; i < plainText.length; i++) {
      if (state.isPaused) {
        await waitForResume();
      }
      textEl.textContent = plainText.slice(0, i + 1);
      await sleep(speed);
    }
    // Set full HTML (with formatting) at the end
    textEl.innerHTML = text;
  }

  // ---- Business Value Card ----
  function showValueCard(title, text, metric = null, position = 'center') {
    ensureDOM();
    const card = qs('#demo-value-card');
    const backdrop = qs('#demo-value-backdrop');
    if (!card) return;
    const titleEl = qs('#demo-value-title');
    const textEl = qs('#demo-value-text');
    if (!titleEl || !textEl) return;

    titleEl.textContent = title;
    textEl.innerHTML = text;

    const metricEl = qs('#demo-value-metric');
    if (metric) {
      const valEl = qs('#demo-value-metric-val');
      qs('#demo-value-metric-label').textContent = metric.label;
      metricEl.style.display = 'flex';
      // Animate the metric value counting up
      _animateMetric(valEl, metric.value);
    } else {
      metricEl.style.display = 'none';
    }

    // Card is always centered via CSS
    card.style.top = '50%';
    card.style.left = '50%';

    if (backdrop) backdrop.classList.add('visible');
    card.classList.add('visible');
  }

  function hideValueCard() {
    const card = qs('#demo-value-card');
    const backdrop = qs('#demo-value-backdrop');
    if (card) card.classList.remove('visible');
    if (backdrop) backdrop.classList.remove('visible');
  }

  /** Animate a metric value with a counting-up effect. */
  function _animateMetric(el, target) {
    // Extract leading number and suffix (e.g. '40%' → 40, '%')
    const match = target.match(/^([\$€£]?)([\d,.]+)(.*)$/);
    if (!match) { el.textContent = target; return; }
    const prefix = match[1];
    const numStr = match[2].replace(/,/g, '');
    const suffix = match[3];
    const num = parseFloat(numStr);
    if (isNaN(num)) { el.textContent = target; return; }

    const isFloat = numStr.includes('.');
    const decimals = isFloat ? (numStr.split('.')[1] || '').length : 0;
    const duration = 1200;
    const start = performance.now();

    function tick(now) {
      const elapsed = now - start;
      const progress = Math.min(elapsed / duration, 1);
      // Ease-out cubic
      const eased = 1 - Math.pow(1 - progress, 3);
      const current = num * eased;
      el.textContent = prefix + (isFloat ? current.toFixed(decimals) : Math.round(current).toLocaleString()) + suffix;
      if (progress < 1) requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  // ---- Progress ----
  function initProgress(totalSteps) {
    state.totalSteps = totalSteps;
    state.currentStep = 0;

    const bar = qs('#demo-progress-bar');
    const indicator = qs('#demo-step-indicator');
    const dots = qs('#demo-step-dots');

    bar.style.display = 'block';
    qs('#demo-progress-fill').style.width = '0%';

    // Build dot indicators
    dots.innerHTML = '';
    for (let i = 0; i < totalSteps; i++) {
      const dot = ce('div', 'demo-step-dot');
      dot.dataset.step = i;
      dots.appendChild(dot);
    }

    indicator.classList.add('visible');
    updateProgress(0);
  }

  function updateProgress(stepIndex) {
    state.currentStep = stepIndex;
    const pct = ((stepIndex + 1) / state.totalSteps) * 100;

    qs('#demo-progress-fill').style.width = `${pct}%`;
    qs('#demo-step-current').textContent = `${stepIndex + 1} / ${state.totalSteps}`;

    // Update dots
    const dots = document.querySelectorAll('.demo-step-dot');
    dots.forEach((dot, i) => {
      dot.classList.remove('completed', 'active');
      if (i < stepIndex) dot.classList.add('completed');
      if (i === stepIndex) dot.classList.add('active');
    });
  }

  function hideProgress() {
    const bar = qs('#demo-progress-bar');
    const indicator = qs('#demo-step-indicator');
    if (bar) bar.style.display = 'none';
    if (indicator) indicator.classList.remove('visible');
  }

  // ---- Click Ripple ----
  function clickRipple(x, y) {
    const ripple = ce('div', 'demo-click-ripple');
    ripple.style.left = `${x}px`;
    ripple.style.top = `${y}px`;
    document.body.appendChild(ripple);
    setTimeout(() => ripple.remove(), 700);
  }

  function clickRippleOnElement(selector) {
    const el = typeof selector === 'string' ? qs(selector) : selector;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    clickRipple(rect.left + rect.width / 2, rect.top + rect.height / 2);
  }

  // ---- Tooltip ----
  function showTooltip(selector, text, position = 'above') {
    hideTooltip(); // Remove any existing

    const el = typeof selector === 'string' ? qs(selector) : selector;
    if (!el) return;

    const rect = el.getBoundingClientRect();
    const tooltip = ce('div', 'demo-tooltip', text);
    tooltip.id = 'demo-active-tooltip';
    document.body.appendChild(tooltip);

    // Position above the element
    const tipRect = tooltip.getBoundingClientRect();
    if (position === 'above') {
      tooltip.style.left = `${rect.left + rect.width / 2 - tipRect.width / 2}px`;
      tooltip.style.top = `${rect.top - tipRect.height - 12}px`;
    } else {
      tooltip.style.left = `${rect.left + rect.width / 2 - tipRect.width / 2}px`;
      tooltip.style.top = `${rect.bottom + 12}px`;
      tooltip.style.transform = 'translateY(0)';
      // Flip arrow
      tooltip.querySelector('::after') // Can't style pseudo in JS, handle via class
      tooltip.classList.add('below');
    }

    requestAnimationFrame(() => tooltip.classList.add('visible'));
  }

  function hideTooltip() {
    const tip = qs('#demo-active-tooltip');
    if (tip) tip.remove();
  }

  // ---- Pause / Resume ----
  function pause() {
    state.isPaused = true;
    const overlay = qs('#demo-pause-overlay');
    if (overlay) overlay.classList.add('active');
  }

  function resume() {
    state.isPaused = false;
    const overlay = qs('#demo-pause-overlay');
    if (overlay) overlay.classList.remove('active');
  }

  function isPaused() {
    return state.isPaused;
  }

  // ---- Title Slide ----
  function showTitleSlide(heading, subheading = '', meta = '') {
    ensureDOM();
    qs('#demo-title-heading').textContent = heading;
    qs('#demo-title-subheading').textContent = subheading;
    qs('#demo-title-meta').textContent = meta;
    qs('#demo-title-slide').classList.add('active');
  }

  function hideTitleSlide() {
    qs('#demo-title-slide').classList.remove('active');
  }

  // ---- Agent Status Indicator ----
  function showStatus(text, mode = 'working') {
    ensureDOM();
    const pill = qs('#demo-status-pill');
    const dot = qs('#demo-status-dot');
    const label = qs('#demo-status-label');
    if (!pill || !dot || !label) return;
    label.textContent = text;
    dot.className = 'demo-status-dot ' + mode;
    pill.classList.add('visible');
  }

  function hideStatus() {
    const pill = qs('#demo-status-pill');
    if (pill) pill.classList.remove('visible');
  }

  // ---- Lifecycle ----
  function init() {
    buildDOM();
    state.isActive = true;
  }

  function destroy() {
    const root = qs('#demo-copilot-root');
    if (root) root.remove();
    state.isActive = false;
  }

  function clearAll() {
    spotlightOff();
    hideCaption();
    hideValueCard();
    hideTooltip();
  }

  // ---- Helpers ----
  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function waitForResume() {
    return new Promise(resolve => {
      const check = setInterval(() => {
        if (!state.isPaused) {
          clearInterval(check);
          resolve();
        }
      }, 100);
    });
  }

  // ---- Keyboard Shortcuts ----
  document.addEventListener('keydown', (e) => {
    if (!state.isActive) return;

    // Ignore when user is typing in an input/textarea/contenteditable
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || e.target.isContentEditable) return;

    if (e.code === 'Space' && e.ctrlKey) {
      e.preventDefault();
      if (state.isPaused) {
        resume();
      } else if (typeof window.__demoCopilotAction === 'function') {
        // Advance to next step
        window.__demoCopilotAction('advance');
      }
    }
    if (e.code === 'Escape') {
      if (state.isPaused) {
        resume();
      } else {
        pause();
      }
    }
  });

  // ---- Public API ----
  window.DemoCopilot = {
    // Lifecycle
    init,
    destroy,
    clearAll,

    // Status
    showStatus,
    hideStatus,

    // Spotlight
    spotlightOn,
    spotlightOff,

    // Captions
    showCaption,
    showCaptionAnimated,
    hideCaption,

    // Business Value
    showValueCard,
    hideValueCard,

    // Progress
    initProgress,
    updateProgress,
    hideProgress,

    // Click Effect
    clickRipple,
    clickRippleOnElement,

    // Tooltips
    showTooltip,
    hideTooltip,

    // Pause / Resume
    pause,
    resume,
    isPaused,

    // Title Slide
    showTitleSlide,
    hideTitleSlide,

    // State
    getState: () => ({ ...state }),
  };

  // Auto-init
  init();

  console.log('%c🎬 D365 Demo Copilot Overlay loaded', 'color: #0078D4; font-weight: bold; font-size: 14px;');
})();
