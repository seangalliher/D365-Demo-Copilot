/**
 * D365 Demo Copilot — Sidecar Chat Panel (Shadow DOM Isolated)
 *
 * In-browser chat UI injected inside a Shadow DOM root so that
 * D365's global CSS cannot interfere with rendering.
 *
 * Architecture:
 *   - Host element (#demo-chat-panel-host) in light DOM
 *   - Shadow root contains all chat HTML + CSS
 *   - Toggle button in light DOM (visible when panel collapsed)
 *   - :host pseudo-class styles the host from within shadow CSS
 *
 * API exposed on window.DemoCopilotChat:
 *   addMessage(role, text)       — Add a chat message
 *   showTyping()                 — Show typing indicator
 *   hideTyping()                 — Hide typing indicator
 *   showPlan(plan)               — Render a plan summary card
 *   updatePlanStep(idx, status)  — Update step status in plan card
 *   showProgress(step, total, label)  — Show demo progress bar
 *   hideProgress()               — Hide progress bar
 *   setStatus(status)            — Set header status text
 *   disable()                    — Disable input
 *   enable()                     — Enable input
 *   showQuickActions(actions)    — Show quick action buttons
 *   hideQuickActions()           — Hide quick action buttons
 *   clear()                      — Clear all messages
 *   collapse() / expand()        — Toggle panel visibility
 *   showWelcome()                — Show the welcome screen
 *   hideWelcome()                — Hide the welcome screen
 *   getWidth()                   — Get the rendered panel width in px
 *
 * Browser -> Python bridge:
 *   window.__demoCopilotSend(text)  — exposed by Playwright
 *   window.__demoCopilotAction(action) — exposed by Playwright
 */

(function () {
  'use strict';

  // Prevent double init
  if (window.DemoCopilotChat) {
    if (document.getElementById('demo-chat-panel-host')) return;
    delete window.DemoCopilotChat;
  }

  // ---- Clean up stale elements ----
  var oldHost = document.getElementById('demo-chat-panel-host');
  if (oldHost) oldHost.remove();
  var oldToggle = document.getElementById('demo-chat-toggle');
  if (oldToggle) oldToggle.remove();
  // Also clean any non-shadow remnants from previous versions
  var oldPanel = document.getElementById('demo-chat-panel');
  if (oldPanel) oldPanel.remove();
  var oldStyles = document.getElementById('demo-chat-panel-styles');
  if (oldStyles) oldStyles.remove();

  // ---- Create Shadow DOM host ----
  var host = document.createElement('div');
  host.id = 'demo-chat-panel-host';
  document.body.appendChild(host);

  var shadow = host.attachShadow({ mode: 'open' });

  // ---- Inject CSS into shadow root ----
  var cssText = window.__demoCopilotCSS || '';
  var styleEl = document.createElement('style');
  styleEl.textContent = cssText;
  shadow.appendChild(styleEl);
  delete window.__demoCopilotCSS;

  // ---- Build panel DOM inside shadow root ----
  var panel = document.createElement('div');
  panel.id = 'demo-chat-panel';
  panel.innerHTML =
    '<!-- Header -->' +
    '<div class="chat-header">' +
      '<div class="chat-header-icon">\uD83E\uDD16</div>' +
      '<div class="chat-header-title">Demo Copilot</div>' +
      '<span class="chat-header-status" id="chat-status">Ready</span>' +
      '<button class="chat-header-voice-toggle" id="chat-voice-toggle" title="Toggle voice narration">\uD83D\uDD07</button>' +
      '<button class="chat-header-collapse" id="chat-collapse" title="Collapse">\u25C0</button>' +
    '</div>' +

    '<!-- Step Tracker (pinned between header and messages) -->' +
    '<div class="chat-step-tracker" id="chat-step-tracker" style="display:none;">' +
      '<div class="chat-tracker-header">' +
        '<span class="chat-tracker-title">Demo Steps</span>' +
        '<button class="chat-tracker-toggle" id="chat-tracker-toggle" title="Toggle step list">\u25BC</button>' +
      '</div>' +
      '<div class="chat-tracker-body" id="chat-tracker-body"></div>' +
    '</div>' +

    '<!-- Welcome -->' +
    '<div class="chat-welcome" id="chat-welcome">' +
      '<div class="chat-welcome-icon">\uD83C\uDFAC</div>' +
      '<h2>D365 Demo Copilot</h2>' +
      '<p>Describe what you\'d like to demonstrate in Dynamics 365 Project Operations, and I\'ll create a live interactive walkthrough.</p>' +
      '<div class="chat-welcome-hints" id="chat-welcome-hints">' +
        '<div class="chat-welcome-hint" data-hint="Show me how to create a time entry and submit it for approval">Create &amp; submit a time entry</div>' +
        '<div class="chat-welcome-hint" data-hint="Walk through creating a new project with tasks and team members">Create a project with WBS</div>' +
        '<div class="chat-welcome-hint" data-hint="Demonstrate the expense entry process including receipt attachment">Submit an expense report</div>' +
      '</div>' +
    '</div>' +

    '<!-- Messages -->' +
    '<div class="chat-messages" id="chat-messages" style="display:none;"></div>' +

    '<!-- Quick actions (during demo) -->' +
    '<div class="chat-quick-actions" id="chat-quick-actions" style="display:none;">' +
      '<button class="chat-quick-btn" data-action="pause">\u23F8 Pause</button>' +
      '<button class="chat-quick-btn" data-action="resume">\u25B6 Resume</button>' +
      '<button class="chat-quick-btn" data-action="skip">\u23ED Skip</button>' +
      '<button class="chat-quick-btn danger" data-action="quit">\u23F9 Stop</button>' +
    '</div>' +

    '<!-- Input -->' +
    '<div class="chat-input-area">' +
      '<div class="chat-input-label">Type your demo request</div>' +
      '<div class="chat-input-wrap pulse" id="chat-input-wrap">' +
        '<textarea class="chat-input-field" id="chat-input" placeholder="e.g. Show me how to create a time entry..." rows="1"></textarea>' +
        '<button class="chat-input-send" id="chat-send" title="Send">\u27A4</button>' +
      '</div>' +
    '</div>';

  shadow.appendChild(panel);

  // ---- Toggle button (in LIGHT DOM for visibility when collapsed) ----
  var toggle = document.createElement('button');
  toggle.id = 'demo-chat-toggle';
  toggle.textContent = '\uD83D\uDCAC';
  toggle.style.cssText = [
    'position:fixed',
    'top:12px',
    'right:12px',
    'z-index:99999',
    'width:44px',
    'height:44px',
    'border-radius:50%',
    'background:#0078D4',
    'border:none',
    'cursor:pointer',
    'display:none',
    'align-items:center',
    'justify-content:center',
    'box-shadow:0 2px 12px rgba(0,0,0,0.3)',
    'pointer-events:auto',
    'color:white',
    'font-size:20px',
    'transition:all 0.2s ease',
    'font-family:inherit',
    'line-height:1'
  ].join(' !important;') + ' !important;';
  document.body.appendChild(toggle);

  // ---- State ----
  var chatState = {
    disabled: false,
    collapsed: false,
    voiceEnabled: false,
    planCardEl: null,
    progressEl: null,
    typingEl: null,
  };

  // ---- Element refs (all inside shadow root) ----
  var els = {
    host: host,
    panel: panel,
    messages: shadow.getElementById('chat-messages'),
    welcome: shadow.getElementById('chat-welcome'),
    input: shadow.getElementById('chat-input'),
    sendBtn: shadow.getElementById('chat-send'),
    status: shadow.getElementById('chat-status'),
    collapseBtn: shadow.getElementById('chat-collapse'),
    toggle: toggle, // in light DOM
    voiceToggle: shadow.getElementById('chat-voice-toggle'),
    quickActions: shadow.getElementById('chat-quick-actions'),
    welcomeHints: shadow.getElementById('chat-welcome-hints'),
    stepTracker: shadow.getElementById('chat-step-tracker'),
    trackerBody: shadow.getElementById('chat-tracker-body'),
    trackerToggle: shadow.getElementById('chat-tracker-toggle'),
  };

  // ---- Step tracker collapse/expand ----
  var trackerCollapsed = false;
  els.trackerToggle.addEventListener('click', function () {
    trackerCollapsed = !trackerCollapsed;
    els.trackerBody.style.display = trackerCollapsed ? 'none' : 'block';
    els.trackerToggle.textContent = trackerCollapsed ? '\u25B6' : '\u25BC';
    els.trackerToggle.title = trackerCollapsed ? 'Show step list' : 'Hide step list';
  });

  // ---- Auto-resize textarea ----
  els.input.addEventListener('input', function () {
    els.input.style.height = 'auto';
    els.input.style.height = Math.min(els.input.scrollHeight, 120) + 'px';
  });

  // ---- Send message ----
  function sendMessage() {
    if (chatState.disabled) return;
    var text = els.input.value.trim();
    if (!text) return;

    addMessage('user', text);
    els.input.value = '';
    els.input.style.height = 'auto';
    hideWelcome();

    if (window.__demoCopilotSend) {
      window.__demoCopilotSend(text);
    } else {
      console.warn('[DemoCopilotChat] __demoCopilotSend not exposed yet');
      addMessage('system', 'Chat backend not connected. Please wait...');
    }
  }

  els.sendBtn.addEventListener('click', sendMessage);
  els.input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  // ---- Collapse / Expand ----
  function collapse() {
    chatState.collapsed = true;
    host.classList.add('collapsed');
    toggle.style.display = 'flex';
    if (window.__demoCopilotAction) {
      window.__demoCopilotAction('panel_collapsed');
    }
  }

  function expand() {
    chatState.collapsed = false;
    host.classList.remove('collapsed');
    toggle.style.display = 'none';
    if (window.__demoCopilotAction) {
      window.__demoCopilotAction('panel_expanded');
    }
    els.input.focus();
  }

  els.collapseBtn.addEventListener('click', collapse);
  toggle.addEventListener('click', expand);

  // ---- Voice toggle ----
  function setVoiceEnabled(enabled) {
    chatState.voiceEnabled = enabled;
    // \uD83D\uDD0A = speaker with sound, \uD83D\uDD07 = muted speaker
    els.voiceToggle.textContent = enabled ? '\uD83D\uDD0A' : '\uD83D\uDD07';
    els.voiceToggle.classList.toggle('active', enabled);
    els.voiceToggle.title = enabled ? 'Voice narration ON (click to mute)' : 'Voice narration OFF (click to enable)';
  }

  els.voiceToggle.addEventListener('click', function () {
    var newState = !chatState.voiceEnabled;
    setVoiceEnabled(newState);
    if (window.__demoCopilotAction) {
      window.__demoCopilotAction(newState ? 'voice_enable' : 'voice_disable');
    }
  });

  // ---- Viewport height sync (fixes 100vh being wrong with browser chrome) ----
  function syncHeight() {
    var h = window.innerHeight;
    if (h > 0) {
      host.style.height = h + 'px';
    }
  }
  syncHeight();
  window.addEventListener('resize', syncHeight);
  // Also handle visual viewport changes (mobile, zoom, etc.)
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', syncHeight);
  }

  // Remove pulse after 6 seconds
  setTimeout(function () {
    var wrap = shadow.getElementById('chat-input-wrap');
    if (wrap) wrap.classList.remove('pulse');
  }, 6000);

  // ---- Welcome hints ----
  els.welcomeHints.addEventListener('click', function (e) {
    var hint = e.target.closest('.chat-welcome-hint');
    if (hint) {
      var text = hint.getAttribute('data-hint');
      els.input.value = text;
      els.input.dispatchEvent(new Event('input'));
      sendMessage();
    }
  });

  // ---- Quick actions ----
  els.quickActions.addEventListener('click', function (e) {
    var btn = e.target.closest('.chat-quick-btn');
    if (btn) {
      var action = btn.getAttribute('data-action');
      if (window.__demoCopilotAction) {
        window.__demoCopilotAction(action);
      }
    }
  });

  // ---- Scroll to bottom ----
  function scrollToBottom() {
    requestAnimationFrame(function () {
      els.messages.scrollTop = els.messages.scrollHeight;
    });
  }

  // ---- API Functions ----

  function addMessage(role, text) {
    if (els.messages.style.display === 'none') {
      els.messages.style.display = 'flex';
    }

    var msg = document.createElement('div');
    msg.className = 'chat-msg ' + role;

    var labels = { user: 'You', assistant: 'Demo Copilot', system: '' };
    var label = labels[role] || role;

    msg.innerHTML =
      (label ? '<div class="chat-msg-label">' + label + '</div>' : '') +
      '<div class="chat-msg-bubble">' + escapeHtml(text) + '</div>';

    if (role === 'assistant' && chatState.typingEl) {
      chatState.typingEl.remove();
      chatState.typingEl = null;
    }

    els.messages.appendChild(msg);
    scrollToBottom();
  }

  function addMessageHtml(role, html) {
    if (els.messages.style.display === 'none') {
      els.messages.style.display = 'flex';
    }

    var msg = document.createElement('div');
    msg.className = 'chat-msg ' + role;

    var labels = { user: 'You', assistant: 'Demo Copilot', system: '' };
    var label = labels[role] || role;

    msg.innerHTML =
      (label ? '<div class="chat-msg-label">' + label + '</div>' : '') +
      '<div class="chat-msg-bubble">' + html + '</div>';

    if (role === 'assistant' && chatState.typingEl) {
      chatState.typingEl.remove();
      chatState.typingEl = null;
    }

    els.messages.appendChild(msg);
    scrollToBottom();
  }

  function showTyping() {
    if (chatState.typingEl) return;
    if (els.messages.style.display === 'none') {
      els.messages.style.display = 'flex';
    }
    var el = document.createElement('div');
    el.className = 'chat-msg assistant';
    el.innerHTML =
      '<div class="chat-msg-label">Demo Copilot</div>' +
      '<div class="chat-typing">' +
        '<div class="chat-typing-dot"></div>' +
        '<div class="chat-typing-dot"></div>' +
        '<div class="chat-typing-dot"></div>' +
      '</div>';
    chatState.typingEl = el;
    els.messages.appendChild(el);
    scrollToBottom();
  }

  function hideTyping() {
    if (chatState.typingEl) {
      chatState.typingEl.remove();
      chatState.typingEl = null;
    }
  }

  function showPlan(plan) {
    if (chatState.planCardEl) {
      chatState.planCardEl.remove();
    }

    hideTyping();

    var card = document.createElement('div');
    card.className = 'chat-msg assistant';

    var stepsHtml = '';
    var stepIdx = 0;
    if (plan.sections) {
      for (var s = 0; s < plan.sections.length; s++) {
        var section = plan.sections[s];
        for (var t = 0; t < section.steps.length; t++) {
          var step = section.steps[t];
          stepsHtml +=
            '<div class="chat-plan-step" data-step-idx="' + stepIdx + '">' +
              '<span class="step-num">' + (stepIdx + 1) + '</span>' +
              '<span>' + escapeHtml(step.title) + '</span>' +
            '</div>';
          stepIdx++;
        }
      }
    }

    card.innerHTML =
      '<div class="chat-msg-label">Demo Copilot</div>' +
      '<div class="chat-plan-card">' +
        '<div class="chat-plan-header">\uD83D\uDCCB ' + escapeHtml(plan.title) + '</div>' +
        '<div class="chat-plan-steps">' + stepsHtml + '</div>' +
        '<div class="chat-plan-footer">' +
          '<button class="chat-plan-btn primary" data-action="start_demo">\u25B6 Start Demo</button>' +
          '<button class="chat-plan-btn secondary" data-action="modify_plan">\u270E Modify</button>' +
        '</div>' +
      '</div>';

    // Button handlers
    card.querySelector('[data-action="start_demo"]').addEventListener('click', function () {
      if (window.__demoCopilotAction) window.__demoCopilotAction('start_demo');
    });
    card.querySelector('[data-action="modify_plan"]').addEventListener('click', function () {
      if (window.__demoCopilotAction) window.__demoCopilotAction('modify_plan');
    });

    if (els.messages.style.display === 'none') {
      els.messages.style.display = 'flex';
    }
    els.messages.appendChild(card);
    chatState.planCardEl = card;
    scrollToBottom();
  }

  function updatePlanStep(idx, status) {
    if (!chatState.planCardEl) return;
    var steps = chatState.planCardEl.querySelectorAll('.chat-plan-step');
    if (idx < steps.length) {
      steps[idx].className = 'chat-plan-step ' + (status || '');
    }
  }

  function showProgress(step, total, label) {
    hideProgress();
    var pct = total > 0 ? Math.round(((step + 1) / total) * 100) : 0;
    var el = document.createElement('div');
    el.className = 'chat-msg system';
    el.innerHTML =
      '<div class="chat-demo-progress">' +
        '<div class="chat-progress-bar-wrap">' +
          '<div class="chat-progress-bar-fill" style="width: ' + pct + '%"></div>' +
        '</div>' +
        '<div class="chat-progress-label">' +
          '<span>' + escapeHtml(label || '') + '</span>' +
          '<span>Step ' + (step + 1) + '/' + total + '</span>' +
        '</div>' +
      '</div>';
    chatState.progressEl = el;
    els.messages.appendChild(el);
    scrollToBottom();
  }

  function updateProgress(step, total, label) {
    if (!chatState.progressEl) {
      showProgress(step, total, label);
      return;
    }
    var pct = total > 0 ? Math.round(((step + 1) / total) * 100) : 0;
    var fill = chatState.progressEl.querySelector('.chat-progress-bar-fill');
    if (fill) fill.style.width = pct + '%';
    var lbl = chatState.progressEl.querySelector('.chat-progress-label');
    if (lbl) {
      lbl.innerHTML = '<span>' + escapeHtml(label || '') + '</span><span>Step ' + (step + 1) + '/' + total + '</span>';
    }
    scrollToBottom();
  }

  function hideProgress() {
    if (chatState.progressEl) {
      chatState.progressEl.remove();
      chatState.progressEl = null;
    }
  }

  function setStatus(status, type) {
    els.status.textContent = status;
    els.status.className = 'chat-header-status' + (type ? ' ' + type : '');
  }

  function disable() {
    chatState.disabled = true;
    els.input.disabled = true;
    els.sendBtn.disabled = true;
  }

  function enable() {
    chatState.disabled = false;
    els.input.disabled = false;
    els.sendBtn.disabled = false;
    els.input.focus();
  }

  function showQuickActions(actions) {
    if (!actions || actions.length === 0) {
      hideQuickActions();
      return;
    }
    els.quickActions.innerHTML = actions.map(function (a) {
      return '<button class="chat-quick-btn' + (a.danger ? ' danger' : '') + '" data-action="' + escapeHtml(a.action) + '">' + escapeHtml(a.label) + '</button>';
    }).join('');
    els.quickActions.style.display = 'flex';
  }

  function hideQuickActions() {
    els.quickActions.style.display = 'none';
  }

  function clearMessages() {
    els.messages.innerHTML = '';
    chatState.planCardEl = null;
    chatState.progressEl = null;
    chatState.typingEl = null;
  }

  function showWelcome() {
    els.welcome.style.display = 'flex';
    els.messages.style.display = 'none';
  }

  function hideWelcome() {
    els.welcome.style.display = 'none';
    if (els.messages.style.display === 'none') {
      els.messages.style.display = 'flex';
    }
  }

  function getWidth() {
    return host.offsetWidth || 400;
  }

  // ---- Step Tracker Functions ----

  function showStepTracker(plan) {
    var html = '';
    var stepIdx = 0;
    if (plan.sections) {
      for (var s = 0; s < plan.sections.length; s++) {
        var section = plan.sections[s];
        if (plan.sections.length > 1) {
          html +=
            '<div class="chat-tracker-section-label">' +
              escapeHtml(section.title) +
            '</div>';
        }
        for (var t = 0; t < section.steps.length; t++) {
          var step = section.steps[t];
          html +=
            '<div class="chat-tracker-step" data-tracker-idx="' + stepIdx + '">' +
              '<span class="tracker-step-indicator">' + (stepIdx + 1) + '</span>' +
              '<span class="tracker-step-title">' + escapeHtml(step.title) + '</span>' +
            '</div>';
          stepIdx++;
        }
      }
    }
    els.trackerBody.innerHTML = html;
    trackerCollapsed = false;
    els.trackerBody.style.display = 'block';
    els.trackerToggle.textContent = '\u25BC';
    els.stepTracker.style.display = 'block';
  }

  function updateTrackerStep(idx, status) {
    var steps = els.trackerBody.querySelectorAll('.chat-tracker-step');
    if (idx < steps.length) {
      steps[idx].className = 'chat-tracker-step ' + (status || '');
      // Auto-scroll the active step into view within the tracker body
      if (status === 'active') {
        steps[idx].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    }
  }

  function hideStepTracker() {
    els.stepTracker.style.display = 'none';
    els.trackerBody.innerHTML = '';
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  // ---- Expose API ----
  window.DemoCopilotChat = {
    addMessage: addMessage,
    addMessageHtml: addMessageHtml,
    showTyping: showTyping,
    hideTyping: hideTyping,
    showPlan: showPlan,
    updatePlanStep: updatePlanStep,
    showProgress: showProgress,
    updateProgress: updateProgress,
    hideProgress: hideProgress,
    setStatus: setStatus,
    disable: disable,
    enable: enable,
    showQuickActions: showQuickActions,
    hideQuickActions: hideQuickActions,
    clear: clearMessages,
    collapse: collapse,
    expand: expand,
    showWelcome: showWelcome,
    hideWelcome: hideWelcome,
    getWidth: getWidth,
    setVoiceEnabled: setVoiceEnabled,
    showStepTracker: showStepTracker,
    updateTrackerStep: updateTrackerStep,
    hideStepTracker: hideStepTracker,
  };

  console.log('[DemoCopilotChat] Sidecar chat panel initialized (Shadow DOM)');
})();
