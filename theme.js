// Shared navbar interactivity (hamburger toggle + scroll-shadow).
// Every page with the dark navbar loads this, including index.html.

(function () {
  'use strict';

  const navbar = document.getElementById('navbar');
  const hamburger = document.getElementById('hamburger');
  const navMobile = document.getElementById('navMobile');

  if (navbar) {
    let scrolled = navbar.classList.contains('scrolled');
    // Passive: this listener never calls preventDefault, and the class is only
    // written when the state actually changes rather than on every event.
    window.addEventListener('scroll', function () {
      const next = window.scrollY > 48;
      if (next === scrolled) return;
      scrolled = next;
      navbar.classList.toggle('scrolled', next);
    }, { passive: true });
  }

  if (hamburger && navMobile) {
    if (!navMobile.id) navMobile.id = 'navMobile';
    hamburger.setAttribute('aria-controls', navMobile.id);
    hamburger.setAttribute('aria-expanded', 'false');

    const setOpen = function (open) {
      navMobile.classList.toggle('open', open);
      hamburger.setAttribute('aria-expanded', String(open));
      hamburger.setAttribute('aria-label', open ? 'Close menu' : 'Open menu');
    };

    hamburger.addEventListener('click', function (event) {
      event.stopPropagation();
      setOpen(!navMobile.classList.contains('open'));
    });

    navMobile.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () { setOpen(false); });
    });

    // Escape closes the menu and returns focus to the control that opened it.
    document.addEventListener('keydown', function (event) {
      if (event.key !== 'Escape' || !navMobile.classList.contains('open')) return;
      setOpen(false);
      hamburger.focus();
    });

    // A tap anywhere outside the panel closes it, which is what the gesture
    // means on a phone.
    document.addEventListener('click', function (event) {
      if (!navMobile.classList.contains('open')) return;
      if (navMobile.contains(event.target) || hamburger.contains(event.target)) return;
      setOpen(false);
    });

    // Leaving the mobile breakpoint leaves the panel stranded open otherwise.
    window.addEventListener('resize', function () {
      if (window.innerWidth > 768 && navMobile.classList.contains('open')) setOpen(false);
    });
  }
})();

// Keep the footer year current without a build step.
(function () {
  var year = String(new Date().getFullYear());
  document.querySelectorAll('[data-current-year]').forEach(function (el) {
    el.textContent = year;
  });
})();

// Mark the nav item for the page you are on, so its position is announced and
// visible rather than left to guesswork.
(function () {
  var here = location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a, .nav-mobile a').forEach(function (link) {
    var target = (link.getAttribute('href') || '').split('#')[0];
    if (!target) return;
    if (target === here) {
      link.setAttribute('aria-current', 'page');
    }
  });
})();

// Optional enterprise training features for Microsoft 365-backed deployments.
// The public site stays anonymous by default. An organization can opt in by
// setting window.JFPEnterprise = { apiUrl: 'https://...', requiredSSO: true }.
(function () {
  'use strict';

  var PROGRESS_KEY = 'phishing-training-progress';
  var IDENTITY_KEY = 'jfp-enterprise-identity';
  var LOG_KEY = 'jfp-enterprise-log-state';

  function safeJsonParse(raw, fallback) {
    if (!raw) return fallback;
    try {
      return JSON.parse(raw);
    } catch (_error) {
      return fallback;
    }
  }

  function getProgress() {
    return safeJsonParse(localStorage.getItem(PROGRESS_KEY), {});
  }

  function getIdentity() {
    return safeJsonParse(localStorage.getItem(IDENTITY_KEY), null);
  }

  function setIdentity(identity) {
    if (!identity || !identity.email) return;
    localStorage.setItem(IDENTITY_KEY, JSON.stringify(identity));
  }

  function clearIdentity() {
    localStorage.removeItem(IDENTITY_KEY);
  }

  function getLoggedModules() {
    return safeJsonParse(localStorage.getItem(LOG_KEY), {});
  }

  function setLoggedModules(data) {
    localStorage.setItem(LOG_KEY, JSON.stringify(data));
  }

  function buildIdentityPayload(identity) {
    if (!identity) return null;
    return {
      email: identity.email || '',
      name: identity.name || '',
      tenantId: identity.tenantId || '',
      objectId: identity.objectId || '',
      authProvider: identity.authProvider || 'microsoft-365'
    };
  }

  function buildTrainingEvent(moduleName, entry) {
    var user = getIdentity();
    var payload = {
      eventType: 'training_completion',
      moduleName: moduleName,
      score: Number(entry.score) || 0,
      total: Number(entry.total) || 0,
      percentage: Number(entry.percentage) || 0,
      completedAt: entry.completedAt || new Date().toISOString(),
      source: 'just-for-phishing',
      user: buildIdentityPayload(user)
    };
    if (!payload.user) {
      payload.user = null;
    }
    return payload;
  }

  function sendTrainingEvent(event) {
    var config = window.JFPEnterprise || {};
    var apiUrl = config.apiUrl;
    if (!apiUrl || !config.enabled && !config.apiUrl) return Promise.resolve();

    return fetch(apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(event)
    }).catch(function () {
      // Best effort only: keep the public site working even if backend logging is
      // not available or the user is not signed in to the enterprise flow.
    });
  }

  function flushCompletedProgress() {
    var config = window.JFPEnterprise || {};
    if (!config.apiUrl && !config.enabled) return;

    var progress = getProgress();
    var logged = getLoggedModules();

    Object.keys(progress).forEach(function (moduleName) {
      var entry = progress[moduleName];
      if (!entry || !entry.completed || !entry.completedAt) return;
      var marker = entry.completedAt;
      if (logged[moduleName] === marker) return;

      var event = buildTrainingEvent(moduleName, entry);
      var user = getIdentity();
      if (config.requiredSSO && !user) {
        logged[moduleName] = marker;
        setLoggedModules(logged);
        return;
      }

      sendTrainingEvent(event).then(function () {
        logged[moduleName] = marker;
        setLoggedModules(logged);
      });
    });
  }

  function readMicrosoftIdentityFromQuery() {
    try {
      var params = new URLSearchParams(window.location.search);
      var email = params.get('userEmail');
      var name = params.get('userName');
      var tenantId = params.get('tenantId');
      var objectId = params.get('objectId');
      if (!email) return null;
      return {
        email: email,
        name: name || email,
        tenantId: tenantId || '',
        objectId: objectId || '',
        authProvider: 'microsoft-365'
      };
    } catch (_error) {
      return null;
    }
  }

  function initialiseEnterpriseTracking() {
    var identity = getIdentity() || readMicrosoftIdentityFromQuery();
    if (identity) setIdentity(identity);
    flushCompletedProgress();
  }

  window.JFPTrainingAuth = window.JFPTrainingAuth || {
    setIdentity: setIdentity,
    clearIdentity: clearIdentity,
    getIdentity: getIdentity,
    readMicrosoftIdentityFromQuery: readMicrosoftIdentityFromQuery,
    flushCompletedProgress: flushCompletedProgress
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialiseEnterpriseTracking);
  } else {
    initialiseEnterpriseTracking();
  }

  window.addEventListener('storage', function (event) {
    if (event.key !== PROGRESS_KEY) return;
    flushCompletedProgress();
  });
})();
