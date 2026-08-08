(function () {
  'use strict';

  var scheduleConfig = null;
  var loaded = false;

  function venueKey(modality) {
    var value = String(modality || '').toLowerCase();
    if (value.indexOf('santa') >= 0) return 'santa';
    if (value.indexOf('recovery') >= 0 || value.indexOf('campestre') >= 0) return 'recovery';
    return 'domicilio';
  }

  function venueEnabled(key) {
    if (key === 'domicilio') return true;
    if (!scheduleConfig || !scheduleConfig.venues || !scheduleConfig.venues[key]) return true;
    return scheduleConfig.venues[key].enabled !== false;
  }

  function serviceAllowed(key, service) {
    if (key === 'domicilio') return true;
    if (!scheduleConfig || !scheduleConfig.venues || !scheduleConfig.venues[key]) return true;
    var services = scheduleConfig.venues[key].services;
    if (!Array.isArray(services) || !services.length) return false;
    return services.indexOf(String(service || '')) >= 0;
  }

  function allowedForCurrentService(key) {
    if (!venueEnabled(key)) return false;
    if (key === 'domicilio') return true;
    var service = (typeof bk !== 'undefined' && bk) ? bk.service : '';
    if (!service) return true;
    return serviceAllowed(key, service);
  }

  function setButtonVisibility(id, key) {
    var btn = document.getElementById(id);
    if (!btn) return;
    var allowedByServiceMarkup = true;
    if (typeof venueAllowed === 'function' && typeof bk !== 'undefined' && bk && bk.venues) {
      allowedByServiceMarkup = venueAllowed(key);
    }
    var visible = allowedByServiceMarkup && allowedForCurrentService(key);
    btn.style.display = visible ? '' : 'none';
    btn.disabled = !visible;
    btn.setAttribute('aria-hidden', visible ? 'false' : 'true');
  }

  function pickFallbackVenue() {
    if (typeof bk === 'undefined' || !bk) return;
    var currentKey = venueKey(bk.modality);
    if (currentKey === 'domicilio') return;
    if (allowedForCurrentService(currentKey) && (!bk.venues || typeof venueAllowed !== 'function' || venueAllowed(currentKey))) return;

    var candidates = [
      { key: 'santa', value: 'Sede Santa Mónica' },
      { key: 'recovery', value: 'Sede Campestre Recovery' },
      { key: 'domicilio', value: 'Domicilio' }
    ];
    for (var i = 0; i < candidates.length; i++) {
      var c = candidates[i];
      var markupAllows = !bk.venues || typeof venueAllowed !== 'function' || venueAllowed(c.key);
      if (markupAllows && allowedForCurrentService(c.key)) {
        bk.modality = c.value;
        break;
      }
    }
  }

  function renderVenueState() {
    if (!loaded) return;
    setButtonVisibility('modSanta', 'santa');
    setButtonVisibility('modRecovery', 'recovery');
    setButtonVisibility('modD', 'domicilio');
    pickFallbackVenue();

    var santa = document.getElementById('modSanta');
    var recovery = document.getElementById('modRecovery');
    var domicilio = document.getElementById('modD');
    if (typeof bk !== 'undefined' && bk) {
      if (santa) santa.classList.toggle('active', bk.modality === 'Sede Santa Mónica');
      if (recovery) recovery.classList.toggle('active', bk.modality === 'Sede Campestre Recovery');
      if (domicilio) domicilio.classList.toggle('active', bk.modality === 'Domicilio');
      var addressWrap = document.getElementById('addressWrap');
      if (addressWrap) addressWrap.style.display = bk.modality === 'Domicilio' ? 'block' : 'none';
      var notice = document.getElementById('dominotice');
      if (notice) notice.style.display = bk.modality === 'Domicilio' ? 'block' : 'none';
    }
  }

  function wrapBookingFunctions() {
    if (typeof window.selectService === 'function' && !window.selectService.__publicScheduleWrapped) {
      var originalSelectService = window.selectService;
      var wrappedSelectService = function (el) {
        var result = originalSelectService.apply(this, arguments);
        renderVenueState();
        return result;
      };
      wrappedSelectService.__publicScheduleWrapped = true;
      window.selectService = wrappedSelectService;
    }

    if (typeof window.updateVenueButtons === 'function' && !window.updateVenueButtons.__publicScheduleWrapped) {
      var originalUpdateVenueButtons = window.updateVenueButtons;
      var wrappedUpdateVenueButtons = function () {
        var result = originalUpdateVenueButtons.apply(this, arguments);
        renderVenueState();
        return result;
      };
      wrappedUpdateVenueButtons.__publicScheduleWrapped = true;
      window.updateVenueButtons = wrappedUpdateVenueButtons;
    }

    if (typeof window.selectModality === 'function' && !window.selectModality.__publicScheduleWrapped) {
      var originalSelectModality = window.selectModality;
      var wrappedSelectModality = function (m) {
        var key = venueKey(m);
        if (!allowedForCurrentService(key)) return;
        return originalSelectModality.apply(this, arguments);
      };
      wrappedSelectModality.__publicScheduleWrapped = true;
      window.selectModality = wrappedSelectModality;
    }
  }

  function loadConfig() {
    if (typeof APPS_SCRIPT_URL === 'undefined') return;
    fetch(APPS_SCRIPT_URL + '?action=publicScheduleConfig&_ts=' + Date.now(), { cache: 'no-store' })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data || !data.ok || !data.config) return;
        scheduleConfig = data.config;
        window.CUIDANDOTE_PUBLIC_SCHEDULE = scheduleConfig;
        loaded = true;
        wrapBookingFunctions();
        renderVenueState();
      })
      .catch(function () {
        // Si la configuración no puede consultarse, conservar el comportamiento actual.
      });
  }

  function init() {
    wrapBookingFunctions();
    loadConfig();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
