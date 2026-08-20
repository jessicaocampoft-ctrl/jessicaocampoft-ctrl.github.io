(function () {
  'use strict';

  var scheduleConfig = null;
  var loaded = false;
  var bookingSubmitting = false;
  var bookingAmbiguous = false;

  function venueKey(modality) {
    var value = String(modality || '').toLowerCase();
    if (value.indexOf('santa') >= 0) return 'santa';
    if (value.indexOf('recovery') >= 0 || value.indexOf('campestre') >= 0) return 'recovery';
    return 'domicilio';
  }

  function venueEnabled(key) {
    if (key === 'recovery') return false;
    if (key === 'domicilio') return true;
    if (!scheduleConfig || !scheduleConfig.venues || !scheduleConfig.venues[key]) return true;
    return scheduleConfig.venues[key].enabled !== false;
  }

  function serviceAllowed(key, service) {
    if (key === 'recovery') return false;
    if (key === 'domicilio') return true;
    if (!scheduleConfig || !scheduleConfig.venues || !scheduleConfig.venues[key]) return true;
    var services = scheduleConfig.venues[key].services;
    if (!Array.isArray(services) || !services.length) return false;
    return services.indexOf(String(service || '')) >= 0;
  }

  function allowedForCurrentService(key) {
    if (key === 'recovery') return false;
    if (!venueEnabled(key)) return false;
    if (key === 'domicilio') return true;
    var service = (typeof bk !== 'undefined' && bk) ? bk.service : '';
    if (!service) return true;
    return serviceAllowed(key, service);
  }

  function stripRecoveryFromMarkup() {
    var btn = document.getElementById('modRecovery');
    if (btn) {
      btn.style.display = 'none';
      btn.disabled = true;
      btn.setAttribute('aria-hidden', 'true');
    }
    document.querySelectorAll('.bk-svc').forEach(function (card) {
      var venues = String(card.dataset.venues || '')
        .split(',')
        .map(function (v) { return v.trim(); })
        .filter(function (v) { return v && v !== 'recovery'; });
      card.dataset.venues = venues.join(',');
      card.dataset.priceRecovery = '';
      var price = card.querySelector('.bk-svc-price');
      if (price) {
        price.textContent = String(price.textContent || '')
          .replace(/\s*\/\s*Recovery\b/g, '')
          .replace(/\s{2,}/g, ' ')
          .trim();
      }
    });
    try {
      if (typeof bk !== 'undefined' && bk) {
        bk.venues = String(bk.venues || '')
          .split(',')
          .map(function (v) { return v.trim(); })
          .filter(function (v) { return v && v !== 'recovery'; })
          .join(',');
        bk.priceRecovery = '';
        if (venueKey(bk.modality) === 'recovery') {
          bk.modality = (',' + bk.venues + ',').indexOf(',santa,') >= 0 ? 'Sede Santa Mónica' : 'Domicilio';
        }
      }
    } catch (_) {}
  }

  function setButtonVisibility(id, key) {
    var btn = document.getElementById(id);
    if (!btn) return;
    if (key === 'recovery') {
      btn.style.display = 'none';
      btn.disabled = true;
      btn.setAttribute('aria-hidden', 'true');
      return;
    }
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
    stripRecoveryFromMarkup();
    var currentKey = venueKey(bk.modality);
    if (currentKey === 'domicilio') return;
    if (allowedForCurrentService(currentKey) && (!bk.venues || typeof venueAllowed !== 'function' || venueAllowed(currentKey))) return;

    var candidates = [
      { key: 'santa', value: 'Sede Santa Mónica' },
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
    stripRecoveryFromMarkup();
    setButtonVisibility('modSanta', 'santa');
    setButtonVisibility('modRecovery', 'recovery');
    setButtonVisibility('modD', 'domicilio');
    pickFallbackVenue();

    var santa = document.getElementById('modSanta');
    var recovery = document.getElementById('modRecovery');
    var domicilio = document.getElementById('modD');
    if (typeof bk !== 'undefined' && bk) {
      if (santa) santa.classList.toggle('active', bk.modality === 'Sede Santa Mónica');
      if (recovery) recovery.classList.remove('active');
      if (domicilio) domicilio.classList.toggle('active', bk.modality === 'Domicilio');
      var addressWrap = document.getElementById('addressWrap');
      if (addressWrap) addressWrap.style.display = bk.modality === 'Domicilio' ? 'block' : 'none';
      var notice = document.getElementById('dominotice');
      if (notice) notice.style.display = bk.modality === 'Domicilio' ? 'block' : 'none';
    }
  }

  function fetchJsonWithTimeout(url, options, timeoutMs) {
    var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = null;
    var opts = Object.assign({}, options || {});
    opts.cache = 'no-store';
    if (controller) {
      opts.signal = controller.signal;
      timer = setTimeout(function () { controller.abort(); }, timeoutMs || 30000);
    }

    return fetch(url, opts).then(function (response) {
      return response.text().then(function (raw) {
        var data;
        try {
          data = JSON.parse(raw);
        } catch (e) {
          var parseError = new Error('El servidor respondió en un formato inesperado.');
          parseError.code = 'INVALID_JSON';
          throw parseError;
        }
        if (!response.ok) {
          var httpError = new Error((data && data.error) || 'El servidor rechazó la solicitud.');
          httpError.code = 'HTTP_ERROR';
          throw httpError;
        }
        return data;
      });
    }).then(function (data) {
      if (timer) clearTimeout(timer);
      return data;
    }, function (error) {
      if (timer) clearTimeout(timer);
      throw error;
    });
  }

  function availabilityUrl() {
    return APPS_SCRIPT_URL
      + '?action=availability&date=' + encodeURIComponent(bk.date || '')
      + '&service=' + encodeURIComponent(bk.service || '')
      + '&modality=' + encodeURIComponent(bk.modality || 'Presencial')
      + '&_ts=' + Date.now();
  }

  function verifyCurrentSlot() {
    if (typeof APPS_SCRIPT_URL === 'undefined' || typeof bk === 'undefined' || !bk) {
      return Promise.reject(new Error('El sistema de agenda no está disponible en este momento.'));
    }
    if (!bk.date || !bk.time) {
      return Promise.reject(new Error('Selecciona nuevamente la fecha y la hora.'));
    }
    if (venueKey(bk.modality) === 'recovery') {
      bk.modality = 'Sede Santa Mónica';
      stripRecoveryFromMarkup();
      return Promise.reject(new Error('Esa sede ya no está disponible para agendamiento. Selecciona Santa Mónica o domicilio.'));
    }
    return fetchJsonWithTimeout(availabilityUrl(), {}, 20000).then(function (data) {
      var slots = data && data.slots ? data.slots : {};
      if (slots[bk.time] !== true) {
        var taken = new Error('Ese horario ya no está disponible. Selecciona otro cupo.');
        taken.code = 'SLOT_TAKEN';
        taken.slots = slots;
        throw taken;
      }
      return data;
    });
  }

  function showStep2AvailabilityError(message) {
    var msg = document.getElementById('bkAvailMsg');
    if (msg) {
      msg.textContent = message;
      msg.style.display = 'block';
      msg.style.color = '#b91c1c';
    }
  }

  function clearBookingError() {
    var error = document.getElementById('bkError');
    if (!error) return;
    error.textContent = '';
    error.classList.remove('visible');
  }

  function showBookingError(message, ambiguous) {
    var error = document.getElementById('bkError');
    if (!error) return;
    error.innerHTML = '';
    var text = document.createElement('span');
    text.textContent = message;
    error.appendChild(text);

    if (ambiguous) {
      var br = document.createElement('br');
      var link = document.createElement('a');
      var ref = (typeof ensureReservationCode === 'function') ? ensureReservationCode() : '';
      var wa = 'Hola, intenté reservar en la página pero el sistema no pudo confirmar si quedó guardada.'
        + (ref ? ' Código: ' + ref + '.' : '')
        + ' ¿Me ayudan a verificarla antes de que la intente nuevamente?';
      link.href = 'https://wa.me/573136467945?text=' + encodeURIComponent(wa);
      link.target = '_blank';
      link.rel = 'noopener';
      link.textContent = 'Verificar por WhatsApp';
      link.style.display = 'inline-block';
      link.style.marginTop = '8px';
      link.style.fontWeight = '700';
      link.style.color = 'inherit';
      error.appendChild(br);
      error.appendChild(link);
    }
    error.classList.add('visible');
  }

  function restoreStep3AfterFailure(message, ambiguous) {
    var sending = document.getElementById('bkSending');
    if (sending) sending.classList.remove('active');
    if (typeof goStep === 'function') goStep(3);
    showBookingError(message, ambiguous);

    var btn = document.getElementById('btnStep3');
    if (btn) {
      btn.disabled = !!ambiguous;
      btn.dataset.bookingLocked = ambiguous ? '1' : '0';
    }
  }

  function collectBookingFields() {
    bk.name = (document.getElementById('bkName') || {}).value ? document.getElementById('bkName').value.trim() : '';
    bk.phone = (document.getElementById('bkPhone') || {}).value ? document.getElementById('bkPhone').value.trim() : '';
    bk.email = (document.getElementById('bkEmail') || {}).value ? document.getElementById('bkEmail').value.trim() : '';
    bk.address = (document.getElementById('bkAddress') || {}).value ? document.getElementById('bkAddress').value.trim() : '';
    bk.notes = (document.getElementById('bkNotes') || {}).value ? document.getElementById('bkNotes').value.trim() : '';
    bk.paraQuien = (document.getElementById('bkParaQuien') || {}).value ? document.getElementById('bkParaQuien').value.trim() : '';
  }

  function bookingPayload() {
    if (!bk.clientTimestamp) bk.clientTimestamp = Date.now();
    var ref = (typeof ensureReservationCode === 'function') ? ensureReservationCode() : (bk.reservationCode || '');
    if (!bk.reservationCode && ref) bk.reservationCode = ref;
    return {
      service: bk.service,
      modality: bk.modality,
      date: bk.date,
      time: bk.time,
      name: bk.name,
      phone: bk.phone,
      email: bk.email,
      address: bk.address,
      notes: bk.notes,
      priceP: bk.priceSanta || bk.priceP,
      priceD: bk.priceD,
      priceRecovery: '',
      priceSelected: typeof currentBookingPrice === 'function' ? currentBookingPrice() : '',
      duration: bk.duration,
      source: 'PAGINA_WEB',
      estadoPago: 'PENDIENTE_PAGO',
      estadoCita: 'RESERVADA',
      policyAccepted: true,
      clientTimestamp: bk.clientTimestamp,
      codigoReserva: bk.reservationCode,
      notaAdmin: '[CODIGO RESERVA: ' + bk.reservationCode + ']' + (bk.paraQuien ? ' [PARA: ' + bk.paraQuien + ']' : '')
    };
  }

  function installReliableBooking() {
    if (typeof window.checkAndContinue === 'function' && !window.checkAndContinue.__reliableBooking) {
      var reliableContinue = function () {
        stripRecoveryFromMarkup();
        if (!bk.date || !bk.time) {
          showStep2AvailabilityError('Selecciona fecha y hora para continuar.');
          return;
        }
        var btn = document.getElementById('btnStep2');
        var origText = btn ? btn.innerHTML : '';
        if (btn) {
          btn.disabled = true;
          btn.innerHTML = 'Verificando...';
        }
        var msg = document.getElementById('bkAvailMsg');
        if (msg) msg.style.display = 'none';

        verifyCurrentSlot().then(function () {
          if (btn) {
            btn.innerHTML = origText;
            btn.disabled = false;
          }
          if (typeof goStep === 'function') goStep(3);
        }).catch(function (error) {
          if (error && error.slots && typeof applyAvailability === 'function') applyAvailability(error.slots);
          showStep2AvailabilityError(error && error.message
            ? error.message
            : 'No pudimos verificar el horario. Intenta nuevamente o escríbenos por WhatsApp.');
          if (btn) {
            btn.innerHTML = origText;
            btn.disabled = false;
          }
        });
      };
      reliableContinue.__reliableBooking = true;
      window.checkAndContinue = reliableContinue;
    }

    if (typeof window.submitBooking === 'function' && !window.submitBooking.__reliableBooking) {
      var reliableSubmit = function () {
        if (bookingSubmitting || bookingAmbiguous) return;
        stripRecoveryFromMarkup();
        collectBookingFields();
        clearBookingError();

        var accepted = document.getElementById('bkPaymentAccept');
        if (!accepted || !accepted.checked) {
          showBookingError('Debes aceptar la política de pago anticipado para continuar.', false);
          return;
        }
        if (!bk.name) {
          showBookingError('Escribe tu nombre completo para continuar.', false);
          return;
        }
        var digits = String(bk.phone || '').replace(/\D/g, '');
        var phoneOk = digits.length === 10 || (digits.length === 12 && digits.indexOf('57') === 0);
        if (!phoneOk) {
          showBookingError('Revisa tu número de WhatsApp. Debe tener 10 dígitos o iniciar con +57.', false);
          return;
        }
        if (!bk.service || !bk.date || !bk.time) {
          showBookingError('Faltan datos de la reserva. Vuelve y selecciona servicio, fecha y hora.', false);
          return;
        }
        if (venueKey(bk.modality) === 'recovery') {
          bk.modality = 'Sede Santa Mónica';
          showBookingError('Esa sede ya no está disponible para agendamiento. Selecciona Santa Mónica o domicilio.', false);
          return;
        }

        bookingSubmitting = true;
        var btn = document.getElementById('btnStep3');
        if (btn) btn.disabled = true;

        verifyCurrentSlot().then(function () {
          document.querySelectorAll('.bk-step').forEach(function (s) { s.classList.remove('active'); });
          var sending = document.getElementById('bkSending');
          if (sending) sending.classList.add('active');

          var payload = bookingPayload();
          return fetchJsonWithTimeout(APPS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify(payload)
          }, 60000).then(function (data) {
            if (!data || data.ok !== true) {
              var rejected = new Error((data && data.error) || 'El servidor no confirmó la reserva.');
              rejected.code = 'BOOKING_REJECTED';
              throw rejected;
            }
            if (!data.id) {
              var noId = new Error('El servidor respondió, pero no devolvió el identificador de la cita.');
              noId.code = 'AMBIGUOUS_RESPONSE';
              throw noId;
            }
            bk.serverBookingId = data.id;
            bookingSubmitting = false;
            if (typeof showSuccess === 'function') showSuccess();
          });
        }).catch(function (error) {
          bookingSubmitting = false;

          if (error && error.code === 'SLOT_TAKEN') {
            if (error.slots && typeof applyAvailability === 'function') applyAvailability(error.slots);
            if (typeof goStep === 'function') goStep(2);
            showStep2AvailabilityError(error.message);
            if (btn) btn.disabled = false;
            return;
          }

          if (error && error.code === 'BOOKING_REJECTED') {
            restoreStep3AfterFailure(error.message, false);
            return;
          }

          bookingAmbiguous = true;
          restoreStep3AfterFailure(
            'No pudimos confirmar automáticamente si la reserva quedó guardada. No la envíes de nuevo todavía; verifícala por WhatsApp con tu código de reserva.',
            true
          );
        });
      };
      reliableSubmit.__reliableBooking = true;
      window.submitBooking = reliableSubmit;
    }
  }

  function wrapBookingFunctions() {
    if (typeof window.selectService === 'function' && !window.selectService.__publicScheduleWrapped) {
      var originalSelectService = window.selectService;
      var wrappedSelectService = function (el) {
        stripRecoveryFromMarkup();
        var result = originalSelectService.apply(this, arguments);
        stripRecoveryFromMarkup();
        renderVenueState();
        return result;
      };
      wrappedSelectService.__publicScheduleWrapped = true;
      window.selectService = wrappedSelectService;
    }

    if (typeof window.updateVenueButtons === 'function' && !window.updateVenueButtons.__publicScheduleWrapped) {
      var originalUpdateVenueButtons = window.updateVenueButtons;
      var wrappedUpdateVenueButtons = function () {
        stripRecoveryFromMarkup();
        var result = originalUpdateVenueButtons.apply(this, arguments);
        stripRecoveryFromMarkup();
        renderVenueState();
        return result;
      };
      wrappedUpdateVenueButtons.__publicScheduleWrapped = true;
      window.updateVenueButtons = wrappedUpdateVenueButtons;
    }

    if (typeof window.selectModality === 'function' && !window.selectModality.__publicScheduleWrapped) {
      var originalSelectModality = window.selectModality;
      var wrappedSelectModality = function (m) {
        if (venueKey(m) === 'recovery') return false;
        var key = venueKey(m);
        if (!allowedForCurrentService(key)) return;
        var result = originalSelectModality.apply(this, arguments);
        stripRecoveryFromMarkup();
        return result;
      };
      wrappedSelectModality.__publicScheduleWrapped = true;
      window.selectModality = wrappedSelectModality;
    }

    installReliableBooking();
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
        stripRecoveryFromMarkup();
        renderVenueState();
      });
  }

  function init() {
    stripRecoveryFromMarkup();
    wrapBookingFunctions();
    renderVenueState();
    loadConfig();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
