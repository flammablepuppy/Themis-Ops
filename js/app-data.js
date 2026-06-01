function normalizeMissionText(value, uppercase = false) {
    const text = value == null ? '' : String(value).trim();
    return uppercase ? text.toUpperCase() : text;
}

function parseMissionDate(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? new Date(NaN) : date;
}

function normalizeMissionLeg(leg) {
    const source = leg && typeof leg === 'object' ? leg : {};

    return {
        takeoffIcao: normalizeMissionText(source.takeoffIcao, true),
        takeoffTime: parseMissionDate(source.takeoffTime),
        landIcao: normalizeMissionText(source.landIcao, true),
        landTime: parseMissionDate(source.landTime)
    };
}

function normalizeMissionLift(lift) {
    const source = lift && typeof lift === 'object' ? lift : {};

    return {
        customer: normalizeMissionText(source.customer ?? source.liftCustomer),
        pax: normalizeMissionText(source.pax ?? source.liftPax),
        cargo: normalizeMissionText(source.cargo ?? source.liftCargo),
        hazmat: normalizeMissionText(source.hazmat ?? source.liftHazmat)
    };
}

function isMissionLiftEmpty(lift) {
    return !lift.customer && !lift.pax && !lift.cargo && !lift.hazmat;
}

function getMissionLiftsForRecord(mission) {
    const source = mission && typeof mission === 'object' ? mission : {};
    const lifts = Array.isArray(source.lifts)
        ? source.lifts.map(normalizeMissionLift).filter(lift => !isMissionLiftEmpty(lift))
        : [];

    if (lifts.length > 0) {
        return lifts;
    }

    const legacyLift = normalizeMissionLift(source);
    return isMissionLiftEmpty(legacyLift) ? [] : [legacyLift];
}

function normalizeMissionRecord(mission, index = 0) {
    if (!mission || typeof mission !== 'object') return null;

    const normalized = { ...mission };
    const missionId = mission.id != null && String(mission.id).trim() !== ''
        ? String(mission.id)
        : createMissionId();
    normalized.id = missionId;
    normalized.missionNum = normalizeMissionText(mission.missionNum, true);
    normalized.tailNum = normalizeMissionText(mission.tailNum);
    normalized.pilot = normalizeMissionText(mission.pilot);
    normalized.copilot = normalizeMissionText(mission.copilot);
    normalized.crewChief = normalizeMissionText(mission.crewChief);
    normalized.loadmaster = normalizeMissionText(mission.loadmaster);
    normalized.lifts = getMissionLiftsForRecord(mission);
    const primaryLift = normalized.lifts[0] || {
        customer: '',
        pax: '',
        cargo: '',
        hazmat: ''
    };
    normalized.liftCustomer = primaryLift.customer;
    normalized.liftPax = primaryLift.pax;
    normalized.liftCargo = primaryLift.cargo;
    normalized.liftHazmat = primaryLift.hazmat;
    normalized.legs = Array.isArray(mission.legs)
        ? mission.legs.map(normalizeMissionLeg).sort((a, b) => getDateTimestamp(a.takeoffTime) - getDateTimestamp(b.takeoffTime))
        : [];

    return normalized;
}

function loadMissions() {
    try {
        const cachedDocument = loadMissionCanonicalDocumentFromCache();
        if (!cachedDocument) return null;

        applyMissionCanonicalDocumentToRuntime(cachedDocument, { persistLocalCache: false });
        return missions;
    } catch {
        return null;
    }
}

function persistMissions() {
    const document = buildMissionCanonicalDocument();
    applyMissionCanonicalDocumentToRuntime(document, { persistLocalCache: true });
    missionCanonicalLocalDirty = true;

    try {
        void queueMissionCanonicalDocumentWrite(document);
    } catch {
        // Ignore storage failures in file:// or restricted browser contexts.
    }
}

function normalizeIcao(code) {
    return (code || '').trim().toUpperCase();
}

function normalizeAirportExactKey(value) {
    return (value == null ? '' : String(value))
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, '');
}

function normalizeAirportSearchText(value) {
    return (value == null ? '' : String(value))
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]+/g, ' ')
        .trim()
        .replace(/\s+/g, ' ');
}

function addAirportExactLookupKey(index, key, airport) {
    const normalizedKey = normalizeAirportExactKey(key);
    if (!normalizedKey) return;

    if (!index.has(normalizedKey)) {
        index.set(normalizedKey, []);
    }

    index.get(normalizedKey).push(airport);
}

function getAirportLookupContext(airport) {
    const parts = [];
    const municipality = normalizeMissionText(airport && airport.municipality);
    const isoRegion = normalizeMissionText(airport && airport.isoRegion);
    const isoCountry = normalizeMissionText(airport && airport.isoCountry);

    if (municipality) {
        parts.push(municipality);
    }

    if (isoRegion) {
        const regionParts = isoRegion.split('-');
        const region = regionParts[regionParts.length - 1] || isoRegion;
        if (region && !parts.includes(region)) {
            parts.push(region);
        }
    }

    if (isoCountry && isoCountry !== 'US') {
        parts.push(isoCountry);
    }

    return parts.join(', ');
}

function getAirportLookupLabel(airport) {
    const ident = airport && airport.ident ? airport.ident : '';
    const name = airport && airport.name ? airport.name : 'Unnamed airport';
    return `${ident} - ${name}`;
}

function getAirportRecordFromValue(value) {
    const trimmed = normalizeMissionText(value);
    if (!trimmed) return null;

    const exactCode = airportData[normalizeIcao(trimmed)];
    if (exactCode) return exactCode;

    const exactMatches = getAirportExactMatches(trimmed);
    if (exactMatches.length === 1) {
        return exactMatches[0];
    }

    return null;
}

function getAirportExactMatches(value) {
    const key = normalizeAirportExactKey(value);
    if (!key) return [];
    return airportExactLookupIndex.get(key) || [];
}

function getAirportLookupSuggestions(value, limit = 6) {
    const query = normalizeAirportSearchText(value);
    if (!query || query.length < 2 || airportLookupEntries.length === 0) return [];

    const exactMatches = getAirportExactMatches(value);
    if (exactMatches.length > 1) {
        return exactMatches.slice(0, limit);
    }

    const queryTokens = query.split(' ').filter(Boolean);
    const ranked = [];

    airportLookupEntries.forEach(airport => {
        const searchText = airport.searchText || '';
        if (!searchText) return;

        let score = null;

        if (searchText === query) {
            score = 1;
        } else if (searchText.startsWith(query)) {
            score = 2;
        } else if (searchText.includes(query)) {
            score = 3 + (searchText.indexOf(query) / 1000);
        } else {
            const hitCount = queryTokens.reduce((count, token) => count + (searchText.includes(token) ? 1 : 0), 0);
            if (hitCount === 0) return;
            score = 4 + ((queryTokens.length - hitCount) / queryTokens.length);
        }

        ranked.push({ airport, score });
    });

    ranked.sort((left, right) => (
        left.score - right.score
        || (left.airport.name || '').localeCompare(right.airport.name || '')
        || (left.airport.ident || '').localeCompare(right.airport.ident || '')
    ));

    return ranked.slice(0, limit).map(item => item.airport);
}

function getAirportFieldResolution(value) {
    const trimmed = normalizeMissionText(value);
    if (!trimmed) {
        return { status: 'empty', value: '' };
    }

    if (isValidAirportIcaoCode(trimmed)) {
        return {
            status: 'code',
            value: normalizeIcao(trimmed)
        };
    }

    const suggestions = getAirportLookupSuggestions(trimmed);
    if (suggestions.length > 0) {
        return {
            status: 'suggestions',
            suggestions
        };
    }

    return {
        status: 'unresolved',
        value: trimmed
    };
}

function resolveAirportCode(value) {
    const code = normalizeIcao(value);
    return /^[A-Z0-9]{4}$/.test(code) ? code : '';
}

function isValidAirportIcaoCode(value) {
    const code = normalizeIcao(value);
    return /^[A-Z0-9]{4}$/.test(code) && Boolean(airportData[code]);
}

function formatAirportLookupOption(airport) {
    const main = getAirportLookupLabel(airport);
    const context = getAirportLookupContext(airport);
    return {
        main,
        context
    };
}

function getAirportLookupPanel(input) {
    if (!input) return null;
    const field = input.closest('.leg-airport-field');
    if (!field) return null;
    return field.querySelector('.airport-lookup-panel');
}

function getAirportNameDisplay(input) {
    if (!input) return null;
    const field = input.closest('.leg-airport-field');
    if (!field) return null;
    return field.querySelector('.airport-name-display');
}

function normalizeAirportInputValue(input) {
    if (!input) return '';

    const rawValue = normalizeMissionText(input.value);
    if (!isValidAirportIcaoCode(rawValue)) {
        return rawValue;
    }

    const normalized = normalizeIcao(rawValue);
    if (input.value !== normalized) {
        const selectionStart = input.selectionStart;
        const selectionEnd = input.selectionEnd;
        const selectionDirection = input.selectionDirection;

        input.value = normalized;

        if (typeof selectionStart === 'number' && typeof selectionEnd === 'number') {
            try {
                input.setSelectionRange(selectionStart, selectionEnd, selectionDirection || 'none');
            } catch {
                // Ignore selection restore failures for unsupported browsers.
            }
        }
    }

    return normalized;
}

function updateAirportNameDisplay(input) {
    const display = getAirportNameDisplay(input);
    if (!display) return;

    const rawValue = normalizeAirportInputValue(input);
    const airport = isValidAirportIcaoCode(rawValue) ? airportData[normalizeIcao(rawValue)] : null;
    const hasAirportName = Boolean(airport && airport.name);

    if (hasAirportName) {
        display.textContent = airport.name;
        display.hidden = false;
    } else {
        display.textContent = '';
        display.hidden = true;
    }

    const row = input.closest('.leg-row');
    if (row) {
        const hasVisibleAirportName = Boolean(row.querySelector('.airport-name-display:not([hidden])'));
        row.classList.toggle('airport-name-visible', hasVisibleAirportName);
    }
}

function getAirportLookupSuggestionsForInput(input) {
    return input && Array.isArray(input._airportLookupSuggestions) ? input._airportLookupSuggestions : [];
}

function getAirportLookupActiveIndex(input, suggestions = getAirportLookupSuggestionsForInput(input)) {
    if (!Array.isArray(suggestions) || suggestions.length === 0) return -1;

    const rawIndex = Number.isInteger(input && input._airportLookupActiveIndex)
        ? input._airportLookupActiveIndex
        : 0;

    if (rawIndex < 0) return 0;
    if (rawIndex >= suggestions.length) return suggestions.length - 1;
    return rawIndex;
}

function setAirportLookupActiveIndex(input, index) {
    if (!input) return;

    const suggestions = getAirportLookupSuggestionsForInput(input);
    if (suggestions.length === 0) {
        input._airportLookupActiveIndex = -1;
        return;
    }

    const nextIndex = ((index % suggestions.length) + suggestions.length) % suggestions.length;
    input._airportLookupActiveIndex = nextIndex;

    const panel = getAirportLookupPanel(input);
    if (panel && !panel.hidden) {
        renderAirportLookupPanel(input, suggestions);
    }
}

function moveAirportLookupSelection(input, delta) {
    const suggestions = getAirportLookupSuggestionsForInput(input);
    if (suggestions.length === 0) return;

    const currentIndex = getAirportLookupActiveIndex(input, suggestions);
    const nextIndex = currentIndex < 0 ? 0 : currentIndex + delta;
    setAirportLookupActiveIndex(input, nextIndex);
}

function hideAirportLookupPanel(input) {
    const panel = getAirportLookupPanel(input);
    if (!panel) return;

    panel.hidden = true;
    panel.innerHTML = '';
    if (input) {
        if (input._airportLookupTimer) {
            clearTimeout(input._airportLookupTimer);
            input._airportLookupTimer = null;
        }
        input._airportLookupRequestToken = (input._airportLookupRequestToken || 0) + 1;
        input._airportLookupSuggestions = [];
        input._airportLookupActiveIndex = -1;
        input.setAttribute('aria-expanded', 'false');
        input.removeAttribute('aria-activedescendant');
    }
}

function selectAirportLookupOption(input, airport) {
    if (!input || !airport) return;

    input.value = airport.ident || '';
    hideAirportLookupPanel(input);
    updateAirportNameDisplay(input);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus({ preventScroll: true });
}

function renderAirportLookupPanel(input, suggestions) {
    const panel = getAirportLookupPanel(input);
    if (!panel) return;

    const activeIndex = getAirportLookupActiveIndex(input, suggestions);
    input._airportLookupSuggestions = suggestions;
    input._airportLookupActiveIndex = activeIndex;
    input.setAttribute('aria-expanded', 'true');

    if (!suggestions.length) {
        panel.innerHTML = '<div class="airport-lookup-empty">No airport matches found.</div>';
        panel.hidden = false;
        return;
    }

    panel.innerHTML = suggestions.map((airport, index) => {
        const option = formatAirportLookupOption(airport);
        const contextHtml = option.context
            ? `<span class="airport-lookup-option-meta">${escapeHtml(option.context)}</span>`
            : '';
        return `
            <button type="button" class="airport-lookup-option${index === activeIndex ? ' is-active' : ''}" role="option" aria-selected="${index === activeIndex ? 'true' : 'false'}" data-airport-index="${index}">
                <span class="airport-lookup-option-main">${escapeHtml(option.main)}</span>
                ${contextHtml}
            </button>
        `;
    }).join('');

    panel.hidden = false;

    panel.querySelectorAll('.airport-lookup-option').forEach(button => {
        const index = Number(button.getAttribute('data-airport-index'));
        const airport = suggestions[index];
        if (!airport) return;

        button.addEventListener('pointerdown', event => {
            event.preventDefault();
            event.stopPropagation();
            selectAirportLookupOption(input, airport);
        });

        button.addEventListener('keydown', event => {
            if (event.key !== 'Enter' && event.key !== ' ') return;
            event.preventDefault();
            event.stopPropagation();
            selectAirportLookupOption(input, airport);
        });
    });

    const activeButton = panel.querySelector('.airport-lookup-option.is-active');
    if (activeButton && typeof activeButton.scrollIntoView === 'function') {
        activeButton.scrollIntoView({ block: 'nearest' });
    }
}

async function updateAirportLookupPanel(input, requestToken = 0) {
    const rawValue = normalizeMissionText(input.value);
    const panel = getAirportLookupPanel(input);
    if (!panel) return;

    if (!rawValue) {
        hideAirportLookupPanel(input);
        updateAirportNameDisplay(input);
        return;
    }

    if (isValidAirportIcaoCode(rawValue)) {
        hideAirportLookupPanel(input);
        updateAirportNameDisplay(input);
        return;
    }

    if (airportLookupEntries.length === 0) {
        panel.innerHTML = '<div class="airport-lookup-empty">Loading airport data...</div>';
        panel.hidden = false;
        input.setAttribute('aria-expanded', 'true');
        input._airportLookupSuggestions = [];

        try {
            await loadAirportData();
        } catch {
            // Fall through and let the unresolved state handle the empty cache.
        }

        if (requestToken && input._airportLookupRequestToken !== requestToken) return;
        if (normalizeMissionText(input.value) !== rawValue) return;

        if (airportLookupEntries.length === 0) {
            panel.innerHTML = '<div class="airport-lookup-empty">Airport lookup unavailable.</div>';
            panel.hidden = false;
            input.setAttribute('aria-expanded', 'true');
            updateAirportNameDisplay(input);
            return;
        }
    }

    if (isValidAirportIcaoCode(rawValue)) {
        hideAirportLookupPanel(input);
        updateAirportNameDisplay(input);
        return;
    }

    const suggestions = getAirportLookupSuggestions(rawValue);
    renderAirportLookupPanel(input, suggestions);
    updateAirportNameDisplay(input);
}

function scheduleAirportLookupPanelUpdate(input) {
    if (!input) return;

    if (input._airportLookupTimer) {
        clearTimeout(input._airportLookupTimer);
    }

    input._airportLookupRequestToken = (input._airportLookupRequestToken || 0) + 1;
    input._airportLookupActiveIndex = 0;
    const requestToken = input._airportLookupRequestToken;
    input._airportLookupTimer = window.setTimeout(() => {
        input._airportLookupTimer = null;
        void updateAirportLookupPanel(input, requestToken);
    }, 120);
}

function attachAirportLookupBehavior(input) {
    if (!input || input._airportLookupAttached) return;

    input._airportLookupAttached = true;
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('spellcheck', 'false');
    input.setAttribute('aria-autocomplete', 'list');
    input.setAttribute('aria-haspopup', 'listbox');
    input.setAttribute('aria-expanded', 'false');

    input.addEventListener('input', () => {
        updateAirportNameDisplay(input);
        scheduleAirportLookupPanelUpdate(input);
    });

    input.addEventListener('focus', () => {
        updateAirportNameDisplay(input);
        scheduleAirportLookupPanelUpdate(input);
    });

    input.addEventListener('blur', () => {
        window.setTimeout(() => hideAirportLookupPanel(input), 0);
    });

    input.addEventListener('keydown', event => {
        if (event.key === 'Escape') {
            hideAirportLookupPanel(input);
            return;
        }

        if (event.key === 'ArrowDown' || event.key === 'ArrowUp' || event.key === 'Enter') {
            const suggestions = getAirportLookupSuggestionsForInput(input);
            const panel = getAirportLookupPanel(input);

            if (!suggestions.length || !panel || panel.hidden) {
                if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                    scheduleAirportLookupPanelUpdate(input);
                    event.preventDefault();
                }
                return;
            }

            event.preventDefault();

            if (event.key === 'ArrowDown') {
                moveAirportLookupSelection(input, 1);
                return;
            }

            if (event.key === 'ArrowUp') {
                moveAirportLookupSelection(input, -1);
                return;
            }

            const activeIndex = getAirportLookupActiveIndex(input, suggestions);
            const airport = suggestions[activeIndex];
            if (airport) {
                selectAirportLookupOption(input, airport);
            }
        }
    });
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function getDateTimestamp(date) {
    return date instanceof Date && !Number.isNaN(date.getTime()) ? date.getTime() : 0;
}

function clearMissionCardFocusHighlight() {
    if (missionCardFocusTimeout !== null) {
        clearTimeout(missionCardFocusTimeout);
        missionCardFocusTimeout = null;
    }

    if (focusedMissionCardId === null) return;

    const focusedCard = document.getElementById(`card-${focusedMissionCardId}`);
    if (focusedCard) {
        focusedCard.classList.remove('focus-highlight');
    }

    focusedMissionCardId = null;
}

function formatTooltipDateTime(date) {
    return date instanceof Date && !Number.isNaN(date.getTime()) ? date.toLocaleString() : 'TBD';
}

function formatTooltipDateDDMMM(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return 'TBD';

    const day = String(date.getDate()).padStart(2, '0');
    const month = date.toLocaleDateString('en-US', { month: 'short' }).toUpperCase();
    return `${day}-${month}`;
}

function formatZuluTime(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';

    const hours = String(date.getUTCHours()).padStart(2, '0');
    const minutes = String(date.getUTCMinutes()).padStart(2, '0');
    return `${hours}:${minutes}Z`;
}

function getMilitaryTimezoneCharacter(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';

    const offsetHours = -(date.getTimezoneOffset() / 60);
    if (offsetHours === 0) return 'Z';
    if (!Number.isInteger(offsetHours) || Math.abs(offsetHours) > 12) return '?';

    const positiveLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M'];
    const negativeLetters = ['N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'];

    if (offsetHours > 0) {
        return positiveLetters[offsetHours - 1] || '';
    }

    return negativeLetters[Math.abs(offsetHours) - 1] || '';
}

function formatLocalMilitaryTime(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';

    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const zone = getMilitaryTimezoneCharacter(date);
    return `${hours}:${minutes}${zone}`;
}

function getAirportDisplayName(code) {
    const normalizedCode = normalizeIcao(code);
    const airport = /^[A-Z0-9]{4}$/.test(normalizedCode) ? airportData[normalizedCode] : null;
    if (airport && airport.name) return airport.name;
    return normalizeMissionText(code) || 'TBD';
}

function getTooltipLegText(leg, index) {
    const takeoffIcao = normalizeMissionText(leg.takeoffIcao, true) || 'TBD';
    const landIcao = normalizeMissionText(leg.landIcao, true) || 'TBD';
    const takeoffDate = formatTooltipDateDDMMM(leg.takeoffTime);
    const landDate = formatTooltipDateDDMMM(leg.landTime);
    return `Leg ${index + 1}: ${takeoffIcao} (${takeoffDate}) → ${landIcao} (${landDate})`;
}

function getTooltipMeasureContext() {
    if (tooltipMeasureFont && tooltipMeasureCtx) return tooltipMeasureCtx;

    const probe = document.createElement('span');
    probe.className = 'tooltip-leg-line';
    probe.style.position = 'absolute';
    probe.style.visibility = 'hidden';
    probe.style.whiteSpace = 'nowrap';
    probe.textContent = 'Leg 1: KJFK (01-JAN) → KLAX (02-JAN)';
    document.body.appendChild(probe);

    const computedStyle = getComputedStyle(probe);
    tooltipMeasureFont = computedStyle.font || `${computedStyle.fontWeight} ${computedStyle.fontSize} ${computedStyle.fontFamily}`;
    probe.remove();

    if (!tooltipMeasureCtx) {
        tooltipMeasureCtx = document.createElement('canvas').getContext('2d');
    }

    tooltipMeasureCtx.font = tooltipMeasureFont;
    return tooltipMeasureCtx;
}

function measureTooltipTextWidth(text) {
    const ctx = getTooltipMeasureContext();
    return ctx.measureText(text).width;
}

function getRouteLabelMeasureContext() {
    if (routeLabelMeasureFont && routeLabelMeasureCtx) return routeLabelMeasureCtx;

    const probe = document.createElement('span');
    probe.className = 'route-icao-label';
    probe.style.position = 'absolute';
    probe.style.visibility = 'hidden';
    probe.style.whiteSpace = 'nowrap';
    probe.style.left = '-9999px';
    probe.style.top = '-9999px';
    probe.textContent = 'KJFK';
    document.body.appendChild(probe);

    const computedStyle = getComputedStyle(probe);
    routeLabelMeasureFont = computedStyle.font || `${computedStyle.fontWeight} ${computedStyle.fontSize} ${computedStyle.fontFamily}`;
    probe.remove();

    if (!routeLabelMeasureCtx) {
        routeLabelMeasureCtx = document.createElement('canvas').getContext('2d');
    }

    routeLabelMeasureCtx.font = routeLabelMeasureFont;
    return routeLabelMeasureCtx;
}

function measureRouteLabelTextWidth(text) {
    const ctx = getRouteLabelMeasureContext();
    return ctx.measureText(text).width;
}

function getRouteIcaoLabelSize(code) {
    const text = normalizeMissionText(code) || '';
    const width = Math.max(ROUTE_ICAO_LABEL_MIN_WIDTH_PX, Math.ceil(measureRouteLabelTextWidth(text)) + (ROUTE_ICAO_LABEL_HORIZONTAL_PADDING_PX * 2));

    return {
        width,
        height: ROUTE_ICAO_LABEL_HEIGHT_PX
    };
}

function getMissionFlightHours(mission) {
    if (!mission || !Array.isArray(mission.legs) || mission.legs.length === 0) return 0;

    return mission.legs.reduce((total, leg) => {
        const takeoffTime = leg && leg.takeoffTime instanceof Date ? leg.takeoffTime.getTime() : Number.NaN;
        const landTime = leg && leg.landTime instanceof Date ? leg.landTime.getTime() : Number.NaN;

        if (Number.isNaN(takeoffTime) || Number.isNaN(landTime) || landTime <= takeoffTime) {
            return total;
        }

        return total + ((landTime - takeoffTime) / MS_PER_HOUR);
    }, 0);
}

function formatMissionFlightHours(mission) {
    return getMissionFlightHours(mission).toFixed(1);
}

function setHoverTooltipWidth(mission) {
    const legs = Array.isArray(mission.legs) ? [...mission.legs].sort((a, b) => getDateTimestamp(a.takeoffTime) - getDateTimestamp(b.takeoffTime)) : [];
    const lines = legs.length > 0
        ? legs.map((leg, index) => getTooltipLegText(leg, index))
        : ['No route legs available.'];

    const missionLine = `Mission: ${mission.missionNum || 'TBD'} | Tail: ${mission.tailNum || 'TBD'} | Flight Hours: ${formatMissionFlightHours(mission)}`;
    lines.unshift(missionLine);

    let contentWidth = 0;
    lines.forEach(line => {
        contentWidth = Math.max(contentWidth, measureTooltipTextWidth(line));
    });

    const horizontalPadding = 24;
    const borderAllowance = 2;
    const targetWidth = Math.ceil(contentWidth + horizontalPadding + borderAllowance);
    const maxWidth = Math.max(280, document.documentElement.clientWidth - 24);

    tooltip.style.width = `${Math.min(targetWidth, maxWidth)}px`;
}

function ensureCurrentTimeBubbleStrip() {
    const timelineWrapper = document.querySelector('.timeline-wrapper');
    if (!timelineWrapper) return null;

    if (currentTimeBubbleStrip && document.body.contains(currentTimeBubbleStrip)) {
        return currentTimeBubbleStrip;
    }

    currentTimeBubbleStrip = document.getElementById('current-time-bubble-strip');
    if (!currentTimeBubbleStrip) {
        currentTimeBubbleStrip = document.createElement('div');
        currentTimeBubbleStrip.id = 'current-time-bubble-strip';
        timelineWrapper.insertAdjacentElement('afterend', currentTimeBubbleStrip);
    }

    currentTimeBubbleStrip.removeAttribute('aria-hidden');
    currentTimeBubbleStrip.style.display = 'block';
    currentTimeBubbleStrip.style.visibility = 'visible';
    currentTimeBubbleStrip.style.position = 'relative';
    currentTimeBubbleStrip.style.height = `${CURRENT_TIME_BUBBLE_HEIGHT_PX}px`;
    currentTimeBubbleStrip.style.overflow = 'visible';
    currentTimeBubbleStrip.style.pointerEvents = 'none';
    return currentTimeBubbleStrip;
}

function applyCurrentTimeBubbleStyles(bubble) {
    bubble.style.position = 'absolute';
    bubble.style.left = '0';
    bubble.style.bottom = '0';
    bubble.style.transform = 'translateX(-50%)';
    bubble.style.display = 'inline-flex';
    bubble.style.flexDirection = 'column';
    bubble.style.alignItems = 'center';
    bubble.style.justifyContent = 'center';
    bubble.style.boxSizing = 'border-box';
    bubble.style.minHeight = `${CURRENT_TIME_BUBBLE_HEIGHT_PX}px`;
    bubble.style.padding = '4px 8px 5px';
    bubble.style.border = '1px solid #007bff';
    bubble.style.borderRadius = '999px';
    bubble.style.background = 'rgba(255,255,255,0.98)';
    bubble.style.color = '#007bff';
    bubble.style.fontSize = '0.7rem';
    bubble.style.fontWeight = '800';
    bubble.style.lineHeight = '1.05';
    bubble.style.whiteSpace = 'nowrap';
    bubble.style.boxShadow = '0 1px 3px rgba(0,0,0,0.15)';
    bubble.style.zIndex = '31';
    bubble.style.pointerEvents = 'auto';
    bubble.style.textAlign = 'center';
    bubble.style.visibility = 'visible';
    bubble.style.cursor = 'pointer';
}

function snapToThreeDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    const end = new Date(start);
    end.setDate(end.getDate() + 2);
    end.setHours(23, 59, 59, 999);

    viewStart = start.getTime();
    viewEnd = end.getTime();

    renderTimeline();
}

function syncCurrentTimeIndicator() {
    const existingIndicator = gridContainer.querySelector('.current-time-line');
    const now = new Date();
    const nowTime = now.getTime();
    syncMissionCardTemporalState(nowTime);
    syncMissionTimelineTemporalState(nowTime);
    const duration = viewEnd - viewStart;
    const strip = ensureCurrentTimeBubbleStrip();

    if (!(duration > 0) || nowTime < viewStart || nowTime > viewEnd) {
        if (existingIndicator) existingIndicator.remove();
        if (strip) strip.hidden = true;
        return;
    }

    const nowPct = ((nowTime - viewStart) / duration) * 100;
    let timeLine = existingIndicator;

    if (!timeLine) {
        timeLine = document.createElement('div');
        timeLine.className = 'current-time-line';
        timeLine.setAttribute('aria-hidden', 'true');

        gridContainer.appendChild(timeLine);
    }

    timeLine.style.left = `${nowPct}%`;
    timeLine.style.bottom = '0';

    if (strip) {
        strip.hidden = false;
        strip.style.display = 'block';
        strip.style.visibility = 'visible';
        const sidebar = document.querySelector('.timeline-sidebar');
        const sidebarWidth = sidebar ? sidebar.getBoundingClientRect().width : 0;
        const viewportWidth = viewport ? viewport.getBoundingClientRect().width : 0;

        strip.style.marginLeft = `${sidebarWidth}px`;
        strip.style.width = `${viewportWidth}px`;
        strip.style.height = `${CURRENT_TIME_BUBBLE_HEIGHT_PX}px`;

        let bubble = strip.querySelector('.current-time-bubble');
        if (!bubble) {
            bubble = document.createElement('div');
            bubble.className = 'current-time-bubble';
            bubble.setAttribute('role', 'button');
            bubble.setAttribute('tabindex', '0');
            bubble.setAttribute('aria-label', 'Snap timeline to 3-day outlook');
            bubble.addEventListener('click', snapToThreeDayOutlook);
            bubble.addEventListener('keydown', event => {
                if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    snapToThreeDayOutlook();
                }
            });
            strip.appendChild(bubble);
        }

        applyCurrentTimeBubbleStyles(bubble);
        bubble.style.left = `${nowPct}%`;

        let zuluLine = bubble.querySelector('.current-time-bubble-utc');
        let localLine = bubble.querySelector('.current-time-bubble-local');

        if (!zuluLine) {
            zuluLine = document.createElement('span');
            zuluLine.className = 'current-time-bubble-line current-time-bubble-utc';
            bubble.appendChild(zuluLine);
        }

        if (!localLine) {
            localLine = document.createElement('span');
            localLine.className = 'current-time-bubble-line current-time-bubble-local';
            bubble.appendChild(localLine);
        }

        zuluLine.textContent = formatZuluTime(now);
        localLine.textContent = formatLocalMilitaryTime(now);
    }
}

function parseCsvLine(line) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];

        if (inQuotes) {
            if (char === '"') {
                if (line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                current += char;
            }
            continue;
        }

        if (char === '"') {
            inQuotes = true;
        } else if (char === ',') {
            values.push(current);
            current = '';
        } else {
            current += char;
        }
    }

    values.push(current);
    return values;
}

function loadAirportData() {
    if (airportDataPromise) return airportDataPromise;

    airportDataPromise = fetch('https://raw.githubusercontent.com/davidmegginson/ourairports-data/main/airports.csv')
        .then(response => {
            if (!response.ok) {
                throw new Error(`Airport CSV request failed with status ${response.status}`);
            }
            return response.text();
        })
        .then(text => {
            const lines = text.replace(/\r/g, '').split('\n').filter(line => line.trim() !== '');
            if (lines.length === 0) {
                throw new Error('Airport CSV was empty.');
            }

            const headers = parseCsvLine(lines[0]).map(header => header.trim().toLowerCase());
            const identIndex = headers.indexOf('ident');
            const latIndex = headers.indexOf('latitude_deg');
            const lonIndex = headers.indexOf('longitude_deg');
            const nameIndex = headers.indexOf('name');
            const municipalityIndex = headers.indexOf('municipality');
            const isoRegionIndex = headers.indexOf('iso_region');
            const isoCountryIndex = headers.indexOf('iso_country');
            const iataCodeIndex = headers.indexOf('iata_code');
            const localCodeIndex = headers.indexOf('local_code');
            const typeIndex = headers.indexOf('type');

            if (identIndex === -1 || latIndex === -1 || lonIndex === -1) {
                throw new Error('Airport CSV missing expected columns.');
            }

            const nextAirportData = {};
            const nextAirportLookupEntries = [];
            const nextAirportExactLookupIndex = new Map();

            for (let i = 1; i < lines.length; i++) {
                const cols = parseCsvLine(lines[i]);
                const ident = normalizeIcao(cols[identIndex]);
                const lat = parseFloat(cols[latIndex]);
                const lon = parseFloat(cols[lonIndex]);
                const name = nameIndex !== -1 ? (cols[nameIndex] || '').trim() : '';
                const municipality = municipalityIndex !== -1 ? (cols[municipalityIndex] || '').trim() : '';
                const isoRegion = isoRegionIndex !== -1 ? (cols[isoRegionIndex] || '').trim() : '';
                const isoCountry = isoCountryIndex !== -1 ? (cols[isoCountryIndex] || '').trim() : '';
                const iataCode = iataCodeIndex !== -1 ? normalizeIcao(cols[iataCodeIndex]) : '';
                const localCode = localCodeIndex !== -1 ? normalizeIcao(cols[localCodeIndex]) : '';
                const type = typeIndex !== -1 ? (cols[typeIndex] || '').trim() : '';

                if (ident && Number.isFinite(lat) && Number.isFinite(lon)) {
                    const airport = {
                        ident,
                        lat,
                        lon,
                        name,
                        municipality,
                        isoRegion,
                        isoCountry,
                        iataCode,
                        localCode,
                        type
                    };

                    airport.nameExactKey = normalizeAirportExactKey(name);
                    airport.searchText = normalizeAirportSearchText([
                        ident,
                        name,
                        municipality,
                        isoRegion,
                        isoCountry,
                        iataCode,
                        localCode,
                        type
                    ].filter(Boolean).join(' '));
                    airport.lookupLabel = getAirportLookupLabel(airport);
                    airport.lookupContext = getAirportLookupContext(airport);

                    nextAirportData[ident] = airport;
                    nextAirportLookupEntries.push(airport);

                    addAirportExactLookupKey(nextAirportExactLookupIndex, ident, airport);
                    addAirportExactLookupKey(nextAirportExactLookupIndex, name, airport);
                    addAirportExactLookupKey(nextAirportExactLookupIndex, iataCode, airport);
                    addAirportExactLookupKey(nextAirportExactLookupIndex, localCode, airport);
                }
            }

            airportData = nextAirportData;
            airportLookupEntries = nextAirportLookupEntries;
            airportExactLookupIndex = nextAirportExactLookupIndex;
            return airportData;
        })
        .catch(error => {
            airportDataPromise = null;
            throw error;
        });

    return airportDataPromise;
}

