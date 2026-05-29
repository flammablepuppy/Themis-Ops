﻿let missions = [];
let viewStart, viewEnd; 
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MS_PER_HOUR = 60 * 60 * 1000;

const viewport = document.getElementById('timeline-viewport');
const gridContainer = document.getElementById('timeline-grid');
const monthsContainer = document.getElementById('timeline-months');
const yearsContainer = document.getElementById('timeline-years');
const missionsContainer = document.getElementById('timeline-missions');
const labelsContainer = document.getElementById('timeline-labels');
const tooltip = document.getElementById('hover-tooltip');
const modal = document.getElementById('mission-modal');
const missionForm = document.getElementById('mission-form');
const legsWrapper = document.getElementById('legs-wrapper');
const missionDefaultsForm = document.getElementById('mission-defaults-form');
const missionHomeFieldInput = document.getElementById('defaultHomeField');
const missionDataPathInput = document.getElementById('defaultMissionDataPath');
const DEFAULTS_STORAGE_KEY = 'themis-ops-mission-defaults';
const MISSIONS_STORAGE_KEY = 'themis-ops-missions';
const MISSION_DATA_HANDLE_DB_NAME = 'themis-ops-mission-data';
const MISSION_DATA_HANDLE_STORE_NAME = 'handles';
const MISSION_DATA_HANDLE_KEY = 'mission-data-handle';
let missionDefaults = createEmptyMissionDefaults();
const tabButtons = document.querySelectorAll('.app-tab');
let airportData = {};
let airportDataPromise = null;
let hoverRouteMap = null;
let hoverTooltipToken = 0;
let hoverTooltipPosition = { x: 0, y: 0 };
let activeTooltipMissionId = null;
let tooltipMeasureCtx = null;
let tooltipMeasureFont = '';
const ROUTE_AIRPORT_DOT_RADIUS_PX = 4;
const ROUTE_AIRPORT_DOT_WEIGHT_PX = 1;
const ROUTE_AIRPORT_DOT_EDGE_PX = ROUTE_AIRPORT_DOT_RADIUS_PX + (ROUTE_AIRPORT_DOT_WEIGHT_PX / 2);
let editingMissionId = null;
let editingMissionBackup = null;
let missionDataFileHandle = null;
let missionDataFileHandleLoadPromise = null;
let missionDataFileWriteQueue = Promise.resolve();
let missionDataFileHandleDisabled = false;
let missionDataFileHandleLoadToken = 0;

function createEmptyMissionDefaults() {
    return {
        homeField: '',
        missionDataPath: ''
    };
}

function loadMissionDefaults() {
    try {
        const raw = localStorage.getItem(DEFAULTS_STORAGE_KEY);
        if (!raw) return createEmptyMissionDefaults();
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== 'object') return createEmptyMissionDefaults();
        const homeField = typeof parsed.homeField === 'string' ? parsed.homeField : '';
        const missionDataPath = typeof parsed.missionDataPath === 'string' ? parsed.missionDataPath : '';
        const legacyHomeField = typeof parsed.tailNum === 'string' ? parsed.tailNum : '';
        return {
            homeField: (homeField || legacyHomeField).trim(),
            missionDataPath: missionDataPath.trim()
        };
    } catch {
        return createEmptyMissionDefaults();
    }
}

function persistMissionDefaults() {
    try {
        localStorage.setItem(DEFAULTS_STORAGE_KEY, JSON.stringify(missionDefaults));
    } catch {
        // Ignore storage failures in file:// or restricted browser contexts.
    }
}

function syncMissionDefaultsForm() {
    if (missionHomeFieldInput) missionHomeFieldInput.value = missionDefaults.homeField ?? '';
    if (missionDataPathInput) missionDataPathInput.value = missionDefaults.missionDataPath ?? '';
}

function readMissionDefaultsForm() {
    return {
        homeField: missionHomeFieldInput ? missionHomeFieldInput.value.trim() : '',
        missionDataPath: missionDataPathInput ? missionDataPathInput.value.trim() : ''
    };
}

function getMissionDataFileName(pathValue) {
    const trimmed = (pathValue || '').trim();
    if (!trimmed) return 'themis_operations_missions.json';

    const parts = trimmed.split(/[\\/]/).filter(Boolean);
    return parts.length > 0 ? parts[parts.length - 1] : 'themis_operations_missions.json';
}

function supportsMissionDataFileAccess() {
    return typeof window.showSaveFilePicker === 'function' && window.isSecureContext;
}

function openMissionDataHandleDatabase() {
    if (typeof indexedDB === 'undefined') return Promise.resolve(null);

    return new Promise((resolve, reject) => {
        const request = indexedDB.open(MISSION_DATA_HANDLE_DB_NAME, 1);

        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(MISSION_DATA_HANDLE_STORE_NAME)) {
                db.createObjectStore(MISSION_DATA_HANDLE_STORE_NAME);
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function readStoredMissionDataHandle() {
    const db = await openMissionDataHandleDatabase();
    if (!db) return null;

    return new Promise((resolve, reject) => {
        const tx = db.transaction(MISSION_DATA_HANDLE_STORE_NAME, 'readonly');
        const store = tx.objectStore(MISSION_DATA_HANDLE_STORE_NAME);
        const request = store.get(MISSION_DATA_HANDLE_KEY);

        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
        tx.oncomplete = () => db.close();
        tx.onerror = () => db.close();
    });
}

async function storeMissionDataHandle(handle) {
    const db = await openMissionDataHandleDatabase();
    if (!db) return;

    return new Promise((resolve, reject) => {
        const tx = db.transaction(MISSION_DATA_HANDLE_STORE_NAME, 'readwrite');
        const store = tx.objectStore(MISSION_DATA_HANDLE_STORE_NAME);
        store.put(handle, MISSION_DATA_HANDLE_KEY);
        tx.oncomplete = () => {
            db.close();
            resolve();
        };
        tx.onerror = () => {
            db.close();
            reject(tx.error);
        };
    });
}

async function clearStoredMissionDataHandle() {
    missionDataFileHandleDisabled = true;
    missionDataFileHandleLoadToken += 1;
    missionDataFileHandle = null;
    missionDataFileHandleLoadPromise = null;

    const db = await openMissionDataHandleDatabase();
    if (!db) return;

    return new Promise((resolve, reject) => {
        const tx = db.transaction(MISSION_DATA_HANDLE_STORE_NAME, 'readwrite');
        const store = tx.objectStore(MISSION_DATA_HANDLE_STORE_NAME);
        store.delete(MISSION_DATA_HANDLE_KEY);
        tx.oncomplete = () => {
            db.close();
            resolve();
        };
        tx.onerror = () => {
            db.close();
            reject(tx.error);
        };
    });
}

async function loadMissionDataHandle() {
    if (missionDataFileHandle) return missionDataFileHandle;
    if (missionDataFileHandleDisabled) return null;
    if (missionDataFileHandleLoadPromise) return missionDataFileHandleLoadPromise;

    const loadToken = missionDataFileHandleLoadToken;
    missionDataFileHandleLoadPromise = (async () => {
        try {
            if (!supportsMissionDataFileAccess()) {
                missionDataFileHandleDisabled = true;
                return null;
            }
            const storedHandle = await readStoredMissionDataHandle();
            if (loadToken !== missionDataFileHandleLoadToken || missionDataFileHandleDisabled) return null;
            missionDataFileHandle = storedHandle || null;
            if (!missionDataFileHandle) {
                missionDataFileHandleDisabled = true;
                return null;
            }

            if (missionDataFileHandle && missionDataPathInput && !missionDefaults.missionDataPath) {
                missionDefaults.missionDataPath = missionDataFileHandle.name || '';
                syncMissionDefaultsForm();
                persistMissionDefaults();
            }

            return missionDataFileHandle;
        } catch {
            return null;
        } finally {
            missionDataFileHandleLoadPromise = null;
        }
    })();

    return missionDataFileHandleLoadPromise;
}

async function requestMissionDataHandle() {
    if (!supportsMissionDataFileAccess()) {
        alert('Disk file selection requires a secure browser context such as localhost or HTTPS.');
        return null;
    }

    const suggestedName = getMissionDataFileName(missionDataPathInput ? missionDataPathInput.value : '');

    try {
        const handle = await window.showSaveFilePicker({
            suggestedName,
            types: [
                {
                    description: 'Mission JSON',
                    accept: { 'application/json': ['.json'] }
                }
            ]
        });

        missionDataFileHandle = handle;
        missionDataFileHandleDisabled = false;
        missionDataFileHandleLoadToken += 1;

        try {
            await storeMissionDataHandle(handle);
        } catch (error) {
            console.warn('Failed to store mission data file handle.', error);
        }

        missionDefaults.missionDataPath = handle.name || suggestedName;
        syncMissionDefaultsForm();
        persistMissionDefaults();
        await queueMissionDataDiskWrite();

        return handle;
    } catch (error) {
        if (error && error.name === 'AbortError') return null;
        console.warn('Mission data file selection failed.', error);
        return null;
    }
}

async function writeMissionDataToDisk() {
    const handle = await loadMissionDataHandle();
    if (!handle) return false;
    const writeToken = missionDataFileHandleLoadToken;

    try {
        if (missionDataFileHandleDisabled || writeToken !== missionDataFileHandleLoadToken) return false;

        if (typeof handle.queryPermission === 'function') {
            const permission = await handle.queryPermission({ mode: 'readwrite' });
            if (permission !== 'granted' && typeof handle.requestPermission === 'function') {
                const requested = await handle.requestPermission({ mode: 'readwrite' });
                if (requested !== 'granted') return false;
            }
        }

        if (missionDataFileHandleDisabled || writeToken !== missionDataFileHandleLoadToken) return false;

        const writable = await handle.createWritable();
        if (missionDataFileHandleDisabled || writeToken !== missionDataFileHandleLoadToken) {
            if (typeof writable.abort === 'function') {
                try {
                    await writable.abort();
                } catch {
                    // Ignore abort failures after invalidation.
                }
            }
            return false;
        }

        await writable.write(JSON.stringify(missions, null, 2));
        await writable.close();
        return true;
    } catch (error) {
        console.warn('Failed to write mission data to disk.', error);
        return false;
    }
}

function queueMissionDataDiskWrite() {
    missionDataFileWriteQueue = missionDataFileWriteQueue
        .then(() => writeMissionDataToDisk())
        .catch(error => {
            console.warn('Mission data disk write queue failed.', error);
        });

    return missionDataFileWriteQueue;
}

function normalizeMissionText(value, uppercase = false) {
    const text = value == null ? '' : String(value).trim();
    return uppercase ? text.toUpperCase() : text;
}

function parseMissionDate(value) {
    const date = value instanceof Date ? new Date(value) : new Date(value);
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

function normalizeMissionRecord(mission, index = 0) {
    if (!mission || typeof mission !== 'object') return null;

    const normalized = { ...mission };
    normalized.id = mission.id != null ? mission.id : (Date.now() + index);
    normalized.missionNum = normalizeMissionText(mission.missionNum, true);
    normalized.tailNum = normalizeMissionText(mission.tailNum);
    normalized.pilot = normalizeMissionText(mission.pilot);
    normalized.copilot = normalizeMissionText(mission.copilot);
    normalized.crewChief = normalizeMissionText(mission.crewChief);
    normalized.loadmaster = normalizeMissionText(mission.loadmaster);
    normalized.liftCustomer = normalizeMissionText(mission.liftCustomer);
    normalized.liftPax = normalizeMissionText(mission.liftPax);
    normalized.liftCargo = normalizeMissionText(mission.liftCargo);
    normalized.liftHazmat = normalizeMissionText(mission.liftHazmat);
    normalized.legs = Array.isArray(mission.legs)
        ? mission.legs.map(normalizeMissionLeg).sort((a, b) => getDateTimestamp(a.takeoffTime) - getDateTimestamp(b.takeoffTime))
        : [];

    return normalized;
}

function loadMissions() {
    try {
        const raw = localStorage.getItem(MISSIONS_STORAGE_KEY);
        if (raw === null) return null;

        try {
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) return [];
            return parsed.map((mission, index) => normalizeMissionRecord(mission, index)).filter(Boolean);
        } catch {
            return [];
        }
    } catch {
        return null;
    }
}

function persistMissions() {
    try {
        localStorage.setItem(MISSIONS_STORAGE_KEY, JSON.stringify(missions));
        void queueMissionDataDiskWrite();
    } catch {
        // Ignore storage failures in file:// or restricted browser contexts.
    }
}

function normalizeIcao(code) {
    return (code || '').trim().toUpperCase();
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

function formatTooltipDateTime(date) {
    return date instanceof Date && !Number.isNaN(date.getTime()) ? date.toLocaleString() : 'TBD';
}

function getAirportDisplayName(code) {
    const airport = airportData[normalizeIcao(code)];
    if (airport && airport.name) return airport.name;
    return normalizeIcao(code) || 'TBD';
}

function getTooltipLegText(leg, index) {
    const takeoffName = getAirportDisplayName(leg.takeoffIcao);
    const landName = getAirportDisplayName(leg.landIcao);
    const takeoffTime = formatTooltipDateTime(leg.takeoffTime);
    const landTime = formatTooltipDateTime(leg.landTime);
    return `Leg ${index + 1}: ${takeoffName} (${takeoffTime}) → ${landName} (${landTime})`;
}

function getTooltipMeasureContext() {
    if (tooltipMeasureFont && tooltipMeasureCtx) return tooltipMeasureCtx;

    const probe = document.createElement('span');
    probe.className = 'tooltip-leg-line';
    probe.style.position = 'absolute';
    probe.style.visibility = 'hidden';
    probe.style.whiteSpace = 'nowrap';
    probe.textContent = 'Leg 1: Sample Airport (00/00/0000, 12:00:00 AM) -> Sample Airport (00/00/0000, 12:00:00 AM)';
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
    return getMissionFlightHours(mission).toFixed(2);
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

            if (identIndex === -1 || latIndex === -1 || lonIndex === -1) {
                throw new Error('Airport CSV missing expected columns.');
            }

            const nextAirportData = {};

            for (let i = 1; i < lines.length; i++) {
                const cols = parseCsvLine(lines[i]);
                const ident = normalizeIcao(cols[identIndex]);
                const lat = parseFloat(cols[latIndex]);
                const lon = parseFloat(cols[lonIndex]);
                const name = nameIndex !== -1 ? (cols[nameIndex] || '').trim() : '';

                if (ident && Number.isFinite(lat) && Number.isFinite(lon)) {
                    nextAirportData[ident] = { lat, lon, name };
                }
            }

            airportData = nextAirportData;
            return airportData;
        })
        .catch(error => {
            airportDataPromise = null;
            throw error;
        });

    return airportDataPromise;
}

function clearHoverRouteMap() {
    hoverTooltipToken += 1;
    activeTooltipMissionId = null;
    if (hoverRouteMap) {
        hoverRouteMap.remove();
        hoverRouteMap = null;
    }
}

function positionHoverTooltip(pageX, pageY) {
    const padding = 12;
    const offset = 16;
    const tooltipWidth = tooltip.offsetWidth;
    const tooltipHeight = tooltip.offsetHeight;
    const viewportWidth = document.documentElement.clientWidth;
    const viewportHeight = document.documentElement.clientHeight;
    const minLeft = window.scrollX + padding;
    const minTop = window.scrollY + padding;
    const maxLeft = Math.max(minLeft, window.scrollX + viewportWidth - tooltipWidth - padding);
    const maxTop = Math.max(minTop, window.scrollY + viewportHeight - tooltipHeight - padding);

    let left = pageX + offset;
    let top = pageY + offset;

    if (left > maxLeft) left = pageX - tooltipWidth - offset;
    if (top > maxTop) top = pageY - tooltipHeight - offset;

    left = Math.min(Math.max(left, minLeft), maxLeft);
    top = Math.min(Math.max(top, minTop), maxTop);

    tooltip.style.left = `${left}px`;
    tooltip.style.top = `${top}px`;
}

function getMissionRouteCodes(mission) {
    const legs = Array.isArray(mission.legs) ? [...mission.legs].sort((a, b) => getDateTimestamp(a.takeoffTime) - getDateTimestamp(b.takeoffTime)) : [];
    const codes = [];

    legs.forEach(leg => {
        const takeoff = normalizeIcao(leg.takeoffIcao);
        const landing = normalizeIcao(leg.landIcao);

        if (takeoff && codes[codes.length - 1] !== takeoff) codes.push(takeoff);
        if (landing && codes[codes.length - 1] !== landing) codes.push(landing);
    });

    return codes;
}

function getRouteSegmentPlacement(map, start, end) {
    const startPoint = map.latLngToLayerPoint(L.latLng(start[0], start[1]));
    const endPoint = map.latLngToLayerPoint(L.latLng(end[0], end[1]));
    const delta = endPoint.subtract(startPoint);
    const length = Math.hypot(delta.x, delta.y);

    if (!length) {
        return {
            latLng: map.layerPointToLatLng(endPoint),
            angle: 0
        };
    }

    const direction = delta.divideBy(length);
    const tipPoint = endPoint.subtract(direction.multiplyBy(ROUTE_AIRPORT_DOT_EDGE_PX));

    return {
        latLng: map.layerPointToLatLng(tipPoint),
        angle: Math.atan2(delta.y, delta.x) * 180 / Math.PI
    };
}

function getRouteLabelPlacement(map, routePoints, index, offsetPx = 16) {
    const currentPoint = map.latLngToLayerPoint(L.latLng(routePoints[index][0], routePoints[index][1]));
    const prevRaw = routePoints[index - 1];
    const nextRaw = routePoints[index + 1];
    const routeCenterPoint = routePoints.reduce((acc, point) => {
        const layerPoint = map.latLngToLayerPoint(L.latLng(point[0], point[1]));
        acc.x += layerPoint.x;
        acc.y += layerPoint.y;
        return acc;
    }, { x: 0, y: 0 });

    routeCenterPoint.x /= routePoints.length;
    routeCenterPoint.y /= routePoints.length;

    let direction = null;

    if (prevRaw && nextRaw) {
        const prevPoint = map.latLngToLayerPoint(L.latLng(prevRaw[0], prevRaw[1]));
        const nextPoint = map.latLngToLayerPoint(L.latLng(nextRaw[0], nextRaw[1]));
        const incoming = currentPoint.subtract(prevPoint);
        const outgoing = nextPoint.subtract(currentPoint);
        const incomingLength = Math.hypot(incoming.x, incoming.y);
        const outgoingLength = Math.hypot(outgoing.x, outgoing.y);
        const incomingUnit = incomingLength ? incoming.divideBy(incomingLength) : null;
        const outgoingUnit = outgoingLength ? outgoing.divideBy(outgoingLength) : null;

        if (incomingUnit && outgoingUnit) {
            direction = incomingUnit.add(outgoingUnit);
            if (!direction.x && !direction.y) {
                direction = outgoingUnit;
            }
        } else {
            direction = outgoingUnit || incomingUnit;
        }
    } else if (nextRaw) {
        const nextPoint = map.latLngToLayerPoint(L.latLng(nextRaw[0], nextRaw[1]));
        direction = nextPoint.subtract(currentPoint);
    } else if (prevRaw) {
        const prevPoint = map.latLngToLayerPoint(L.latLng(prevRaw[0], prevRaw[1]));
        direction = currentPoint.subtract(prevPoint);
    }

    if (!direction || (!direction.x && !direction.y)) {
        return map.layerPointToLatLng(currentPoint);
    }

    const length = Math.hypot(direction.x, direction.y) || 1;
    const unit = direction.divideBy(length);
    const normal = L.point(-unit.y, unit.x);
    const outward = currentPoint.add(normal.multiplyBy(offsetPx));
    const inward = currentPoint.subtract(normal.multiplyBy(offsetPx));

    const outwardDistance = Math.hypot(outward.x - routeCenterPoint.x, outward.y - routeCenterPoint.y);
    const inwardDistance = Math.hypot(inward.x - routeCenterPoint.x, inward.y - routeCenterPoint.y);
    const placementPoint = outwardDistance >= inwardDistance ? outward : inward;

    return map.layerPointToLatLng(placementPoint);
}

function createRouteDirectionIcon(angle) {
    return L.divIcon({
        className: 'route-direction-marker',
        html: `
            <svg class="route-direction-arrow" viewBox="0 0 20 16" aria-hidden="true">
                <g transform="rotate(${angle} 18 8)">
                    <path class="route-direction-arrow-shaft" d="M2 8h7" />
                    <path class="route-direction-arrow-head" d="M9 1l9 7-9 7z" />
                </g>
            </svg>
        `,
        iconSize: [20, 16],
        iconAnchor: [18, 8]
    });
}

function createAirportCircleMarker(latlng) {
    return L.circleMarker(latlng, {
        radius: ROUTE_AIRPORT_DOT_RADIUS_PX,
        color: '#000',
        weight: ROUTE_AIRPORT_DOT_WEIGHT_PX,
        fillColor: '#000',
        fillOpacity: 1,
        opacity: 1,
        interactive: false,
        bubblingMouseEvents: false
    });
}

function buildMissionTooltipHTML(mission) {
    const legs = Array.isArray(mission.legs) ? [...mission.legs].sort((a, b) => getDateTimestamp(a.takeoffTime) - getDateTimestamp(b.takeoffTime)) : [];
    const routeLines = legs.length > 0
        ? legs.map((leg, index) => {
            const takeoffName = escapeHtml(getAirportDisplayName(leg.takeoffIcao));
            const landName = escapeHtml(getAirportDisplayName(leg.landIcao));
            const takeoffTime = escapeHtml(formatTooltipDateTime(leg.takeoffTime));
            const landTime = escapeHtml(formatTooltipDateTime(leg.landTime));
            return `<div class="tooltip-leg-line">Leg ${index + 1}: ${takeoffName} (${takeoffTime}) → ${landName} (${landTime})</div>`;
        }).join('')
        : '<div>No route legs available.</div>';
    const flightHours = escapeHtml(formatMissionFlightHours(mission));

    return `
        <div class="tooltip-summary">
            <div><strong>Mission:</strong> ${escapeHtml(mission.missionNum || 'TBD')} | <strong>Tail:</strong> ${escapeHtml(mission.tailNum || 'TBD')} | <strong>Flight Hours:</strong> ${flightHours}</div>
            <hr class="tooltip-separator">
            <div class="tooltip-itinerary">${routeLines}</div>
        </div>
        <div class="tooltip-route-section">
            <div id="hover-route-map-status" class="tooltip-route-status">Loading route map...</div>
            <div id="hover-route-map" class="tooltip-route-map"></div>
        </div>
    `;
}

async function renderHoverRouteMap(mission, token) {
    let routeMapEl = tooltip.querySelector('#hover-route-map');
    let statusEl = tooltip.querySelector('#hover-route-map-status');
    if (!routeMapEl || !statusEl) return;

    if (typeof L === 'undefined') {
        statusEl.textContent = 'Route map unavailable.';
        routeMapEl.style.visibility = 'hidden';
        return;
    }

    const routeCodes = getMissionRouteCodes(mission);
    if (routeCodes.length < 2) {
        statusEl.textContent = 'Route map unavailable for this mission.';
        routeMapEl.style.visibility = 'hidden';
        return;
    }

    routeMapEl.style.visibility = 'hidden';

    try {
        await loadAirportData();
        if (token !== hoverTooltipToken) return;

        tooltip.innerHTML = buildMissionTooltipHTML(mission);
        setHoverTooltipWidth(mission);

        routeMapEl = tooltip.querySelector('#hover-route-map');
        statusEl = tooltip.querySelector('#hover-route-map-status');
        if (!routeMapEl || !statusEl) return;

        await new Promise(resolve => requestAnimationFrame(resolve));
        if (token !== hoverTooltipToken) return;

        const routePointEntries = [];
        const markerCodes = new Set();

        routeCodes.forEach(code => {
            const airport = airportData[code];
            if (!airport) return;
            routePointEntries.push({
                code,
                point: [airport.lat, airport.lon]
            });
        });

        const routePoints = routePointEntries.map(entry => entry.point);

        if (routePoints.length < 2) {
            statusEl.textContent = 'Route map unavailable for this mission.';
            routeMapEl.style.visibility = 'hidden';
            return;
        }

        if (hoverRouteMap) {
            hoverRouteMap.remove();
            hoverRouteMap = null;
        }

        routeMapEl.innerHTML = '';
        hoverRouteMap = L.map(routeMapEl, {
            dragging: false,
            zoomControl: false,
            scrollWheelZoom: false,
            doubleClickZoom: false,
            boxZoom: false,
            touchZoom: false,
            keyboard: false,
            tap: false,
            attributionControl: false
        });

        L.tileLayer('https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png', {
            subdomains: 'abcd',
            maxZoom: 19
        }).addTo(hoverRouteMap);

        const routeLayer = L.layerGroup().addTo(hoverRouteMap);

        L.polyline(routePoints, { color: '#d71920', weight: 4 }).addTo(routeLayer);
        hoverRouteMap.fitBounds(routePoints, { padding: [18, 18] });

        routePointEntries.forEach(entry => {
            createAirportCircleMarker(entry.point).addTo(routeLayer);
        });

        routePointEntries.forEach((entry, index) => {
            const { code } = entry;
            if (markerCodes.has(code)) return;
            markerCodes.add(code);
            const labelLatLng = getRouteLabelPlacement(hoverRouteMap, routePoints, index);
            L.marker(labelLatLng, {
                interactive: false,
                keyboard: false,
                icon: L.divIcon({
                    className: 'route-icao-label',
                    html: escapeHtml(code),
                    iconSize: null,
                    iconAnchor: [0, 0]
                })
            }).addTo(routeLayer);
        });

        for (let i = 0; i < routePoints.length - 1; i++) {
            const start = routePoints[i];
            const end = routePoints[i + 1];
            if (start[0] === end[0] && start[1] === end[1]) continue;

            const placement = getRouteSegmentPlacement(hoverRouteMap, start, end);
            L.marker(placement.latLng, {
                interactive: false,
                keyboard: false,
                icon: createRouteDirectionIcon(placement.angle)
            }).addTo(routeLayer);
        }

        routeMapEl.style.visibility = 'visible';
        statusEl.textContent = '';

        hoverRouteMap.invalidateSize();
        positionHoverTooltip(hoverTooltipPosition.x, hoverTooltipPosition.y);
    } catch (error) {
        if (token !== hoverTooltipToken) return;
        console.error('Failed to render route map', error);
        statusEl.textContent = 'Route map unavailable.';
        routeMapEl.style.visibility = 'hidden';
        if (hoverRouteMap) {
            hoverRouteMap.remove();
            hoverRouteMap = null;
        }
    }
}

function showMissionTooltip(mission, event) {
    clearHoverRouteMap();
    activeTooltipMissionId = mission.id;
    hoverTooltipPosition = { x: event.pageX, y: event.pageY };
    const token = hoverTooltipToken;

    tooltip.innerHTML = buildMissionTooltipHTML(mission);
    tooltip.style.display = 'flex';
    setHoverTooltipWidth(mission);
    positionHoverTooltip(event.pageX, event.pageY);
    renderHoverRouteMap(mission, token);
}

function cloneMissionRecord(mission) {
    return {
        ...mission,
        legs: Array.isArray(mission.legs)
            ? mission.legs.map(leg => ({
                ...leg,
                takeoffTime: new Date(leg.takeoffTime),
                landTime: new Date(leg.landTime)
            }))
            : []
    };
}

function getEditingMissionRecord() {
    if (editingMissionId == null) return null;
    return missions.find(mission => String(mission.id) === String(editingMissionId)) || null;
}

function applyMissionDataInPlace(target, missionData) {
    const id = target.id;
    Object.keys(target).forEach(key => {
        if (key !== 'id') delete target[key];
    });
    Object.assign(target, { id, ...missionData });
}

function buildMissionDataFromForm() {
    const legNodes = document.querySelectorAll('.leg-container');
    let legs = [];
    let timeValid = true;

    legNodes.forEach(node => {
        const tkIcao = node.querySelector('.leg-tk-icao').value.trim().toUpperCase();
        const ldIcao = node.querySelector('.leg-ld-icao').value.trim().toUpperCase();
        const tkTime = parseDateTimeLocalValue(node.querySelector('.leg-tk-time').value);
        const ldTime = parseDateTimeLocalValue(node.querySelector('.leg-ld-time').value);

        if (!tkTime || !ldTime || ldTime <= tkTime) timeValid = false;
        legs.push({
            takeoffIcao: tkIcao,
            takeoffTime: tkTime || new Date(NaN),
            landIcao: ldIcao,
            landTime: ldTime || new Date(NaN)
        });
    });

    return {
        hasLegs: legNodes.length > 0,
        timeValid,
        missionData: {
            missionNum: document.getElementById('missionNum').value.toUpperCase(),
            tailNum: document.getElementById('tailNum').value,
            pilot: document.getElementById('pilot').value,
            copilot: document.getElementById('copilot').value,
            crewChief: document.getElementById('crewChief').value,
            loadmaster: document.getElementById('loadmaster').value,
            liftCustomer: document.getElementById('liftCustomer').value,
            liftPax: document.getElementById('liftPax').value,
            liftCargo: document.getElementById('liftCargo').value,
            liftHazmat: document.getElementById('liftHazmat').value,
            legs
        }
    };
}

function refreshTooltipForMission(missionId) {
    if (tooltip.style.display === 'none') return;
    if (activeTooltipMissionId == null || String(activeTooltipMissionId) !== String(missionId)) return;

    const mission = missions.find(item => String(item.id) === String(missionId));
    if (!mission) return;

    showMissionTooltip(mission, {
        pageX: hoverTooltipPosition.x,
        pageY: hoverTooltipPosition.y
    });
}

function syncEditingMissionPreview() {
    const mission = getEditingMissionRecord();
    if (!mission) return;

    const snapshot = buildMissionDataFromForm();
    if (!snapshot.timeValid) return;

    applyMissionDataInPlace(mission, snapshot.missionData);
    refreshTooltipForMission(mission.id);
}

function getMissionHomeField() {
    return (missionDefaults.homeField || '').trim().toUpperCase();
}

function createDateTimeLocal(daysAhead, hour, minute) {
    const date = new Date();
    date.setDate(date.getDate() + daysAhead);
    date.setHours(hour, minute, 0, 0);
    return date;
}

function parseDateTimeLocalValue(value) {
    if (!value) return null;
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
}

function getInitialLegDefaults() {
    const homeField = getMissionHomeField();
    const takeoffTime = createDateTimeLocal(1, 9, 0);
    const landTime = createDateTimeLocal(1, 12, 30);

    return {
        takeoffIcao: homeField,
        takeoffTime,
        landIcao: homeField,
        landTime
    };
}

function getNextLegDefaults() {
    const legNodes = legsWrapper.querySelectorAll('.leg-container');
    if (legNodes.length === 0) return getInitialLegDefaults();

    const previousLeg = legNodes[legNodes.length - 1];
    const previousLandIcaoInput = previousLeg.querySelector('.leg-ld-icao');
    const previousLandTimeInput = previousLeg.querySelector('.leg-ld-time');
    const previousLandTime = parseDateTimeLocalValue(previousLandTimeInput ? previousLandTimeInput.value : '');

    const takeoffIcao = previousLandIcaoInput ? previousLandIcaoInput.value.trim().toUpperCase() : '';
    const takeoffTime = previousLandTime ? new Date(previousLandTime) : createDateTimeLocal(1, 9, 0);
    if (previousLandTime) takeoffTime.setMinutes(takeoffTime.getMinutes() + 60);

    const landTime = new Date(takeoffTime);
    landTime.setMinutes(landTime.getMinutes() + 180);

    return {
        takeoffIcao,
        takeoffTime,
        landIcao: '',
        landTime
    };
}

function activateTab(panelId) {
    tabButtons.forEach(button => {
        const isActive = button.dataset.tab === panelId;
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-selected', isActive ? 'true' : 'false');
    });

    document.querySelectorAll('.tab-panel').forEach(panel => {
        const isActive = panel.id === panelId;
        panel.hidden = !isActive;
    });

    if (panelId === 'timeline-panel') {
        renderTimeline();
    }
}

function init() {
    const now = new Date();
    let currentFY = now.getFullYear();
    if (now.getMonth() >= 9) currentFY += 1; 
    document.getElementById('fy-input').value = currentFY;

    missionDefaults = loadMissionDefaults();
    syncMissionDefaultsForm();

    const loadedMissions = loadMissions();
    if (loadedMissions !== null) {
        missions = loadedMissions;
    } else {
        //addDummyData();
    }

    snapToSevenDayOutlook();
    renderMissionCards();
    void loadMissionDataHandle().then(handle => {
        if (handle) void queueMissionDataDiskWrite();
    });
    void loadAirportData().catch(() => {});
}

function snapToSevenDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    const end = new Date(start);
    end.setDate(end.getDate() + 8);
    end.setHours(23, 59, 59, 999);

    viewStart = start.getTime();
    viewEnd = end.getTime();

    renderTimeline();
}

function snapToFourteenDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    const end = new Date(start);
    end.setDate(end.getDate() + 15);
    end.setHours(23, 59, 59, 999);

    viewStart = start.getTime();
    viewEnd = end.getTime();

    renderTimeline();
}

function snapToThirtyDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate()-1);
    const end = new Date(start);
    end.setDate(end.getDate() + 31);
    end.setHours(23, 59, 59, 999);

    viewStart = start.getTime();
    viewEnd = end.getTime();

    renderTimeline();
}

function snapToFY(fy) {
    viewStart = new Date(fy - 1, 9, 1).getTime(); 
    viewEnd = new Date(fy, 8, 30, 23, 59, 59).getTime(); 
    renderTimeline();
}

document.getElementById('btn-snap-7day').addEventListener('click', snapToSevenDayOutlook);
document.getElementById('btn-snap-14day').addEventListener('click', snapToFourteenDayOutlook);
document.getElementById('btn-snap-30day').addEventListener('click', snapToThirtyDayOutlook);
document.getElementById('btn-snap-fy').addEventListener('click', () => snapToFY(parseInt(document.getElementById('fy-input').value, 10)));
tabButtons.forEach(button => {
    button.addEventListener('click', () => activateTab(button.dataset.tab));
});

// Zoom & Pan functionality
viewport.addEventListener('wheel', (e) => {
    e.preventDefault();
    const rect = viewport.getBoundingClientRect();
    const mousePct = (e.clientX - rect.left) / rect.width; 
    const zoomFactor = e.deltaY > 0 ? 1.1 : 0.9; 
    const currentDuration = viewEnd - viewStart;
    let newDuration = currentDuration * zoomFactor;
    
    const minZoom = 3 * MS_PER_DAY;
    const maxZoom = 2 * 365 * MS_PER_DAY;
    if (newDuration < minZoom) newDuration = minZoom;
    if (newDuration > maxZoom) newDuration = maxZoom;

    const anchorTime = viewStart + mousePct * currentDuration;
    viewStart = anchorTime - (mousePct * newDuration);
    viewEnd = anchorTime + ((1 - mousePct) * newDuration);

    renderTimeline();
});

let isDragging = false, lastX = 0;
viewport.addEventListener('mousedown', (e) => { isDragging = true; lastX = e.clientX; });
window.addEventListener('mouseup', () => isDragging = false);
window.addEventListener('mouseleave', () => isDragging = false);
window.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    const dx = e.clientX - lastX;
    lastX = e.clientX;
    const msPerPixel = (viewEnd - viewStart) / viewport.getBoundingClientRect().width;
    const timeShift = dx * msPerPixel;
    viewStart -= timeShift; viewEnd -= timeShift;
    renderTimeline();
});

function getMissionTimes(mission) {
    if(!mission.legs || mission.legs.length === 0) return { start: 0, end: 0 };
    const firstLeg = mission.legs[0] || {};
    const lastLeg = mission.legs[mission.legs.length - 1] || {};
    const start = getDateTimestamp(firstLeg.takeoffTime);
    const end = getDateTimestamp(lastLeg.landTime);
    return { start, end };
}

function zoomToMonth(dateObj) {
    const start = new Date(dateObj.getFullYear(), dateObj.getMonth(), 1);
    const end = new Date(dateObj.getFullYear(), dateObj.getMonth() + 1, 0, 23, 59, 59);
    viewStart = start.getTime();
    viewEnd = end.getTime();
    renderTimeline();
}

function zoomToYear(dateObj) {
    const start = new Date(dateObj.getFullYear(), 0, 1);
    const end = new Date(dateObj.getFullYear(), 11, 31, 23, 59, 59);
    viewStart = start.getTime();
    viewEnd = end.getTime();
    renderTimeline();
}

function renderTimeline() {
    gridContainer.innerHTML = '';
    monthsContainer.innerHTML = '';
    yearsContainer.innerHTML = '';
    missionsContainer.innerHTML = '';
    labelsContainer.innerHTML = '';

    // Add Day Divider Line
    const dayDivider = document.createElement('div');
    dayDivider.className = 'day-divider-line';
    gridContainer.appendChild(dayDivider);

    const startDate = new Date(viewStart);
    const duration = viewEnd - viewStart;
    const rect = viewport.getBoundingClientRect();
    const msPerPx = duration / rect.width;
    const pxPerDay = MS_PER_DAY / msPerPx;
    
    // Highlight current day
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const tomorrowStart = todayStart + MS_PER_DAY;

    if (todayStart < viewEnd && tomorrowStart > viewStart) {
        const leftPct = ((todayStart - viewStart) / duration) * 100;
        const widthPct = (MS_PER_DAY / duration) * 100;

        const todayHighlight = document.createElement('div');
        todayHighlight.className = 'today-highlight';
        todayHighlight.style.left = `${leftPct}%`;
        todayHighlight.style.width = `${widthPct}%`;

        gridContainer.appendChild(todayHighlight);
    }

    // Current time indicator
    const nowTime = now.getTime();
    if (nowTime >= viewStart && nowTime <= viewEnd) {
        const nowPct = ((nowTime - viewStart) / duration) * 100;

        const timeLine = document.createElement('div');
        timeLine.className = 'current-time-line';
        timeLine.style.left = `${nowPct}%`;

        gridContainer.appendChild(timeLine);
    }

    // Determine zoom granularity
    const totalDaysVisible = duration / MS_PER_DAY;

    // Adaptive divider rendering
    if (totalDaysVisible <= 365) {
        const visibleDays = [];
        let iterDay = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
        while (iterDay <= new Date(viewEnd)) {
            visibleDays.push(new Date(iterDay));
            iterDay.setDate(iterDay.getDate() + 1);
        }

        const dayLabelMode = getDayLabelMode(visibleDays, viewStart, duration, rect.width, pxPerDay);

        if (dayLabelMode !== 'none') {
            visibleDays.forEach(dayDate => {
                const time = dayDate.getTime();
                const leftPx = ((time - viewStart) / duration) * rect.width;
                const leftPct = (leftPx / rect.width) * 100;
                const widthPct = (MS_PER_DAY / duration) * 100;

                if (isWeekend(dayDate)) {
                    const weekendHighlight = document.createElement('div');
                    weekendHighlight.className = 'weekend-highlight';
                    weekendHighlight.style.left = `${leftPct}%`;
                    weekendHighlight.style.width = `${widthPct}%`;
                    gridContainer.appendChild(weekendHighlight);
                }

                const marker = document.createElement('div');
                marker.className = 'grid-marker day';
                marker.style.left = `${leftPct}%`;

                const labelText = getDayLabelText(dayDate, dayLabelMode);
                if (labelText && canFitDayLabel(labelText, leftPx, pxPerDay, rect.width)) {
                    marker.innerHTML = `<div class="grid-label day-label">${labelText}</div>`;
                }
                gridContainer.appendChild(marker);
            });
        } else {
            let iterWeek = new Date(startDate);
            iterWeek.setDate(iterWeek.getDate() - iterWeek.getDay());

            while (iterWeek <= new Date(viewEnd)) {
                const time = iterWeek.getTime();
                const leftPct = ((time - viewStart) / duration) * 100;

                const marker = document.createElement('div');
                marker.className = 'week-marker';
                marker.style.left = `${leftPct}%`;

                gridContainer.appendChild(marker);
                iterWeek.setDate(iterWeek.getDate() + 7);
            }
        }
    }

    // YEAR LABELS
    const visibleYears = [];
    let iterYear = new Date(startDate.getFullYear(), 0, 1);
    if (iterYear.getTime() > viewStart) {
        iterYear.setFullYear(iterYear.getFullYear() - 1);
    }
    while (iterYear.getTime() <= viewEnd + (366 * MS_PER_DAY)) {
        visibleYears.push(new Date(iterYear));
        iterYear.setFullYear(iterYear.getFullYear() + 1);
    }

    visibleYears.forEach((yearDate, idx) => {
        const yearStart = yearDate.getTime();
        const nextYear = visibleYears[idx + 1];
        const nextYearStart = nextYear ? nextYear.getTime() : viewEnd;

        const leftPx = ((yearStart - viewStart) / duration) * rect.width;
        const nextLeftPx = ((nextYearStart - viewStart) / duration) * rect.width;

        const label = document.createElement('div');
        label.className = 'year-label';
        label.textContent = yearDate.getFullYear();

        const labelWidth = 80;
        const minLeft = 12;
        let labelLeft = Math.max(leftPx + 6, minLeft);

        if (nextYear) {
            labelLeft = Math.min(labelLeft, nextLeftPx - labelWidth);
        }
        label.style.left = `${labelLeft}px`;
        label.addEventListener('click', () => zoomToYear(yearDate));
        yearsContainer.appendChild(label);
    });

    // MONTH + FY MARKERS
    const visibleMonths = [];
    const stackMonthLabels = totalDaysVisible > 365;
    let iterMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    if (iterMonth.getTime() > viewStart) {
        iterMonth.setMonth(iterMonth.getMonth() - 1);
    }
    while (iterMonth.getTime() <= viewEnd + (32 * MS_PER_DAY)) {
        visibleMonths.push(new Date(iterMonth));
        iterMonth.setMonth(iterMonth.getMonth() + 1);
    }

    visibleMonths.forEach((monthDate, idx) => {
        const monthStart = monthDate.getTime();
        const nextMonth = visibleMonths[idx + 1];
        const nextMonthStart = nextMonth ? nextMonth.getTime() : viewEnd;

        const leftPx = ((monthStart - viewStart) / duration) * rect.width;
        const nextLeftPx = ((nextMonthStart - viewStart) / duration) * rect.width;

        const isFYBoundary = monthDate.getMonth() === 9;
        const marker = document.createElement('div');
        marker.className = isFYBoundary ? 'fy-marker' : 'month-marker';
        marker.style.left = `${(leftPx / rect.width) * 100}%`;
        gridContainer.appendChild(marker);

        const label = document.createElement('div');
        label.className = 'month-label';
        label.textContent = monthDate.toLocaleString('default', { month: 'long' });

        const labelWidth = 110;
        const minLeft = 12;
        let labelLeft = Math.max(leftPx + 6, minLeft);

        if (!stackMonthLabels && nextMonth) {
            labelLeft = Math.min(labelLeft, nextLeftPx - labelWidth);
        }
        label.style.left = `${labelLeft}px`;
        label.style.transform = 'none';
        label.addEventListener('click', () => zoomToMonth(monthDate));
        monthsContainer.appendChild(label);
    });

    // --- TAIL ROW ASSIGNMENT ---
    const uniqueTails = [...new Set(missions.map(m => m.tailNum ? m.tailNum.toUpperCase() : 'TBD'))].sort();
    const rowHeight = 46;
    const rowOffset = 24; // Start tail rows below the day-label divider.

    // Dynamically adjust viewport height to fit all tails
    const requiredHeight = 72 + rowOffset + (uniqueTails.length * rowHeight);
    viewport.style.height = `${requiredHeight}px`;

    // Render Tail Rows and Labels inside the fixed Sidebar
    uniqueTails.forEach((tail, index) => {
        const topPx = rowOffset + (index * rowHeight);

        // Divider line across the viewport grid
        const line = document.createElement('div');
        line.className = 'timeline-row-line';
        line.style.top = `${topPx + rowHeight}px`;
        missionsContainer.appendChild(line);

        // Fixed Tail Label in Sidebar (matches vertical offset perfectly)
        const label = document.createElement('div');
        label.className = 'tail-label';
        label.style.top = `${topPx + 8}px`; 
        label.textContent = tail;
        labelsContainer.appendChild(label);
    });

    // Render Missions into assigned Tail Rows
    missions.forEach(mission => {
        const times = getMissionTimes(mission);
        if (times.end < viewStart || times.start > viewEnd) return;

        const leftPct = ((times.start - viewStart) / duration) * 100;
        const widthPct = ((times.end - times.start) / duration) * 100;

        const bar = document.createElement('div');
        bar.className = 'timeline-bar';
        
        // Highlight missing crew or past missions
        const isCompleteCrew = mission.pilot && mission.copilot && mission.crewChief && mission.loadmaster;
        if (!isCompleteCrew) {
            bar.classList.add('unassigned-crew');
        }
        if (times.end < nowTime) {
            bar.classList.add('past');
        }

        bar.style.left = `${Math.max(0, leftPct)}%`;
        bar.style.width = leftPct < 0 ? `${widthPct + leftPct}%` : `${widthPct}%`;
        if(leftPct + widthPct > 100) bar.style.width = `${100 - leftPct}%`;
        
        // Assign top position based on matching tail row
        const tailStr = mission.tailNum ? mission.tailNum.toUpperCase() : 'TBD';
        const tailIndex = uniqueTails.indexOf(tailStr);
        bar.style.top = `${rowOffset + (tailIndex * rowHeight) + 8}px`; 
        
        const firstIcao = mission.legs[0].takeoffIcao;
        const lastIcao = mission.legs[mission.legs.length - 1].landIcao;
            bar.innerHTML = `<span>${mission.missionNum} (${firstIcao}&rarr;${lastIcao})</span>`;

        bar.addEventListener('mouseenter', (e) => {
            showMissionTooltip(mission, e);

            const card = document.getElementById(`card-${mission.id}`);
            if (card) {
                card.classList.add('highlight');
            }
        });

        bar.addEventListener('mousemove', (e) => {
            hoverTooltipPosition = { x: e.pageX, y: e.pageY };
            positionHoverTooltip(e.pageX, e.pageY);
        });
        bar.addEventListener('mouseleave', () => {
            tooltip.style.display = 'none';
            clearHoverRouteMap();
            const card = document.getElementById(`card-${mission.id}`);
            if(card) card.classList.remove('highlight');
        });

        bar.addEventListener('click', () => focusMissionCard(mission.id));
        missionsContainer.appendChild(bar);
    });
}

function renderMissionCards() {
    const list = document.getElementById('mission-list');
    list.innerHTML = '';
    const sorted = [...missions].sort((a, b) => getMissionTimes(b).start - getMissionTimes(a).start);

    sorted.forEach(mission => {
        const card = document.createElement('div');
        card.className = 'mission-card';
        card.id = `card-${mission.id}`;

        const hasAnyCrew = !!(mission.pilot || mission.copilot || mission.crewChief || mission.loadmaster);
        const isComplete = !!(mission.missionNum && mission.tailNum && mission.legs.length > 0 && 
                              mission.pilot && mission.copilot && mission.crewChief && mission.loadmaster);

        if (!hasAnyCrew) card.classList.add('no-crew');
        const checkMark = isComplete ? `<span class="status-check">&#10003;</span>` : '';

        let legsHTML = mission.legs.map((leg, idx) => `
                <li><strong>Leg ${idx+1}:</strong> ${leg.takeoffIcao} (${leg.takeoffTime.toLocaleString(undefined, {month:'short', day:'numeric', hour:'2-digit', minute:'2-digit'})}) &rarr;
                ${leg.landIcao} (${leg.landTime.toLocaleString(undefined, {month:'short', day:'numeric', hour:'2-digit', minute:'2-digit'})})</li>
        `).join('');

        card.innerHTML = `
            <div class="card-col col-mission">
                <div class="mission-header">
                    <h3>${checkMark} Mission: ${mission.missionNum}</h3>
                </div>
                <div class="mission-details">
                    <p class="mission-tail-line"><strong>Tail #:</strong> ${mission.tailNum || '<em>TBD</em>'}</p>
                    <ul class="itinerary-list">${legsHTML}</ul>
                    <span class="edit-hint">Click card anywhere to edit</span>
                </div>
            </div>
            
            <div class="card-col col-crew">
                <div class="crew-section">
                    <h4>Crew</h4>
                    <p><strong>Pilot:</strong> ${mission.pilot || '<em>TBD</em>'}</p>
                    <p><strong>Co-Pilot:</strong> ${mission.copilot || '<em>TBD</em>'}</p>
                    <p><strong>Crew Chief:</strong> ${mission.crewChief || '<em>TBD</em>'}</p>
                    <p><strong>Loadmaster(s):</strong> ${mission.loadmaster || '<em>TBD</em>'}</p>
                </div>
            </div>
            
            <div class="card-col col-lift">
                <div class="lift-section">
                    <h4>Lift</h4>
                    <p><strong>Customer:</strong> ${mission.liftCustomer || '0'}</p>
                    <p><strong>Pax:</strong> ${mission.liftPax || '0'}</p>
                    <p><strong>Cargo:</strong> ${mission.liftCargo || '<em>None</em>'}</p>
                    <p><strong>Hazmat:</strong> ${mission.liftHazmat || '<em>None</em>'}</p>
                </div>
            </div>
            
            <div class="card-col col-buttons">
                <button type="button" class="btn-locate">Locate</button>
                <button type="button" class="btn-delete">Delete</button>
            </div>
        `;
        
        card.querySelector('.btn-locate').addEventListener('click', (event) => locateMission(mission.id, event));
        card.querySelector('.btn-delete').addEventListener('click', (event) => deleteMission(mission.id, event));
        card.addEventListener('click', () => openEditModal(mission));
        list.appendChild(card);
    });
}

function locateMission(id, event) {
    event.stopPropagation();
    const mission = missions.find(m => m.id === id);
    if (!mission) return;
    
    const times = getMissionTimes(mission);
    if (times.start === 0) return;

    const missionStartDay = new Date(times.start);
    missionStartDay.setHours(0, 0, 0, 0);

    const missionEndFrame = new Date(missionStartDay.getTime());
    missionEndFrame.setDate(missionEndFrame.getDate() + 2);
    missionEndFrame.setHours(23, 59, 59, 999);

    viewStart = missionStartDay.getTime();
    viewEnd = missionEndFrame.getTime();

    renderTimeline();
    document.getElementById('timeline-viewport').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function focusMissionCard(id) {
    const card = document.getElementById(`card-${id}`);
    if (!card) return;

    card.classList.add('highlight');
    card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function deleteMission(id, event) {
    event.stopPropagation();
    if(confirm("Are you sure you want to delete this mission?")) {
        missions = missions.filter(m => m.id !== id);
        persistMissions();
        renderTimeline();
        renderMissionCards();
    }
}

document.getElementById('btn-import-csv').addEventListener('click', () => document.getElementById('csv-file-input').click());
document.getElementById('csv-file-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
        const text = event.target.result;
        const lines = text.split('\n').filter(l => l.trim() !== '');
        if(lines.length < 2) return alert("CSV is empty or missing data.");
        
        const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
        const missionMap = new Map();

        for(let i=1; i<lines.length; i++) {
            const cols = lines[i].split(',').map(c => c.trim().replace(/^"|"$/g, ''));
            let row = {};
            headers.forEach((h, idx) => row[h] = cols[idx]);
            
            const mNum = row.mission;
            if(!mNum) continue;

            if(!missionMap.has(mNum)) {
                missionMap.set(mNum, {
                    id: Date.now() + i,
                    missionNum: mNum.toUpperCase(),
                    tailNum: row.tail || '',
                    pilot: row.pilot || '',
                    copilot: row.copilot || '',
                    crewChief: row.crewchief || '',
                    loadmaster: row.loadmaster || '',
                    liftPax: row.pax || '',
                    liftCargo: row.cargo || '',
                    liftHazmat: row.hazmat || '',
                    legs: []
                });
            }
            
            if(row.takeoff_icao && row.takeoff_time && row.land_icao && row.land_time) {
                missionMap.get(mNum).legs.push({
                    takeoffIcao: row.takeoff_icao.toUpperCase(),
                    takeoffTime: new Date(row.takeoff_time),
                    landIcao: row.land_icao.toUpperCase(),
                    landTime: new Date(row.land_time)
                });
            }
        }
        
        missionMap.forEach(m => {
            if(m.legs.length > 0) {
                m.legs.sort((a,b) => a.takeoffTime - b.takeoffTime);
                const normalizedMission = normalizeMissionRecord(m);
                if (normalizedMission) missions.push(normalizedMission);
            }
        });
        
        persistMissions();
        renderTimeline();
        renderMissionCards();
        alert("CSV Imported Successfully!");
        e.target.value = ''; 
    };
    reader.readAsText(file);
});

document.getElementById('btn-export-csv').addEventListener('click', exportMissionsToCSV);
missionDefaultsForm.addEventListener('submit', function(e) {
    e.preventDefault();
    missionDefaults = readMissionDefaultsForm();
    persistMissionDefaults();
});

missionDefaultsForm.addEventListener('input', event => {
    if (event.target === missionDataPathInput && !missionDataFileHandleDisabled) {
        void clearStoredMissionDataHandle();
    }
    missionDefaults = readMissionDefaultsForm();
    persistMissionDefaults();
});

document.getElementById('btn-clear-defaults').addEventListener('click', () => {
    missionDefaults = createEmptyMissionDefaults();
    persistMissionDefaults();
    syncMissionDefaultsForm();
    void clearStoredMissionDataHandle();
});

document.getElementById('btn-choose-mission-data-file').addEventListener('click', () => {
    void requestMissionDataHandle();
});

function toDateTimeLocal(date) {
    const pad = n => n < 10 ? '0' + n : n;
    return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) + 'T' + pad(date.getHours()) + ':' + pad(date.getMinutes());
}

let dayLabelMeasureCtx = null;
let dayLabelMeasureFont = '';

function getDayLabelMeasureFont() {
    if (dayLabelMeasureFont) return dayLabelMeasureFont;

    const probe = document.createElement('span');
    probe.className = 'grid-label day-label';
    probe.style.position = 'absolute';
    probe.style.visibility = 'hidden';
    probe.style.whiteSpace = 'nowrap';
    probe.textContent = '00 - Wednesday';
    document.body.appendChild(probe);

    const computedStyle = getComputedStyle(probe);
    dayLabelMeasureFont = computedStyle.font || `${computedStyle.fontWeight} ${computedStyle.fontSize} ${computedStyle.fontFamily}`;
    probe.remove();

    return dayLabelMeasureFont;
}

function measureDayLabelWidth(text) {
    if (!dayLabelMeasureCtx) {
        dayLabelMeasureCtx = document.createElement('canvas').getContext('2d');
    }

    dayLabelMeasureCtx.font = getDayLabelMeasureFont();
    return dayLabelMeasureCtx.measureText(text).width;
}

function canFitDayLabel(text, leftPx, pxPerDay, viewportWidth) {
    if (leftPx < 0 || leftPx >= viewportWidth) return false;
    const paddingAllowance = 8;
    const buffer = 4;
    const availableWidth = Math.min(pxPerDay, viewportWidth - leftPx);
    return availableWidth >= (measureDayLabelWidth(text) + paddingAllowance + buffer);
}

function getDayLabelText(date, mode) {
    const dayNumber = String(date.getDate()).padStart(2, '0');

    if (mode === 'full') {
        const dayOfWeek = date.toLocaleDateString('en-US', { weekday: 'long' });
        return `${dayNumber} - ${dayOfWeek}`;
    }

    if (mode === 'short') {
        const dayAbbrev = date.toLocaleDateString('en-US', { weekday: 'short' });
        return `${dayNumber} - ${dayAbbrev}`;
    }

    if (mode === 'day') {
        return dayNumber;
    }

    return '';
}

function getDayLabelMode(dayDates, viewStart, duration, viewportWidth, pxPerDay) {
    const modes = ['full', 'short', 'day'];
    const fullVisibleDays = dayDates.filter(date => isFullDayVisible(date, viewStart, viewStart + duration));
    const testDays = fullVisibleDays.length > 0 ? fullVisibleDays : dayDates;

    for (const mode of modes) {
        const fitsEveryLabel = testDays.every(date => {
            const leftPx = ((date.getTime() - viewStart) / duration) * viewportWidth;
            const labelText = getDayLabelText(date, mode);
            return canFitDayLabel(labelText, leftPx, pxPerDay, viewportWidth);
        });

        if (fitsEveryLabel) return mode;
    }

    return 'none';
}

function isFullDayVisible(date, viewStart, viewEnd) {
    const dayStart = date.getTime();
    const dayEnd = dayStart + MS_PER_DAY - 1;
    return dayStart >= viewStart && dayEnd <= viewEnd;
}

function isWeekend(date) {
    const day = date.getDay();
    return day === 0 || day === 6;
}

function csvEscape(value) {
    const text = value == null ? '' : String(value);
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function exportMissionsToCSV() {
    const headers = ['mission', 'tail', 'pilot', 'copilot', 'crewchief', 'loadmaster', 'pax', 'cargo', 'hazmat', 'takeoff_icao', 'takeoff_time', 'land_icao', 'land_time'];
    const rows = [headers.join(',')];

    missions.forEach(mission => {
        if (!mission.legs || mission.legs.length === 0) return;

        mission.legs.forEach(leg => {
            rows.push([
                mission.missionNum || '',
                mission.tailNum || '',
                mission.pilot || '',
                mission.copilot || '',
                mission.crewChief || '',
                mission.loadmaster || '',
                mission.liftPax || '',
                mission.liftCargo || '',
                mission.liftHazmat || '',
                leg.takeoffIcao || '',
                toDateTimeLocal(leg.takeoffTime),
                leg.landIcao || '',
                toDateTimeLocal(leg.landTime)
            ].map(csvEscape).join(','));
        });
    });

    const blob = new Blob([rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'themis_operations_export.csv';
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function addLegRow(legData = null) {
    const div = document.createElement('div');
    div.className = 'leg-container';
    const defaultLegData = legData || getNextLegDefaults();
    div.innerHTML = `
        <div class="form-row leg-row">
            <div class="form-col leg-col-short"><label class="leg-label">Origin</label><input type="text" class="leg-tk-icao" required maxlength="4" value="${defaultLegData.takeoffIcao || ''}"></div>
            <div class="form-col leg-col-wide"><label class="leg-label">Takeoff Time</label><input type="datetime-local" class="leg-tk-time" required value="${toDateTimeLocal(defaultLegData.takeoffTime)}"></div>
            <div class="form-col leg-col-short"><label class="leg-label">Destination</label><input type="text" class="leg-ld-icao" required maxlength="4" value="${defaultLegData.landIcao || ''}"></div>
            <div class="form-col leg-col-wide"><label class="leg-label">Land Time</label><input type="datetime-local" class="leg-ld-time" required value="${toDateTimeLocal(defaultLegData.landTime)}"></div>
            <div><button type="button" class="btn btn-remove-leg leg-remove-btn" title="Remove Leg">&times;</button></div>
        </div>
    `;    
    div.querySelector('.btn-remove-leg').addEventListener('click', () => {
        div.remove();
        syncEditingMissionPreview();
    });
    legsWrapper.appendChild(div);
    if (!legData) syncEditingMissionPreview();
}

document.getElementById('btn-add-leg').addEventListener('click', () => addLegRow());
missionForm.addEventListener('input', event => {
    if (event.target.closest('.leg-container')) syncEditingMissionPreview();
});
missionForm.addEventListener('change', event => {
    if (event.target.closest('.leg-container')) syncEditingMissionPreview();
});

document.getElementById('btn-new-mission').addEventListener('click', () => {
    missionForm.reset();
    document.getElementById('editMissionId').value = "";
    legsWrapper.innerHTML = '';
    addLegRow();
    document.getElementById('modal-title').innerText = "New Mission";
    document.getElementById('btn-submit-form').innerText = "Create Mission";
    modal.style.display = "block";
});

function clearEditSession() {
    editingMissionId = null;
    editingMissionBackup = null;
}

function startEditSession(mission) {
    editingMissionId = mission.id;
    editingMissionBackup = cloneMissionRecord(mission);
}

function restoreEditSession() {
    if (editingMissionId == null || !editingMissionBackup) {
        clearEditSession();
        return;
    }

    const mission = getEditingMissionRecord();
    if (mission) {
        applyMissionDataInPlace(mission, cloneMissionRecord(editingMissionBackup));
    }

    const restoredMissionId = editingMissionId;
    clearEditSession();
    refreshTooltipForMission(restoredMissionId);
}

document.getElementById('close-modal').addEventListener('click', () => {
    restoreEditSession();
    modal.style.display = "none";
});

function openEditModal(mission) {
    startEditSession(mission);
    document.getElementById('editMissionId').value = mission.id;
    document.getElementById('missionNum').value = mission.missionNum;
    document.getElementById('tailNum').value = mission.tailNum;
    document.getElementById('pilot').value = mission.pilot;
    document.getElementById('copilot').value = mission.copilot;
    document.getElementById('crewChief').value = mission.crewChief;
    document.getElementById('loadmaster').value = mission.loadmaster;
    document.getElementById('liftCustomer').value = mission.liftCustomer || '';
    document.getElementById('liftPax').value = mission.liftPax || '';
    document.getElementById('liftCargo').value = mission.liftCargo || '';
    document.getElementById('liftHazmat').value = mission.liftHazmat || '';
    
    legsWrapper.innerHTML = '';
    mission.legs.forEach(leg => addLegRow(leg));
    
    document.getElementById('modal-title').innerText = "Edit Mission";
    document.getElementById('btn-submit-form').innerText = "Update Mission";
    modal.style.display = "block";
}

missionForm.addEventListener('submit', function(e) {
    e.preventDefault();
    const { hasLegs, timeValid, missionData } = buildMissionDataFromForm();
    if (!hasLegs) return alert("You must add at least one flight leg.");
    if (!timeValid) return alert("Landing time must be after takeoff time for all legs.");
    missionData.legs.sort((a, b) => a.takeoffTime - b.takeoffTime);

    const editId = document.getElementById('editMissionId').value;
    const committedMissionId = editId || null;
    if (editId) {
        const index = missions.findIndex(m => m.id == editId);
        if (index > -1) {
            const normalizedMission = normalizeMissionRecord({ id: missions[index].id, ...missionData });
            if (normalizedMission) applyMissionDataInPlace(missions[index], normalizedMission);
        }
    } else {
        const normalizedMission = normalizeMissionRecord({ id: Date.now(), ...missionData });
        if (normalizedMission) missions.push(normalizedMission);
    }

    persistMissions();
    clearEditSession();
    modal.style.display = "none";
    renderTimeline();
    renderMissionCards();
    if (committedMissionId != null) {
        refreshTooltipForMission(committedMissionId);
    }
});

function addDummyData() {
    const now = new Date();
    const y = now.getFullYear(), m = now.getMonth(), d = now.getDate();

    missions.push({
        id: 1, missionNum: 'CNV4469', tailNum: '695',
        legs: [
            { takeoffIcao: 'KNUW', takeoffTime: new Date(y, m, d - 2, 8, 0), landIcao: 'KNZY', landTime: new Date(y, m, d - 2, 14, 0) },
            { takeoffIcao: 'KNZY', takeoffTime: new Date(y, m, d - 1, 10, 0), landIcao: 'KNUW', landTime: new Date(y, m, d - 1, 19, 0) }
        ],
        pilot: 'John', copilot: 'Bhil', crewChief: 'Smeal', loadmaster: 'Valentine',
        liftCustomer: "DEVGRU", liftPax: 30, liftCargo: '3000', liftHazmat: 'None'
    });

    missions.push({
        id: 2, missionNum: 'CNV4869', tailNum: '834',
        legs: [
            { takeoffIcao: 'ETAR', takeoffTime: new Date(y, m, d + 2, 10, 0), landIcao: 'KADW', landTime: new Date(y, m, d + 2, 20, 0) }
        ],
        pilot: '', copilot: '', crewChief: '', loadmaster: '',
        liftPax: 91, liftCargo: '500', liftHazmat: 'Yes'
    });

    persistMissions();
}

init();
