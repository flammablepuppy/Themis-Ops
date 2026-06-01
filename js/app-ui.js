function clearHoverRouteMap() {
    hoverTooltipToken += 1;
    activeTooltipMissionId = null;
    if (hoverRouteMap) {
        hoverRouteMap.remove();
        hoverRouteMap = null;
    }
}

function hideMissionTooltip() {
    tooltip.style.display = 'none';
    clearHoverRouteMap();
}

function clearTouchTimelineBarLongPress() {
    if (touchTimelineBarLongPressTimer !== null) {
        clearTimeout(touchTimelineBarLongPressTimer);
        touchTimelineBarLongPressTimer = null;
    }
    touchTimelineBarLongPressMissionId = null;
}

function openMissionCardFromTimeline(mission) {
    if (!mission) return;

    clearTouchTimelineBarLongPress();
    hideMissionTooltip();
    focusMissionCard(mission.id);
    openEditModal(mission);
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
        const takeoff = resolveAirportCode(leg.takeoffIcao) || normalizeIcao(leg.takeoffIcao);
        const landing = resolveAirportCode(leg.landIcao) || normalizeIcao(leg.landIcao);

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

function getRouteOffsetPlacement(map, point, offsetX = 0, offsetY = 0) {
    const layerPoint = map.latLngToLayerPoint(L.latLng(point));
    return map.layerPointToLatLng(layerPoint.add(L.point(offsetX, offsetY)));
}

function getRouteLegNumberPlacement(map, end, offsetX = 0, offsetY = 0) {
    return getRouteOffsetPlacement(map, end, offsetX, offsetY);
}

function distancePointToSegment(point, segmentStart, segmentEnd) {
    const delta = segmentEnd.subtract(segmentStart);
    const lengthSquared = (delta.x * delta.x) + (delta.y * delta.y);

    if (!lengthSquared) {
        return Math.hypot(point.x - segmentStart.x, point.y - segmentStart.y);
    }

    const t = Math.max(0, Math.min(1, (((point.x - segmentStart.x) * delta.x) + ((point.y - segmentStart.y) * delta.y)) / lengthSquared));
    const projection = L.point(segmentStart.x + (delta.x * t), segmentStart.y + (delta.y * t));
    return Math.hypot(point.x - projection.x, point.y - projection.y);
}

function distancePointToPoint(pointA, pointB) {
    return Math.hypot(pointA.x - pointB.x, pointA.y - pointB.y);
}

function getRouteLabelSize(labelSize = null) {
    const width = labelSize && Number.isFinite(labelSize.width) ? labelSize.width : ROUTE_ICAO_LABEL_MIN_WIDTH_PX;
    const height = labelSize && Number.isFinite(labelSize.height) ? labelSize.height : ROUTE_ICAO_LABEL_HEIGHT_PX;

    return {
        width,
        height,
        halfWidth: width / 2,
        halfHeight: height / 2
    };
}

function getRouteLabelBounds(placementPoint, labelSize = null) {
    const { width, height, halfWidth, halfHeight } = getRouteLabelSize(labelSize);

    return {
        left: placementPoint.x - halfWidth,
        right: placementPoint.x + halfWidth,
        top: placementPoint.y - halfHeight,
        bottom: placementPoint.y + halfHeight,
        width,
        height
    };
}

function pointInRect(point, rect) {
    return point.x >= rect.left && point.x <= rect.right && point.y >= rect.top && point.y <= rect.bottom;
}

function rectsOverlap(rectA, rectB) {
    return rectA.left < rectB.right && rectA.right > rectB.left && rectA.top < rectB.bottom && rectA.bottom > rectB.top;
}

function rectIntersectsCircle(rect, center, radius) {
    const closestX = Math.max(rect.left, Math.min(center.x, rect.right));
    const closestY = Math.max(rect.top, Math.min(center.y, rect.bottom));
    const dx = center.x - closestX;
    const dy = center.y - closestY;
    return ((dx * dx) + (dy * dy)) < (radius * radius);
}

function distancePointToRect(point, rect) {
    const dx = point.x < rect.left
        ? rect.left - point.x
        : point.x > rect.right
            ? point.x - rect.right
            : 0;
    const dy = point.y < rect.top
        ? rect.top - point.y
        : point.y > rect.bottom
            ? point.y - rect.bottom
            : 0;
    return Math.hypot(dx, dy);
}

function distanceRectToRect(rectA, rectB) {
    const dx = rectA.right < rectB.left
        ? rectB.left - rectA.right
        : rectB.right < rectA.left
            ? rectA.left - rectB.right
            : 0;
    const dy = rectA.bottom < rectB.top
        ? rectB.top - rectA.bottom
        : rectB.bottom < rectA.top
            ? rectA.top - rectB.bottom
            : 0;
    return Math.hypot(dx, dy);
}

function segmentsIntersect(pointA, pointB, pointC, pointD) {
    const EPSILON = 1e-9;

    function orientation(p, q, r) {
        const value = ((q.y - p.y) * (r.x - q.x)) - ((q.x - p.x) * (r.y - q.y));
        if (Math.abs(value) <= EPSILON) return 0;
        return value > 0 ? 1 : 2;
    }

    function onSegment(p, q, r) {
        return q.x >= Math.min(p.x, r.x) - EPSILON &&
            q.x <= Math.max(p.x, r.x) + EPSILON &&
            q.y >= Math.min(p.y, r.y) - EPSILON &&
            q.y <= Math.max(p.y, r.y) + EPSILON;
    }

    const o1 = orientation(pointA, pointB, pointC);
    const o2 = orientation(pointA, pointB, pointD);
    const o3 = orientation(pointC, pointD, pointA);
    const o4 = orientation(pointC, pointD, pointB);

    if (o1 !== o2 && o3 !== o4) return true;
    if (o1 === 0 && onSegment(pointA, pointC, pointB)) return true;
    if (o2 === 0 && onSegment(pointA, pointD, pointB)) return true;
    if (o3 === 0 && onSegment(pointC, pointA, pointD)) return true;
    if (o4 === 0 && onSegment(pointC, pointB, pointD)) return true;
    return false;
}

function rectIntersectsSegment(rect, startPoint, endPoint) {
    if (pointInRect(startPoint, rect) || pointInRect(endPoint, rect)) return true;

    const topLeft = L.point(rect.left, rect.top);
    const topRight = L.point(rect.right, rect.top);
    const bottomLeft = L.point(rect.left, rect.bottom);
    const bottomRight = L.point(rect.right, rect.bottom);

    return segmentsIntersect(startPoint, endPoint, topLeft, topRight) ||
        segmentsIntersect(startPoint, endPoint, topRight, bottomRight) ||
        segmentsIntersect(startPoint, endPoint, bottomRight, bottomLeft) ||
        segmentsIntersect(startPoint, endPoint, bottomLeft, topLeft);
}

function getRouteLabelPlacement(map, routePoints, index, labelSize = null, placedLabelBounds = [], markerPoints = null) {
    const currentPoint = map.latLngToLayerPoint(L.latLng(routePoints[index][0], routePoints[index][1]));
    const { halfWidth: labelHalfWidth, halfHeight: labelHalfHeight } = getRouteLabelSize(labelSize);
    const labelRadius = Math.max(labelHalfWidth, labelHalfHeight);
    const mapSize = map.getSize ? map.getSize() : null;
    const segmentPoints = [];
    const arrowPoints = [];
    const dotPoints = [];
    const searchAngles = [];

    for (let i = 0; i < routePoints.length - 1; i++) {
        const startRaw = routePoints[i];
        const endRaw = routePoints[i + 1];
        if (startRaw[0] === endRaw[0] && startRaw[1] === endRaw[1]) continue;

        const startPoint = map.latLngToLayerPoint(L.latLng(startRaw[0], startRaw[1]));
        const endPoint = map.latLngToLayerPoint(L.latLng(endRaw[0], endRaw[1]));
        segmentPoints.push([startPoint, endPoint]);
        arrowPoints.push(map.latLngToLayerPoint(getRouteSegmentPlacement(map, startRaw, endRaw).latLng));
    }

    const dotSourcePoints = Array.isArray(markerPoints) && markerPoints.length === routePoints.length ? markerPoints : routePoints;

    dotSourcePoints.forEach(point => {
        dotPoints.push(map.latLngToLayerPoint(L.latLng(point)));
    });

    for (let angleDeg = 0; angleDeg < 360; angleDeg += ROUTE_ICAO_LABEL_SEARCH_ANGLE_STEP_DEG) {
        const angleRad = angleDeg * Math.PI / 180;
        const dx = Math.cos(angleRad);
        const dy = Math.sin(angleRad);
        searchAngles.push({
            angleDeg,
            angleIndex: searchAngles.length,
            dx,
            dy,
            baseOffset: ROUTE_AIRPORT_DOT_EDGE_PX + ROUTE_ICAO_LABEL_VERTICAL_GAP_PX + (Math.abs(dx) * labelHalfWidth) + (Math.abs(dy) * labelHalfHeight)
        });
    }

    const minOffset = Math.max(0, searchAngles.reduce((min, entry) => Math.min(min, entry.baseOffset), Number.POSITIVE_INFINITY));

    function scoreCandidate(placementPoint) {
        const bounds = getRouteLabelBounds(placementPoint, labelSize);
        const boundsClearance = mapSize
            ? Math.min(
                bounds.left,
                mapSize.x - bounds.right,
                bounds.top,
                mapSize.y - bounds.bottom
            )
            : Number.POSITIVE_INFINITY;

        let labelClearance = Number.POSITIVE_INFINITY;
        let dotClearance = Number.POSITIVE_INFINITY;
        let segmentClearance = Number.POSITIVE_INFINITY;
        let arrowClearance = Number.POSITIVE_INFINITY;
        let collisionCount = 0;

        placedLabelBounds.forEach(otherBounds => {
            if (rectsOverlap(bounds, otherBounds)) collisionCount += 1;
            labelClearance = Math.min(labelClearance, distanceRectToRect(bounds, otherBounds));
        });

        dotPoints.forEach(dotPoint => {
            if (rectIntersectsCircle(bounds, dotPoint, ROUTE_AIRPORT_DOT_EDGE_PX)) collisionCount += 1;
            dotClearance = Math.min(dotClearance, distancePointToRect(dotPoint, bounds) - ROUTE_AIRPORT_DOT_EDGE_PX);
        });

        segmentPoints.forEach(([startPoint, endPoint]) => {
            if (rectIntersectsSegment(bounds, startPoint, endPoint)) collisionCount += 1;
            segmentClearance = Math.min(
                segmentClearance,
                distancePointToSegment(placementPoint, startPoint, endPoint) - labelRadius - ROUTE_ICAO_LABEL_LINE_CLEARANCE_PX
            );
        });

        arrowPoints.forEach(arrowPoint => {
            if (rectIntersectsCircle(bounds, arrowPoint, ROUTE_DIRECTION_MARKER_SAFE_RADIUS_PX)) collisionCount += 1;
            arrowClearance = Math.min(
                arrowClearance,
                distancePointToRect(arrowPoint, bounds) - ROUTE_DIRECTION_MARKER_SAFE_RADIUS_PX
            );
        });

        return {
            placementPoint,
            bounds,
            boundsClearance,
            segmentClearance,
            labelClearance,
            dotClearance,
            arrowClearance,
            clearance: Math.min(boundsClearance, segmentClearance, labelClearance, dotClearance, arrowClearance),
            collisionCount
        };
    }

    let bestCandidate = null;

    for (let offset = minOffset; offset <= (minOffset + ROUTE_ICAO_LABEL_MAX_SEARCH_PX); offset += ROUTE_ICAO_LABEL_SEARCH_STEP_PX) {
        let bestClearAtOffset = null;
        let bestAtOffset = null;

        for (let angleIndex = 0; angleIndex < searchAngles.length; angleIndex++) {
            const angle = searchAngles[angleIndex];
            if (offset + 1e-6 < angle.baseOffset) continue;

            const candidate = scoreCandidate(currentPoint.add(L.point(angle.dx * offset, angle.dy * offset)));
            candidate.angleIndex = angleIndex;
            candidate.angleDeg = angle.angleDeg;
            candidate.offset = offset;
            candidate.sideIndex = angleIndex;
            candidate.side = `angle-${angle.angleDeg}`;

            if (candidate.collisionCount === 0 && candidate.boundsClearance >= 0) {
                if (!bestClearAtOffset ||
                    candidate.clearance > bestClearAtOffset.clearance ||
                    (candidate.clearance === bestClearAtOffset.clearance && candidate.boundsClearance > bestClearAtOffset.boundsClearance) ||
                    (candidate.clearance === bestClearAtOffset.clearance && candidate.boundsClearance === bestClearAtOffset.boundsClearance && candidate.angleIndex < bestClearAtOffset.angleIndex)) {
                    bestClearAtOffset = candidate;
                }
            }

            if (!bestAtOffset ||
                candidate.collisionCount < bestAtOffset.collisionCount ||
                (candidate.collisionCount === bestAtOffset.collisionCount && candidate.clearance > bestAtOffset.clearance) ||
                (candidate.collisionCount === bestAtOffset.collisionCount && candidate.clearance === bestAtOffset.clearance && candidate.boundsClearance > bestAtOffset.boundsClearance) ||
                (candidate.collisionCount === bestAtOffset.collisionCount && candidate.clearance === bestAtOffset.clearance && candidate.boundsClearance === bestAtOffset.boundsClearance && candidate.angleIndex < bestAtOffset.angleIndex)) {
                bestAtOffset = candidate;
            }
        }

        if (bestClearAtOffset) {
            return {
                latLng: map.layerPointToLatLng(bestClearAtOffset.placementPoint),
                bounds: bestClearAtOffset.bounds,
                side: bestClearAtOffset.side,
                offset: bestClearAtOffset.offset
            };
        }

        if (bestAtOffset && (
            !bestCandidate ||
            bestAtOffset.collisionCount < bestCandidate.collisionCount ||
            (bestAtOffset.collisionCount === bestCandidate.collisionCount && bestAtOffset.clearance > bestCandidate.clearance) ||
            (bestAtOffset.collisionCount === bestCandidate.collisionCount && bestAtOffset.clearance === bestCandidate.clearance && bestAtOffset.boundsClearance > bestCandidate.boundsClearance) ||
            (bestAtOffset.collisionCount === bestCandidate.collisionCount && bestAtOffset.clearance === bestCandidate.clearance && bestAtOffset.boundsClearance === bestCandidate.boundsClearance && bestAtOffset.offset < bestCandidate.offset)
        )) {
            bestCandidate = bestAtOffset;
        }
    }

    const fallbackCandidate = bestCandidate || scoreCandidate(currentPoint);
    return {
        latLng: map.layerPointToLatLng(fallbackCandidate.placementPoint),
        bounds: fallbackCandidate.bounds,
        side: fallbackCandidate.side,
        offset: fallbackCandidate.offset
    };
}

function createRouteIcaoLabelIcon(code, labelSize = null) {
    const label = normalizeMissionText(code) || '';
    const resolvedSize = labelSize && Number.isFinite(labelSize.width) && Number.isFinite(labelSize.height)
        ? labelSize
        : getRouteIcaoLabelSize(label);

    return L.divIcon({
        className: 'route-icao-label',
        html: escapeHtml(label),
        iconSize: [resolvedSize.width, resolvedSize.height],
        iconAnchor: [Math.ceil(resolvedSize.width / 2), Math.ceil(resolvedSize.height / 2)]
    });
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

function createRouteLegNumberIcon(legNumber) {
    const label = String(legNumber);
    const width = Math.max(14, (label.length * 8) + 4);

    return L.divIcon({
        className: 'route-leg-number-marker',
        html: escapeHtml(label),
        iconSize: [width, 14],
        iconAnchor: [Math.ceil(width / 2), 7]
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
        ? legs.map((leg, index) => `<div class="tooltip-leg-line">${escapeHtml(getTooltipLegText(leg, index))}</div>`).join('')
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
        const routeLabelSizesByCode = new Map(routePointEntries.map(entry => [entry.code, getRouteIcaoLabelSize(entry.code)]));
        const routeLabelPadding = Math.max(
            18,
            Math.ceil(routePointEntries.reduce((max, entry) => {
                const labelSize = routeLabelSizesByCode.get(entry.code);
                if (!labelSize) return max;
                const placementExtent = Math.hypot(labelSize.width / 2, labelSize.height / 2) + ROUTE_AIRPORT_DOT_EDGE_PX + ROUTE_ICAO_LABEL_VERTICAL_GAP_PX + ROUTE_ICAO_LABEL_CLEARANCE_PX;
                return Math.max(max, placementExtent);
            }, 0))
        );

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
        hoverRouteMap.fitBounds(routePoints, { padding: [routeLabelPadding, routeLabelPadding] });

        // Repeated stops fan out to the right so duplicate destinations stay readable.
        // The first return to the mission origin stays centered on the origin dot.
        const routeOriginCode = routePointEntries[0] ? routePointEntries[0].code : '';
        const routePointOccurrenceCounts = new Map();
        const routePointPlacements = routePointEntries.map(entry => {
            const occurrenceIndex = routePointOccurrenceCounts.get(entry.code) || 0;
            routePointOccurrenceCounts.set(entry.code, occurrenceIndex + 1);

            if (occurrenceIndex === 0) {
                return L.latLng(entry.point);
            }

            if (entry.code === routeOriginCode && occurrenceIndex === 1) {
                return L.latLng(entry.point);
            }

            return getRouteOffsetPlacement(
                hoverRouteMap,
                entry.point,
                (entry.code === routeOriginCode ? occurrenceIndex - 1 : occurrenceIndex) * ROUTE_AIRPORT_DOT_RADIUS_PX,
                0
            );
        });

        routePointPlacements.forEach(point => {
            createAirportCircleMarker(point).addTo(routeLayer);
        });

        const placedLabelBounds = [];

        routePointEntries.forEach((entry, index) => {
            const { code } = entry;
            if (markerCodes.has(code)) return;
            markerCodes.add(code);
            const labelSize = routeLabelSizesByCode.get(code) || getRouteIcaoLabelSize(code);
            const labelPlacement = getRouteLabelPlacement(hoverRouteMap, routePoints, index, labelSize, placedLabelBounds, routePointPlacements);
            placedLabelBounds.push(labelPlacement.bounds);
            L.marker(labelPlacement.latLng, {
                interactive: false,
                keyboard: false,
                icon: createRouteIcaoLabelIcon(code, labelSize)
            }).addTo(routeLayer);
        });

        let legNumber = 1;
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

            const legPlacement = getRouteLegNumberPlacement(hoverRouteMap, routePointPlacements[i + 1]);
            L.marker(legPlacement, {
                interactive: false,
                keyboard: false,
                zIndexOffset: 1000,
                icon: createRouteLegNumberIcon(legNumber)
            }).addTo(routeLayer);
            legNumber += 1;
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
    const lifts = getMissionLiftsForRecord(mission);
    return {
        ...mission,
        lifts: lifts.map(lift => ({
            customer: lift.customer || '',
            pax: lift.pax || '',
            cargo: lift.cargo || '',
            hazmat: lift.hazmat || ''
        })),
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

function readLiftDataFromForm() {
    const liftNodes = document.querySelectorAll('.lift-container');
    const lifts = [];

    liftNodes.forEach(node => {
        const lift = {
            customer: normalizeMissionText(node.querySelector('.lift-customer')?.value),
            pax: normalizeMissionText(node.querySelector('.lift-pax')?.value),
            cargo: normalizeMissionText(node.querySelector('.lift-cargo')?.value),
            hazmat: normalizeMissionText(node.querySelector('.lift-hazmat')?.value)
        };

        if (!isMissionLiftEmpty(lift)) {
            lifts.push(lift);
        }
    });

    return lifts;
}

function buildMissionLiftFields(lifts) {
    const primaryLift = Array.isArray(lifts) && lifts.length > 0 ? lifts[0] : null;

    return {
        liftCustomer: primaryLift ? primaryLift.customer : '',
        liftPax: primaryLift ? primaryLift.pax : '',
        liftCargo: primaryLift ? primaryLift.cargo : '',
        liftHazmat: primaryLift ? primaryLift.hazmat : ''
    };
}

function readMissionDataFromForm() {
    const legNodes = document.querySelectorAll('.leg-container');
    const lifts = readLiftDataFromForm();
    let legs = [];
    let timeValid = true;

    legNodes.forEach(node => {
        const tkIcao = normalizeMissionText(node.querySelector('.leg-tk-icao').value);
        const ldIcao = normalizeMissionText(node.querySelector('.leg-ld-icao').value);
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
            tailNum: normalizeMissionText(document.getElementById('tailNum').value),
            pilot: document.getElementById('pilot').value,
            copilot: document.getElementById('copilot').value,
            crewChief: document.getElementById('crewChief').value,
            loadmaster: document.getElementById('loadmaster').value,
            ...buildMissionLiftFields(lifts),
            lifts,
            legs
        }
    };
}

async function buildMissionDataFromForm() {
    await loadAirportData().catch(() => {});

    const legNodes = document.querySelectorAll('.leg-container');
    const lifts = readLiftDataFromForm();
    let legs = [];
    let timeValid = true;
    let airportIssue = null;

    legNodes.forEach(node => {
        const tkInput = node.querySelector('.leg-tk-icao');
        const ldInput = node.querySelector('.leg-ld-icao');
        const tkTimeInput = node.querySelector('.leg-tk-time');
        const ldTimeInput = node.querySelector('.leg-ld-time');

        const tkText = normalizeMissionText(tkInput.value);
        const ldText = normalizeMissionText(ldInput.value);
        const tkValidIcao = isValidAirportIcaoCode(tkText);
        const ldValidIcao = isValidAirportIcaoCode(ldText);
        const tkSuggestions = tkValidIcao ? [] : getAirportLookupSuggestions(tkText);
        const ldSuggestions = ldValidIcao ? [] : getAirportLookupSuggestions(ldText);

        if (!tkValidIcao && !airportIssue) {
            airportIssue = {
                input: tkInput,
                label: 'Origin',
                suggestions: tkSuggestions
            };
        }

        if (!ldValidIcao && !airportIssue) {
            airportIssue = {
                input: ldInput,
                label: 'Destination',
                suggestions: ldSuggestions
            };
        }

        const tkTime = parseDateTimeLocalValue(tkTimeInput.value);
        const ldTime = parseDateTimeLocalValue(ldTimeInput.value);

        if (!tkTime || !ldTime || ldTime <= tkTime) timeValid = false;

        if (!airportIssue) {
            legs.push({
                takeoffIcao: normalizeIcao(tkText),
                takeoffTime: tkTime || new Date(NaN),
                landIcao: normalizeIcao(ldText),
                landTime: ldTime || new Date(NaN)
            });
        }
    });

    if (airportIssue) {
        return {
            hasLegs: legNodes.length > 0,
            timeValid,
            airportIssue,
            missionData: null
        };
    }

    return {
        hasLegs: legNodes.length > 0,
        timeValid,
        missionData: {
            missionNum: document.getElementById('missionNum').value.toUpperCase(),
            tailNum: normalizeMissionText(document.getElementById('tailNum').value),
            pilot: document.getElementById('pilot').value,
            copilot: document.getElementById('copilot').value,
            crewChief: document.getElementById('crewChief').value,
            loadmaster: document.getElementById('loadmaster').value,
            ...buildMissionLiftFields(lifts),
            lifts,
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

    const snapshot = readMissionDataFromForm();
    if (!snapshot.timeValid) return;

    applyMissionDataInPlace(mission, snapshot.missionData);
    refreshTooltipForMission(mission.id);
}

function getMissionHomeField() {
    return (missionDefaults.homeField || '').trim().toUpperCase();
}

function parseTailNumberList(value) {
    const source = Array.isArray(value)
        ? value
        : String(value || '').split(/[\n,;]+/);

    const seen = new Set();
    const tails = [];

    source.forEach(entry => {
        const tail = normalizeMissionText(entry);
        if (!tail) return;

        const key = tail.toUpperCase();
        if (seen.has(key)) return;

        seen.add(key);
        tails.push(tail);
    });

    return tails;
}

function getAvailableTailNumbers(selectedTail = '') {
    const configuredTails = parseTailNumberList(missionDefaults.availableTailNumbers);
    const selected = normalizeMissionText(selectedTail);

    if (!selected) return configuredTails;

    const selectedKey = selected.toUpperCase();
    if (configuredTails.some(tail => tail.toUpperCase() === selectedKey)) {
        return configuredTails;
    }

    return [selected, ...configuredTails];
}

function registerAvailableTailNumber(tailNumber) {
    const tail = normalizeMissionText(tailNumber);
    if (!tail) return false;

    const tailNumbers = parseTailNumberList(missionDefaults.availableTailNumbers);
    const tailKey = tail.toUpperCase();
    if (tailNumbers.some(item => item.toUpperCase() === tailKey)) {
        return false;
    }

    tailNumbers.push(tail);
    missionDefaults.availableTailNumbers = tailNumbers.join('\n');
    syncMissionDefaultsForm();
    persistMissionDefaults();
    syncTailNumberSelect(tail);
    return true;
}

function syncTailNumberSelect(selectedTail = null) {
    const tailInput = document.getElementById('tailNum');
    const tailOptions = document.getElementById('tailNumOptions');
    if (!tailInput || !tailOptions) return;

    const desiredTail = normalizeMissionText(selectedTail !== null ? selectedTail : tailInput.value);
    const tailNumbers = getAvailableTailNumbers(desiredTail);

    tailOptions.innerHTML = tailNumbers.map(tail => `<option value="${escapeHtml(tail)}"></option>`).join('');

    if (desiredTail) {
        tailInput.value = desiredTail;
    } else {
        tailInput.value = '';
    }
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

function syncLegLandTimeFromTakeoff(legNode) {
    const takeoffInput = legNode.querySelector('.leg-tk-time');
    const landInput = legNode.querySelector('.leg-ld-time');
    if (!takeoffInput || !landInput) return;

    const takeoffTime = parseDateTimeLocalValue(takeoffInput.value);
    if (!takeoffTime) return;

    const landTime = new Date(takeoffTime);
    landTime.setHours(landTime.getHours() + 3);
    landInput.value = toDateTimeLocal(landTime);
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

function updateTimelineTip() {
    const tip = document.getElementById('timeline-tip');
    if (!tip) return;
    const isTouch = window.matchMedia('(hover: none) and (pointer: coarse)').matches;
    tip.innerHTML = isTouch
        ? '<em>Tip: Pinch to zoom. Swipe to pan. Tap missions to view details.</em>'
        : '<em>Tip: Scroll to zoom. Drag to pan. Hover over missions to locate details.</em>';
}

function init() {
    const now = new Date();
    let currentFY = now.getFullYear();
    if (now.getMonth() >= 9) currentFY += 1;
    document.getElementById('fy-input').value = currentFY;
    updateTimelineTip();
    attachMissionCanonicalSyncObservers();

    missionDefaults = loadMissionDefaults();
    syncMissionDefaultsForm();

    const loadedMissions = loadMissions();
    if (loadedMissions !== null) {
        missions = loadedMissions;
    }

    snapToSevenDayOutlook();
    renderMissionCards();
    if (!currentTimeIndicatorRefreshTimer) {
        currentTimeIndicatorRefreshTimer = window.setInterval(syncCurrentTimeIndicator, 1000);
    }
    activateMissionSyncBackend();
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
}, { passive: false });

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

// Touch pan & pinch-zoom
function getTouchPinchDist(touches) {
    return Math.hypot(touches[1].clientX - touches[0].clientX, touches[1].clientY - touches[0].clientY);
}

viewport.addEventListener('touchstart', (e) => {
    e.preventDefault();
    touchTapMissionId = activeTooltipMissionId;
    hideMissionTooltip();
    touchMoved = false;
    touchActiveCount = e.touches.length;
    if (e.touches.length !== 1) {
        clearTouchTimelineBarLongPress();
    }

    if (e.touches.length === 1) {
        touchPanStartX = e.touches[0].clientX;
        touchPanStartViewStart = viewStart;
        touchPanStartViewEnd = viewEnd;
    } else if (e.touches.length === 2) {
        const rect = viewport.getBoundingClientRect();
        const midX = (e.touches[0].clientX + e.touches[1].clientX) / 2;
        touchPinchStartDist = getTouchPinchDist(e.touches);
        touchPinchStartMidPct = (midX - rect.left) / rect.width;
        touchPinchStartViewStart = viewStart;
        touchPinchStartViewEnd = viewEnd;
    }
}, { passive: false });

viewport.addEventListener('touchmove', (e) => {
    e.preventDefault();
    touchMoved = true;
    clearTouchTimelineBarLongPress();

    if (e.touches.length === 1 && touchActiveCount === 1) {
        const dx = e.touches[0].clientX - touchPanStartX;
        const rect = viewport.getBoundingClientRect();
        const msPerPixel = (touchPanStartViewEnd - touchPanStartViewStart) / rect.width;
        viewStart = touchPanStartViewStart - dx * msPerPixel;
        viewEnd = touchPanStartViewEnd - dx * msPerPixel;
        renderTimeline();
    } else if (e.touches.length === 2 && touchPinchStartDist > 0) {
        const currentDist = getTouchPinchDist(e.touches);
        const scale = touchPinchStartDist / currentDist;
        const baseDuration = touchPinchStartViewEnd - touchPinchStartViewStart;
        let newDuration = Math.min(Math.max(baseDuration * scale, 3 * MS_PER_DAY), 2 * 365 * MS_PER_DAY);
        const anchorTime = touchPinchStartViewStart + touchPinchStartMidPct * baseDuration;
        viewStart = anchorTime - touchPinchStartMidPct * newDuration;
        viewEnd = anchorTime + (1 - touchPinchStartMidPct) * newDuration;
        renderTimeline();
    }
}, { passive: false });

viewport.addEventListener('touchend', (e) => {
    touchActiveCount = e.touches.length;
    if (e.touches.length === 0) {
        clearTouchTimelineBarLongPress();
    }
    if (e.touches.length === 1) {
        // Dropped from 2 fingers to 1: restart pan tracking
        touchMoved = false;
        touchPanStartX = e.touches[0].clientX;
        touchPanStartViewStart = viewStart;
        touchPanStartViewEnd = viewEnd;
    }
}, { passive: true });

viewport.addEventListener('touchcancel', clearTouchTimelineBarLongPress, { passive: true });

function getMissionTimes(mission) {
    if(!mission.legs || mission.legs.length === 0) return { start: 0, end: 0 };
    const firstLeg = mission.legs[0] || {};
    const lastLeg = mission.legs[mission.legs.length - 1] || {};
    const start = getDateTimestamp(firstLeg.takeoffTime);
    const end = getDateTimestamp(lastLeg.landTime);
    return { start, end };
}

function hasCompleteMissionCrew(mission) {
    return Boolean(mission && mission.pilot && mission.copilot && mission.crewChief && mission.loadmaster);
}

function isMissionInProgress(mission, nowTime = Date.now()) {
    const { start, end } = getMissionTimes(mission);
    return Number.isFinite(start) &&
        Number.isFinite(end) &&
        start > 0 &&
        end > 0 &&
        start <= nowTime &&
        end > nowTime;
}

function getMissionCardStartTime(mission) {
    const start = getMissionTimes(mission).start;
    return start > 0 ? start : Number.POSITIVE_INFINITY;
}

function isMissionCardPast(mission, nowTime = Date.now()) {
    const start = getMissionCardStartTime(mission);
    return !Number.isFinite(start) || start < nowTime;
}

function getMissionCardOrderValue(mission, nowTime = Date.now()) {
    const start = getMissionCardStartTime(mission);
    if (!Number.isFinite(start)) return Number.MAX_SAFE_INTEGER;
    return start + (start < nowTime ? MISSION_CARD_PAST_ORDER_OFFSET : 0);
}

function compareMissionCardsByTakeoffTime(left, right, nowTime = Date.now()) {
    const leftStart = getMissionCardStartTime(left);
    const rightStart = getMissionCardStartTime(right);
    const leftPast = !Number.isFinite(leftStart) || leftStart < nowTime;
    const rightPast = !Number.isFinite(rightStart) || rightStart < nowTime;

    if (leftPast !== rightPast) return leftPast ? 1 : -1;
    if (leftStart !== rightStart) return leftStart - rightStart;

    return (left.missionNum || '').localeCompare(right.missionNum || '')
        || String(left.id || '').localeCompare(String(right.id || ''));
}

function applyMissionCardTemporalState(card, mission, nowTime = Date.now()) {
    if (!card || !mission) return;

    const isPast = isMissionCardPast(mission, nowTime);
    const isOngoing = isMissionInProgress(mission, nowTime);
    card.classList.toggle('past-mission', isPast);
    card.classList.toggle('ongoing-mission', isOngoing);
    card.dataset.missionPast = isPast ? '1' : '0';
    card.style.order = String(getMissionCardOrderValue(mission, nowTime));
}

function syncMissionCardTemporalState(nowTime = Date.now()) {
    const list = document.getElementById('mission-list');
    if (!list || list.children.length === 0) return;

    const missionById = new Map(missions.map(mission => [String(mission.id), mission]));
    list.querySelectorAll('.mission-card').forEach(card => {
        const mission = missionById.get(card.dataset.missionId || String(card.id || '').replace(/^card-/, ''));
        if (!mission) return;
        applyMissionCardTemporalState(card, mission, nowTime);
    });
}

function applyMissionTimelineTemporalState(bar, mission, nowTime = Date.now()) {
    if (!bar || !mission) return;

    const { start, end } = getMissionTimes(mission);
    const hasCrew = hasCompleteMissionCrew(mission);
    const isOngoing = isMissionInProgress(mission, nowTime);
    const isPast = Number.isFinite(end) && end > 0 && end <= nowTime;
    const isReadyToStart = !isOngoing && !isPast && hasCrew && Number.isFinite(start) && start > nowTime;
    const isUnassignedCrew = !isOngoing && !isPast && !hasCrew;

    bar.classList.toggle('ongoing', isOngoing);
    bar.classList.toggle('past', isPast);
    bar.classList.toggle('ready-to-start', isReadyToStart);
    bar.classList.toggle('unassigned-crew', isUnassignedCrew);
}

function syncMissionTimelineTemporalState(nowTime = Date.now()) {
    const list = document.getElementById('timeline-missions');
    if (!list || list.children.length === 0) return;

    const missionById = new Map(missions.map(mission => [String(mission.id), mission]));
    list.querySelectorAll('.timeline-bar').forEach(bar => {
        const mission = missionById.get(bar.dataset.missionId || String(bar.id || '').replace(/^timeline-/, ''));
        if (!mission) return;
        applyMissionTimelineTemporalState(bar, mission, nowTime);
    });
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
    const nowTime = now.getTime();
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

    syncCurrentTimeIndicator();

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
    const rowHeight = 92;
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
        label.style.top = `${topPx + 16}px`;
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
        bar.id = `timeline-${mission.id}`;
        bar.dataset.missionId = String(mission.id);
        applyMissionTimelineTemporalState(bar, mission, nowTime);

        const clampedLeft = Math.max(0, leftPct);
        const clampedRight = Math.min(100, leftPct + widthPct);
        bar.style.left = `${clampedLeft}%`;
        bar.style.width = `${Math.max(0, clampedRight - clampedLeft)}%`;
        
        // Assign top position based on matching tail row
        const tailStr = mission.tailNum ? mission.tailNum.toUpperCase() : 'TBD';
        const tailIndex = uniqueTails.indexOf(tailStr);
        bar.style.top = `${rowOffset + (tailIndex * rowHeight) + 16}px`;
        
        const firstIcao = escapeHtml(mission.legs[0].takeoffIcao);
        const lastIcao = escapeHtml(mission.legs[mission.legs.length - 1].landIcao);
        bar.innerHTML = `<span class="timeline-bar-mission">${escapeHtml(mission.missionNum)}</span><span class="timeline-bar-route">(${firstIcao}&rarr;${lastIcao})</span>`;

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
            hideMissionTooltip();
            const card = document.getElementById(`card-${mission.id}`);
            if(card) card.classList.remove('highlight');
        });

        bar.addEventListener('click', () => openMissionCardFromTimeline(mission));
        bar.addEventListener('touchstart', (e) => {
            if (e.touches.length !== 1) {
                clearTouchTimelineBarLongPress();
                return;
            }

            clearTouchTimelineBarLongPress();
            touchTimelineBarLongPressMissionId = mission.id;
            touchTimelineBarLongPressTimer = window.setTimeout(() => {
                if (String(touchTimelineBarLongPressMissionId) !== String(mission.id)) return;

                touchMoved = true;
                clearTouchTimelineBarLongPress();
                openMissionCardFromTimeline(mission);
            }, TIMELINE_BAR_LONG_PRESS_DELAY_MS);
        }, { passive: true });
        bar.addEventListener('touchend', (e) => {
            clearTouchTimelineBarLongPress();
            if (touchMoved) return;
            if (touchTapMissionId != null && String(touchTapMissionId) === String(mission.id)) return;
            const touch = e.changedTouches[0];
            if (!touch) return;
            showMissionTooltip(mission, { pageX: touch.pageX, pageY: touch.pageY });
        }, { passive: true });
        missionsContainer.appendChild(bar);
    });
}

function renderMissionCards() {
    const list = document.getElementById('mission-list');
    clearMissionCardFocusHighlight();
    list.innerHTML = '';
    const nowTime = Date.now();
    const sorted = [...missions].sort((a, b) => compareMissionCardsByTakeoffTime(a, b, nowTime));

    sorted.forEach(mission => {
        const card = document.createElement('div');
        card.className = 'mission-card';
        card.id = `card-${mission.id}`;
        card.dataset.missionId = String(mission.id);
        applyMissionCardTemporalState(card, mission, nowTime);

        const isComplete = !!(mission.missionNum && mission.tailNum && mission.legs.length > 0 && hasCompleteMissionCrew(mission));

        if (!isComplete) card.classList.add('no-crew');
        const checkMark = isComplete ? `<span class="status-check">&#10003;</span>` : '';

        let legsHTML = mission.legs.map((leg, idx) => `
                <li><strong>Leg ${idx+1}:</strong> ${escapeHtml(leg.takeoffIcao)} (${leg.takeoffTime.toLocaleString(undefined, {month:'short', day:'numeric', hour:'2-digit', minute:'2-digit'})}) &rarr;
                ${escapeHtml(leg.landIcao)} (${leg.landTime.toLocaleString(undefined, {month:'short', day:'numeric', hour:'2-digit', minute:'2-digit'})})</li>
        `).join('');

        card.innerHTML = `
            <div class="card-col col-mission">
                <div class="mission-header">
                    <h3>${checkMark} Mission: ${escapeHtml(mission.missionNum)}</h3>
                </div>
                <div class="mission-details">
                    <p class="mission-tail-line"><strong>Tail #:</strong> ${mission.tailNum ? escapeHtml(mission.tailNum) : '<em>TBD</em>'}</p>
                    <ul class="itinerary-list">${legsHTML}</ul>
                    <span class="edit-hint">Click card anywhere to edit</span>
                </div>
            </div>

            <div class="card-col col-crew">
                <div class="crew-section">
                    <h4>Crew</h4>
                    <p><strong>Pilot:</strong> ${mission.pilot ? escapeHtml(mission.pilot) : '<em>TBD</em>'}</p>
                    <p><strong>Co-Pilot:</strong> ${mission.copilot ? escapeHtml(mission.copilot) : '<em>TBD</em>'}</p>
                    <p><strong>Crew Chief:</strong> ${mission.crewChief ? escapeHtml(mission.crewChief) : '<em>TBD</em>'}</p>
                    <p><strong>Loadmaster(s):</strong> ${mission.loadmaster ? escapeHtml(mission.loadmaster) : '<em>TBD</em>'}</p>
                </div>
            </div>

            <div class="card-col col-lift">
                <div class="lift-section">
                    <h4>Lift</h4>
                    ${buildMissionLiftCardHTML(mission)}
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

function animateMissionEntryElement(element, animationClass) {
    if (!element) return;

    element.classList.add(animationClass);
    window.setTimeout(() => {
        if (element.isConnected) {
            element.classList.remove(animationClass);
        }
    }, MISSION_ENTRY_ANIMATION_DURATION_MS);
}

function formatMissionLiftDisplayValue(value, fallback) {
    return value ? escapeHtml(String(value)) : fallback;
}

function buildMissionLiftCardHTML(mission) {
    const lifts = getMissionLiftsForRecord(mission);
    if (lifts.length === 0) {
        return '<p><em>No lifts added.</em></p>';
    }

    return lifts.map((lift, index) => `
        <div class="lift-entry">
            <div class="lift-entry-title">Lift ${index + 1}</div>
            <p><strong>Customer:</strong> ${formatMissionLiftDisplayValue(lift.customer, '0')}</p>
            <p><strong>Pax:</strong> ${formatMissionLiftDisplayValue(lift.pax, '0')}</p>
            <p><strong>Cargo:</strong> ${formatMissionLiftDisplayValue(lift.cargo, '<em>None</em>')}</p>
            <p><strong>Hazmat:</strong> ${formatMissionLiftDisplayValue(lift.hazmat, '<em>None</em>')}</p>
        </div>
    `).join('');
}

function animateAddedMission(id) {
    window.requestAnimationFrame(() => {
        animateMissionEntryElement(document.getElementById(`card-${id}`), 'mission-card-entry-animating');
        animateMissionEntryElement(document.getElementById(`timeline-${id}`), 'mission-timeline-entry-animating');
    });
}

function createMissionRemovalGhost(element, host, positionStrategy, animationClass = 'mission-card-delete-animating') {
    if (!element || !host) return;

    const rect = element.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return;

    const ghost = element.cloneNode(true);
    ghost.removeAttribute('id');
    ghost.classList.add(animationClass);
    ghost.style.margin = '0';
    ghost.style.pointerEvents = 'none';
    ghost.style.zIndex = '9999';
    ghost.style.transformOrigin = 'center center';
    ghost.style.width = `${rect.width}px`;
    ghost.style.height = `${rect.height}px`;

    if (positionStrategy === 'absolute') {
        const hostRect = host.getBoundingClientRect();
        ghost.style.position = 'absolute';
        ghost.style.left = `${rect.left - hostRect.left}px`;
        ghost.style.top = `${rect.top - hostRect.top}px`;
    } else {
        ghost.style.position = 'fixed';
        ghost.style.left = `${rect.left}px`;
        ghost.style.top = `${rect.top}px`;
    }

    host.appendChild(ghost);
    window.setTimeout(() => ghost.remove(), MISSION_ENTRY_ANIMATION_DURATION_MS + 40);
}

function animateDeletedMission(id) {
    createMissionRemovalGhost(document.getElementById(`card-${id}`), document.body, 'fixed', 'mission-card-delete-animating');
    createMissionRemovalGhost(document.getElementById(`timeline-${id}`), viewport, 'absolute', 'mission-timeline-delete-animating');
}

function animateSyncedMissionViews(missionIds = null) {
    const resolvedMissionIds = missionIds == null ? null : Array.from(missionIds, id => String(id)).filter(Boolean);
    window.requestAnimationFrame(() => {
        const cards = resolvedMissionIds
            ? resolvedMissionIds.map(id => document.getElementById(`card-${id}`)).filter(Boolean)
            : document.querySelectorAll('.mission-card');
        cards.forEach(card => {
            animateMissionEntryElement(card, 'mission-card-sync-animating');
        });

        const bars = resolvedMissionIds
            ? resolvedMissionIds.map(id => document.getElementById(`timeline-${id}`)).filter(Boolean)
            : document.querySelectorAll('.timeline-bar');
        bars.forEach(bar => {
            animateMissionEntryElement(bar, 'mission-timeline-sync-animating');
        });
    });
}

function renderMissionViewsAfterSync(animationState = false) {
    const normalizedState = typeof animationState === 'boolean'
        ? {
            shouldAnimateRefresh: animationState,
            addedMissionIds: [],
            removedMissionIds: [],
            changedMissionIds: [],
            animateAllMissionIds: animationState
        }
        : {
            shouldAnimateRefresh: Boolean(animationState && animationState.shouldAnimateRefresh),
            addedMissionIds: Array.isArray(animationState && animationState.addedMissionIds)
                ? animationState.addedMissionIds.map(id => String(id))
                : [],
            removedMissionIds: Array.isArray(animationState && animationState.removedMissionIds)
                ? animationState.removedMissionIds.map(id => String(id))
                : [],
            changedMissionIds: Array.isArray(animationState && animationState.changedMissionIds)
                ? animationState.changedMissionIds.map(id => String(id))
                : [],
            animateAllMissionIds: Boolean(animationState && animationState.animateAllMissionIds)
        };

    normalizedState.removedMissionIds.forEach(id => animateDeletedMission(id));

    renderTimeline();
    renderMissionCards();

    if (normalizedState.animateAllMissionIds) {
        animateSyncedMissionViews();
    } else if (normalizedState.changedMissionIds.length > 0) {
        animateSyncedMissionViews(normalizedState.changedMissionIds);
    }

    normalizedState.addedMissionIds.forEach(id => {
        animateAddedMission(id);
    });

    if (activeTooltipMissionId != null) {
        refreshTooltipForMission(activeTooltipMissionId);
    }
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
    clearMissionCardFocusHighlight();

    const card = document.getElementById(`card-${id}`);
    if (!card) return;

    focusedMissionCardId = id;
    card.classList.remove('highlight');
    card.classList.add('focus-highlight');
    card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    missionCardFocusTimeout = window.setTimeout(() => {
        if (focusedMissionCardId !== id) return;

        card.classList.remove('focus-highlight');
        focusedMissionCardId = null;
        missionCardFocusTimeout = null;
    }, 5000);
}

function deleteMission(id, event) {
    event.stopPropagation();
    if(confirm("Are you sure you want to delete this mission?")) {
        animateDeletedMission(id);
        markMissionDeleted(id);
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
        
        const headers = parseCsvLine(lines[0]).map(h => h.trim().toLowerCase());
        const missionMap = new Map();
        const importedMissionIds = [];

        for(let i=1; i<lines.length; i++) {
            const cols = parseCsvLine(lines[i]).map(c => c.trim());
            let row = {};
            headers.forEach((h, idx) => row[h] = cols[idx]);
            
            const mNum = row.mission;
            if(!mNum) continue;

            if(!missionMap.has(mNum)) {
                const importedLift = {
                    customer: row.customer || '',
                    pax: row.pax || '',
                    cargo: row.cargo || '',
                    hazmat: row.hazmat || ''
                };
                missionMap.set(mNum, {
                    id: createMissionId('csv'),
                    missionNum: mNum.toUpperCase(),
                    tailNum: row.tail || '',
                    pilot: row.pilot || '',
                    copilot: row.copilot || '',
                    crewChief: row.crewchief || '',
                    loadmaster: row.loadmaster || '',
                    liftCustomer: importedLift.customer,
                    liftPax: importedLift.pax,
                    liftCargo: importedLift.cargo,
                    liftHazmat: importedLift.hazmat,
                    lifts: isMissionLiftEmpty(importedLift) ? [] : [importedLift],
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
                if (normalizedMission) {
                    missions.push(normalizedMission);
                    markMissionUpdated(normalizedMission.id);
                    importedMissionIds.push(normalizedMission.id);
                }
            }
        });
        
        persistMissions();
        renderTimeline();
        renderMissionCards();
        importedMissionIds.forEach(id => animateAddedMission(id));
        window.setTimeout(() => alert("CSV Imported Successfully!"), MISSION_ENTRY_ANIMATION_DURATION_MS + 50);
        e.target.value = ''; 
    };
    reader.readAsText(file);
});

document.getElementById('btn-export-csv').addEventListener('click', exportMissionsToCSV);
missionDefaultsForm.addEventListener('submit', function(e) {
    e.preventDefault();
    missionDefaults = readMissionDefaultsForm();
    persistMissionDefaults();
    syncMissionDefaultsForm();
    activateMissionSyncBackend();
});

missionDefaultsForm.addEventListener('input', event => {
    if (event.target === missionDataPathInput && !missionDataFileHandleDisabled) {
        void clearStoredMissionDataHandle();
    }
    const oneDriveConfigInputs = [
        missionOneDriveTenantIdInput,
        missionOneDriveClientIdInput,
        missionOneDriveFilePathInput,
        missionOneDriveRedirectUriInput
    ];
    const oneDriveAuthInputs = [
        missionOneDriveTenantIdInput,
        missionOneDriveClientIdInput,
        missionOneDriveRedirectUriInput
    ];
    const sharePointConfigInputs = [
        missionSharePointTenantIdInput,
        missionSharePointClientIdInput,
        missionSharePointSiteIdInput,
        missionSharePointListIdentifierInput,
        missionSharePointItemTitleInput,
        missionSharePointPayloadFieldInput,
        missionSharePointRedirectUriInput
    ];
    const sharePointAuthInputs = [
        missionSharePointTenantIdInput,
        missionSharePointClientIdInput,
        missionSharePointRedirectUriInput
    ];
    const googleCloudConfigInputs = [
        missionGoogleCloudClientIdInput,
        missionGoogleCloudFileNameInput,
        missionGoogleCloudFolderIdInput
    ];
    const googleCloudAuthInputs = [
        missionGoogleCloudClientIdInput
    ];
    const wasBackendChange = event.target === missionSyncBackendInput;
    const wasOneDriveConfigChange = oneDriveConfigInputs.includes(event.target);
    const wasOneDriveAuthChange = oneDriveAuthInputs.includes(event.target);
    const wasSharePointConfigChange = sharePointConfigInputs.includes(event.target);
    const wasSharePointAuthChange = sharePointAuthInputs.includes(event.target);
    const wasGoogleCloudConfigChange = googleCloudConfigInputs.includes(event.target);
    const wasGoogleCloudAuthChange = googleCloudAuthInputs.includes(event.target);

    missionDefaults = readMissionDefaultsForm();

    if (wasOneDriveConfigChange && event.target !== missionOneDriveItemIdInput && event.target !== missionOneDriveItemETagInput) {
        if (wasOneDriveAuthChange) {
            resetOneDriveAuthClient();
        }
        clearOneDriveItemReference();
    }

    if (wasSharePointConfigChange && event.target !== missionSharePointItemIdInput && event.target !== missionSharePointItemETagInput) {
        if (wasSharePointAuthChange) {
            resetSharePointAuthClient();
        }
        clearSharePointItemReference();
    }

    if (wasGoogleCloudConfigChange && event.target !== missionGoogleCloudFileIdInput) {
        if (wasGoogleCloudAuthChange) {
            resetGoogleCloudAuthClient();
        }
        clearGoogleCloudItemReference();
    }

    persistMissionDefaults();
    syncMissionDefaultsForm();

    if (event.target === missionTailNumbersInput) {
        syncTailNumberSelect(document.getElementById('tailNum')?.value || '');
    }

    if (wasBackendChange) {
        activateMissionSyncBackend();
    } else if (wasOneDriveConfigChange) {
        updateOneDriveStatusFromState();
        if (isOneDriveSyncMode() && hasOneDriveSyncConfiguration()) {
            startMissionCanonicalSyncLoop();
        }
    } else if (wasSharePointConfigChange) {
        updateSharePointStatusFromState();
        if (isSharePointSyncMode() && hasSharePointSyncConfiguration()) {
            startMissionCanonicalSyncLoop();
        }
    } else if (wasGoogleCloudConfigChange) {
        updateGoogleCloudStatusFromState();
        if (isGoogleCloudSyncMode() && hasGoogleCloudSyncConfiguration()) {
            startMissionCanonicalSyncLoop();
        }
    }
});

document.getElementById('btn-clear-defaults').addEventListener('click', () => {
    missionDefaults = createEmptyMissionDefaults();
    persistMissionDefaults();
    syncMissionDefaultsForm();
    syncTailNumberSelect('');
    resetOneDriveAuthClient();
    clearOneDriveItemReference();
    resetSharePointAuthClient();
    clearSharePointItemReference();
    resetGoogleCloudAuthClient();
    clearGoogleCloudItemReference();
    void clearStoredMissionDataHandle();
});

document.getElementById('btn-choose-mission-data-file').addEventListener('click', () => {
    void requestMissionDataHandle();
});

sharePointConnectButton?.addEventListener('click', () => {
    void requestSharePointConnection();
});

sharePointSyncNowButton?.addEventListener('click', () => {
    void syncSharePointNow();
});

oneDriveConnectButton?.addEventListener('click', () => {
    void requestOneDriveConnection();
});

oneDriveSyncNowButton?.addEventListener('click', () => {
    void syncOneDriveNow();
});

googleCloudConnectButton?.addEventListener('click', () => {
    void requestGoogleCloudConnection();
});

googleCloudSyncNowButton?.addEventListener('click', () => {
    void syncGoogleCloudNow();
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
    const headers = ['mission', 'tail', 'pilot', 'copilot', 'crewchief', 'loadmaster', 'customer', 'pax', 'cargo', 'hazmat', 'takeoff_icao', 'takeoff_time', 'land_icao', 'land_time'];
    const rows = [headers.join(',')];

    missions.forEach(mission => {
        if (!mission.legs || mission.legs.length === 0) return;

        const primaryLift = getMissionLiftsForRecord(mission)[0] || {
            customer: '',
            pax: '',
            cargo: '',
            hazmat: ''
        };

        mission.legs.forEach(leg => {
            rows.push([
                mission.missionNum || '',
                mission.tailNum || '',
                mission.pilot || '',
                mission.copilot || '',
                mission.crewChief || '',
                mission.loadmaster || '',
                primaryLift.customer || '',
                primaryLift.pax || '',
                primaryLift.cargo || '',
                primaryLift.hazmat || '',
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
            <div class="form-col leg-col-short leg-airport-field">
                <label class="leg-label">Origin</label>
                <input type="text" class="leg-tk-icao" required value="${escapeHtml(defaultLegData.takeoffIcao || '')}" aria-autocomplete="list" aria-expanded="false">
                <div class="airport-name-display" aria-live="polite" hidden></div>
                <div class="airport-lookup-panel" role="listbox" hidden></div>
            </div>
            <div class="form-col leg-col-wide"><label class="leg-label">Takeoff Time</label><input type="datetime-local" class="leg-tk-time" required value="${toDateTimeLocal(defaultLegData.takeoffTime)}"></div>
            <div class="form-col leg-col-short leg-airport-field">
                <label class="leg-label">Destination</label>
                <input type="text" class="leg-ld-icao" required value="${escapeHtml(defaultLegData.landIcao || '')}" aria-autocomplete="list" aria-expanded="false">
                <div class="airport-name-display" aria-live="polite" hidden></div>
                <div class="airport-lookup-panel" role="listbox" hidden></div>
            </div>
            <div class="form-col leg-col-wide"><label class="leg-label">Land Time</label><input type="datetime-local" class="leg-ld-time" required value="${toDateTimeLocal(defaultLegData.landTime)}"></div>
            <div><button type="button" class="btn btn-remove-leg leg-remove-btn" title="Remove Leg">&times;</button></div>
        </div>
    `;    
    div.querySelector('.btn-remove-leg').addEventListener('click', () => {
        div.remove();
        syncEditingMissionPreview();
    });
    const takeoffTimeInput = div.querySelector('.leg-tk-time');
    if (takeoffTimeInput) {
        const updateLandTime = () => syncLegLandTimeFromTakeoff(div);
        takeoffTimeInput.addEventListener('input', updateLandTime);
        takeoffTimeInput.addEventListener('change', updateLandTime);
    }
    const takeoffIcaoInput = div.querySelector('.leg-tk-icao');
    const landIcaoInput = div.querySelector('.leg-ld-icao');
    attachAirportLookupBehavior(takeoffIcaoInput);
    attachAirportLookupBehavior(landIcaoInput);
    legsWrapper.appendChild(div);
    updateAirportNameDisplay(takeoffIcaoInput);
    updateAirportNameDisplay(landIcaoInput);
    if (takeoffIcaoInput) scheduleAirportLookupPanelUpdate(takeoffIcaoInput);
    if (landIcaoInput) scheduleAirportLookupPanelUpdate(landIcaoInput);
    if (!legData) syncEditingMissionPreview();
}

function addLiftRow(liftData = null) {
    const div = document.createElement('div');
    div.className = 'lift-container';
    const defaultLiftData = liftData || {
        customer: '',
        pax: '',
        cargo: '',
        hazmat: ''
    };
    div.innerHTML = `
        <div class="form-row lift-row">
            <div class="form-col lift-col">
                <label class="lift-label">Customer</label>
                <input type="text" class="lift-customer" value="${escapeHtml(defaultLiftData.customer || '')}">
            </div>
            <div class="form-col lift-col">
                <label class="lift-label">Pax</label>
                <input type="number" class="lift-pax" min="0" value="${escapeHtml(defaultLiftData.pax || '')}">
            </div>
            <div class="form-col lift-col">
                <label class="lift-label">Cargo</label>
                <input type="text" class="lift-cargo" value="${escapeHtml(defaultLiftData.cargo || '')}">
            </div>
            <div class="form-col lift-col">
                <label class="lift-label">Hazmat</label>
                <input type="text" class="lift-hazmat" value="${escapeHtml(defaultLiftData.hazmat || '')}">
            </div>
            <div><button type="button" class="btn btn-remove-leg lift-remove-btn" title="Remove Lift">&times;</button></div>
        </div>
    `;
    div.querySelector('.lift-remove-btn').addEventListener('click', () => {
        div.remove();
        syncEditingMissionPreview();
    });
    liftsWrapper.appendChild(div);
    if (!liftData) syncEditingMissionPreview();
}

document.getElementById('btn-add-leg').addEventListener('click', () => addLegRow());
document.getElementById('btn-add-lift').addEventListener('click', () => addLiftRow());
missionForm.addEventListener('input', event => {
    if (event.target.closest('.leg-container') || event.target.closest('.lift-container')) syncEditingMissionPreview();
});
missionForm.addEventListener('change', event => {
    if (event.target.closest('.leg-container') || event.target.closest('.lift-container')) syncEditingMissionPreview();
});

document.getElementById('btn-new-mission').addEventListener('click', () => {
    missionForm.reset();
    document.getElementById('editMissionId').value = "";
    legsWrapper.innerHTML = '';
    liftsWrapper.innerHTML = '';
    syncTailNumberSelect('');
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
    syncTailNumberSelect(mission.tailNum);
    document.getElementById('pilot').value = mission.pilot;
    document.getElementById('copilot').value = mission.copilot;
    document.getElementById('crewChief').value = mission.crewChief;
    document.getElementById('loadmaster').value = mission.loadmaster;
    liftsWrapper.innerHTML = '';
    const missionLifts = getMissionLiftsForRecord(mission);
    missionLifts.forEach(lift => addLiftRow(lift));
    
    legsWrapper.innerHTML = '';
    const missionLegs = Array.isArray(mission.legs) ? mission.legs : [];
    missionLegs.forEach(leg => addLegRow(leg));
    
    document.getElementById('modal-title').innerText = "Edit Mission";
    document.getElementById('btn-submit-form').innerText = "Update Mission";
    modal.style.display = "block";
}

missionForm.addEventListener('submit', async function(e) {
    e.preventDefault();
    const { hasLegs, timeValid, missionData, airportIssue } = await buildMissionDataFromForm();
    if (!hasLegs) return alert("You must add at least one flight leg.");
    if (!timeValid) return alert("Landing time must be after takeoff time for all legs.");
    if (!missionData) {
        if (airportIssue && airportIssue.input) {
            renderAirportLookupPanel(airportIssue.input, airportIssue.suggestions || []);
            airportIssue.input.focus({ preventScroll: true });
        }
        if (airportIssue) {
            alert(
                airportIssue.suggestions && airportIssue.suggestions.length > 0
                    ? `Select an airport from the ${airportIssue.label.toLowerCase()} matches before submitting.`
                    : `No airport matches found for the ${airportIssue.label.toLowerCase()} field.`
            );
        }
        return;
    }
    missionData.legs.sort((a, b) => a.takeoffTime - b.takeoffTime);

    const tailNum = normalizeMissionText(missionData.tailNum);
    missionData.tailNum = tailNum;
    if (tailNum) {
        registerAvailableTailNumber(tailNum);
    }

    const editId = document.getElementById('editMissionId').value;
    const committedMissionId = editId || null;
    let newlyCreatedMissionId = null;
    if (editId) {
        const index = missions.findIndex(m => m.id == editId);
        if (index > -1) {
            const normalizedMission = normalizeMissionRecord({ id: missions[index].id, ...missionData });
            if (normalizedMission) {
                applyMissionDataInPlace(missions[index], normalizedMission);
                markMissionUpdated(missions[index].id);
            }
        }
    } else {
        const normalizedMission = normalizeMissionRecord({ id: createMissionId(), ...missionData });
        if (normalizedMission) {
            missions.push(normalizedMission);
            markMissionUpdated(normalizedMission.id);
            newlyCreatedMissionId = normalizedMission.id;
        }
    }

    persistMissions();
    clearEditSession();
    modal.style.display = "none";
    renderTimeline();
    renderMissionCards();
    if (newlyCreatedMissionId) {
        animateAddedMission(newlyCreatedMissionId);
    }
    if (committedMissionId != null) {
        refreshTooltipForMission(committedMissionId);
    }
});

