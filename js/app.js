let missions = [];
let viewStart, viewEnd; 
const MS_PER_DAY = 24 * 60 * 60 * 1000;

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

function init() {
    const now = new Date();
    let currentFY = now.getFullYear();
    if (now.getMonth() >= 9) currentFY += 1; 
    document.getElementById('fy-input').value = currentFY;
    
    snapToCurrentMonth();
    addDummyData();
}

function snapToCurrentMonth() {
    const now = new Date();
    viewStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime(); 
    viewEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59).getTime(); 
    renderTimeline();
}

function snapToSevenDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const end = new Date(start);
    end.setDate(end.getDate() + 7);
    end.setHours(23, 59, 59, 999);

    viewStart = start.getTime();
    viewEnd = end.getTime();

    renderTimeline();
}

function snapToFourteenDayOutlook() {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const end = new Date(start);
    end.setDate(end.getDate() + 14);
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

document.getElementById('btn-snap-month').addEventListener('click', snapToCurrentMonth);
document.getElementById('btn-snap-7day').addEventListener('click', snapToSevenDayOutlook);
document.getElementById('btn-snap-14day').addEventListener('click', snapToFourteenDayOutlook);
document.getElementById('btn-snap-fy').addEventListener('click', () => snapToFY(parseInt(document.getElementById('fy-input').value, 10)));

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
    const start = mission.legs[0].takeoffTime.getTime();
    const end = mission.legs[mission.legs.length - 1].landTime.getTime();
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
    if (totalDaysVisible <= 90) {
        let iterDay = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
        while (iterDay <= new Date(viewEnd)) {
            const time = iterDay.getTime();
            const leftPct = ((time - viewStart) / duration) * 100;

            const marker = document.createElement('div');
            marker.className = 'grid-marker day';
            marker.style.left = `${leftPct}%`;

            if (pxPerDay > 22) {
                marker.innerHTML = `<div class="grid-label day-label">${iterDay.getDate()}</div>`;
            }
            gridContainer.appendChild(marker);
            iterDay.setDate(iterDay.getDate() + 1);
        }
    } else if (totalDaysVisible <= 365) {
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

        if (nextMonth) {
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
    const rowOffset = 10;

    // Dynamically adjust viewport height to fit all tails
    const requiredHeight = 104 + (uniqueTails.length * rowHeight) + rowOffset;
    viewport.style.height = `${Math.max(250, requiredHeight)}px`;

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
            tooltip.style.display = 'block';
            let itineraryHTML = mission.legs.map((l, i) => 
                `Leg ${i+1}: ${l.takeoffIcao} (${l.takeoffTime.toLocaleString()}) &rarr; ${l.landIcao} (${l.landTime.toLocaleString()})`
            ).join('<br>');

            tooltip.innerHTML = `
                <strong>Mission:</strong> ${mission.missionNum} | <strong>Tail:</strong> ${mission.tailNum}<br>
                <hr class="tooltip-separator">
                ${itineraryHTML}
            `;
            
            const card = document.getElementById(`card-${mission.id}`);
            if(card) { card.classList.add('highlight'); card.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
        });

        bar.addEventListener('mousemove', (e) => tooltip.style.cssText = `display: block; left: ${e.pageX + 15}px; top: ${e.pageY + 15}px;`);
        bar.addEventListener('mouseleave', () => {
            tooltip.style.display = 'none';
            const card = document.getElementById(`card-${mission.id}`);
            if(card) card.classList.remove('highlight');
        });

        bar.addEventListener('click', () => openEditModal(mission));
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

function deleteMission(id, event) {
    event.stopPropagation();
    if(confirm("Are you sure you want to delete this mission?")) {
        missions = missions.filter(m => m.id !== id);
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
                missions.push(m);
            }
        });
        
        renderTimeline();
        renderMissionCards();
        alert("CSV Imported Successfully!");
        e.target.value = ''; 
    };
    reader.readAsText(file);
});

document.getElementById('btn-export-csv').addEventListener('click', exportMissionsToCSV);

function toDateTimeLocal(date) {
    const pad = n => n < 10 ? '0' + n : n;
    return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) + 'T' + pad(date.getHours()) + ':' + pad(date.getMinutes());
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
    div.innerHTML = `
        <div class="form-row leg-row">
            <div class="form-col leg-col-short"><label class="leg-label">TK ICAO</label><input type="text" class="leg-tk-icao" required maxlength="4" value="${legData ? legData.takeoffIcao : ''}"></div>
            <div class="form-col leg-col-wide"><label class="leg-label">Takeoff Time</label><input type="datetime-local" class="leg-tk-time" required value="${legData ? toDateTimeLocal(legData.takeoffTime) : ''}"></div>
            <div class="form-col leg-col-short"><label class="leg-label">LD ICAO</label><input type="text" class="leg-ld-icao" required maxlength="4" value="${legData ? legData.landIcao : ''}"></div>
            <div class="form-col leg-col-wide"><label class="leg-label">Land Time</label><input type="datetime-local" class="leg-ld-time" required value="${legData ? toDateTimeLocal(legData.landTime) : ''}"></div>
            <div><button type="button" class="btn btn-remove-leg leg-remove-btn" title="Remove Leg">&times;</button></div>
        </div>
    `;
    div.querySelector('.btn-remove-leg').addEventListener('click', () => div.remove());
    legsWrapper.appendChild(div);
}

document.getElementById('btn-add-leg').addEventListener('click', () => addLegRow());

document.getElementById('btn-new-mission').addEventListener('click', () => {
    missionForm.reset();
    document.getElementById('editMissionId').value = "";
    legsWrapper.innerHTML = '';
    addLegRow();
    document.getElementById('modal-title').innerText = "New Mission";
    document.getElementById('btn-submit-form').innerText = "Create Mission";
    modal.style.display = "block";
});

document.getElementById('close-modal').addEventListener('click', () => modal.style.display = "none");

function openEditModal(mission) {
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
    
    const legNodes = document.querySelectorAll('.leg-container');
    if(legNodes.length === 0) return alert("You must add at least one flight leg.");

    let legs = [];
    let timeValid = true;

    legNodes.forEach(node => {
        const tkIcao = node.querySelector('.leg-tk-icao').value.toUpperCase();
        const ldIcao = node.querySelector('.leg-ld-icao').value.toUpperCase();
        const tkTime = new Date(node.querySelector('.leg-tk-time').value);
        const ldTime = new Date(node.querySelector('.leg-ld-time').value);

        if (ldTime <= tkTime) timeValid = false;
        legs.push({ takeoffIcao: tkIcao, takeoffTime: tkTime, landIcao: ldIcao, landTime: ldTime });
    });

    if (!timeValid) return alert("Landing time must be after takeoff time for all legs.");
    legs.sort((a, b) => a.takeoffTime - b.takeoffTime);

    const missionData = {
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
        legs: legs
    };

    const editId = document.getElementById('editMissionId').value;
    if (editId) {
        const index = missions.findIndex(m => m.id == editId);
        if (index > -1) missions[index] = { ...missions[index], ...missionData };
    } else {
        missions.push({ id: Date.now(), ...missionData });
    }

    modal.style.display = "none";
    renderTimeline();
    renderMissionCards();
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

    renderTimeline();
    renderMissionCards();
}

init();
