let missions = [];
let viewStart, viewEnd; 
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MS_PER_HOUR = 60 * 60 * 1000;
const CURRENT_TIME_BUBBLE_HEIGHT_PX = 38;

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
const liftsWrapper = document.getElementById('lifts-wrapper');
const missionDefaultsForm = document.getElementById('mission-defaults-form');
const missionHomeFieldInput = document.getElementById('defaultHomeField');
const missionTailNumbersInput = document.getElementById('defaultTailNumbers');
const missionDataPathInput = document.getElementById('defaultMissionDataPath');
const missionSyncBackendInput = document.getElementById('defaultSyncBackend');
const missionSharePointTenantIdInput = document.getElementById('defaultSharePointTenantId');
const missionSharePointClientIdInput = document.getElementById('defaultSharePointClientId');
const missionSharePointSiteIdInput = document.getElementById('defaultSharePointSiteId');
const missionSharePointListIdentifierInput = document.getElementById('defaultSharePointListIdentifier');
const missionSharePointItemTitleInput = document.getElementById('defaultSharePointItemTitle');
const missionSharePointPayloadFieldInput = document.getElementById('defaultSharePointPayloadField');
const missionSharePointRedirectUriInput = document.getElementById('defaultSharePointRedirectUri');
const missionSharePointItemIdInput = document.getElementById('defaultSharePointItemId');
const missionSharePointItemETagInput = document.getElementById('defaultSharePointItemETag');
const missionSharePointStatusInput = document.getElementById('sharepoint-sync-status');
const sharePointStatusPanel = document.getElementById('sharepoint-sync-status-panel');
const missionOneDriveTenantIdInput = document.getElementById('defaultOneDriveTenantId');
const missionOneDriveClientIdInput = document.getElementById('defaultOneDriveClientId');
const missionOneDriveFilePathInput = document.getElementById('defaultOneDriveFilePath');
const missionOneDriveRedirectUriInput = document.getElementById('defaultOneDriveRedirectUri');
const missionOneDriveItemIdInput = document.getElementById('defaultOneDriveItemId');
const missionOneDriveItemETagInput = document.getElementById('defaultOneDriveItemETag');
const missionOneDriveStatusInput = document.getElementById('onedrive-sync-status');
const oneDriveStatusPanel = document.getElementById('onedrive-sync-status-panel');
const missionGoogleCloudClientIdInput = document.getElementById('defaultGoogleCloudClientId');
const missionGoogleCloudFileNameInput = document.getElementById('defaultGoogleCloudFileName');
const missionGoogleCloudFolderIdInput = document.getElementById('defaultGoogleCloudFolderId');
const missionGoogleCloudFileIdInput = document.getElementById('defaultGoogleCloudFileId');
const missionGoogleCloudStatusInput = document.getElementById('googlecloud-sync-status');
const googleCloudStatusPanel = document.getElementById('googlecloud-sync-status-panel');
const localSyncSettingsContainer = document.getElementById('local-sync-settings');
const sharePointSyncSettingsContainer = document.getElementById('sharepoint-sync-settings');
const oneDriveSyncSettingsContainer = document.getElementById('onedrive-sync-settings');
const googleCloudSyncSettingsContainer = document.getElementById('googlecloud-sync-settings');
const sharePointConnectButton = document.getElementById('btn-sharepoint-connect');
const sharePointSyncNowButton = document.getElementById('btn-sharepoint-sync-now');
const oneDriveConnectButton = document.getElementById('btn-onedrive-connect');
const oneDriveSyncNowButton = document.getElementById('btn-onedrive-sync-now');
const googleCloudConnectButton = document.getElementById('btn-googlecloud-connect');
const googleCloudSyncNowButton = document.getElementById('btn-googlecloud-sync-now');
const DEFAULTS_STORAGE_KEY = 'themis-ops-mission-defaults';
const MISSIONS_STORAGE_KEY = 'themis-ops-missions';
const MISSION_DATA_HANDLE_DB_NAME = 'themis-ops-mission-data';
const MISSION_DATA_HANDLE_STORE_NAME = 'handles';
const MISSION_DATA_HANDLE_KEY = 'mission-data-handle';
const MISSION_CANONICAL_SCHEMA_VERSION = 1;
const MISSION_CANONICAL_CLIENT_ID_KEY = 'themis-ops-mission-client-id';
const MISSION_CANONICAL_SYNC_POLL_MS = 2000;
const MISSION_SHAREPOINT_SYNC_POLL_MS = 5000;
const MISSION_ONEDRIVE_SYNC_POLL_MS = 5000;
const MISSION_GOOGLECLOUD_SYNC_POLL_MS = 5000;
const SHAREPOINT_CANONICAL_ITEM_TITLE = 'Themis Operations';
const SHAREPOINT_DEFAULT_PAYLOAD_FIELD = 'Payload';
const SHAREPOINT_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const GOOGLE_CLOUD_API_BASE_URL = 'https://www.googleapis.com/drive/v3';
const GOOGLE_CLOUD_UPLOAD_BASE_URL = 'https://www.googleapis.com/upload/drive/v3';
const GOOGLE_CLOUD_CANONICAL_FILE_MIME_TYPE = 'application/json';
const MISSION_ENTRY_ANIMATION_DURATION_MS = 1500;
const MISSION_CARD_PAST_ORDER_OFFSET = 1e15;
const TIMELINE_BAR_LONG_PRESS_DELAY_MS = 500;
let missionDefaults = createEmptyMissionDefaults();
const tabButtons = document.querySelectorAll('.app-tab');
let airportData = {};
let airportDataPromise = null;
let airportLookupEntries = [];
let airportExactLookupIndex = new Map();
let hoverRouteMap = null;
let hoverTooltipToken = 0;
let hoverTooltipPosition = { x: 0, y: 0 };
let activeTooltipMissionId = null;
let tooltipMeasureCtx = null;
let tooltipMeasureFont = '';
let routeLabelMeasureCtx = null;
let routeLabelMeasureFont = '';
const ROUTE_AIRPORT_DOT_RADIUS_PX = 7;
const ROUTE_AIRPORT_DOT_WEIGHT_PX = 1;
const ROUTE_AIRPORT_DOT_EDGE_PX = ROUTE_AIRPORT_DOT_RADIUS_PX + (ROUTE_AIRPORT_DOT_WEIGHT_PX / 2);
const ROUTE_ICAO_LABEL_MIN_WIDTH_PX = 30;
const ROUTE_ICAO_LABEL_HEIGHT_PX = 18;
const ROUTE_ICAO_LABEL_HORIZONTAL_PADDING_PX = 5;
const ROUTE_ICAO_LABEL_CLEARANCE_PX = 8;
const ROUTE_ICAO_LABEL_VERTICAL_GAP_PX = 2;
const ROUTE_ICAO_LABEL_SEARCH_STEP_PX = 2;
const ROUTE_ICAO_LABEL_SEARCH_ANGLE_STEP_DEG = 8;
const ROUTE_ICAO_LABEL_MAX_SEARCH_PX = 48;
const ROUTE_ICAO_LABEL_LINE_CLEARANCE_PX = 2;
const ROUTE_DIRECTION_MARKER_SAFE_RADIUS_PX = 6;
let editingMissionId = null;
let editingMissionBackup = null;
let missionDataFileHandle = null;
let missionDataFileHandleLoadPromise = null;
let missionDataFileWriteQueue = Promise.resolve();
let missionDataFileHandleDisabled = false;
let missionDataFileHandleLoadToken = 0;
let missionCanonicalDocument = null;
let missionSyncMetaById = Object.create(null);
let missionCanonicalSyncTimer = null;
let missionCanonicalWriteInFlight = false;
let missionCanonicalReloadInProgress = false;
let missionCanonicalLocalDirty = false;
let missionCanonicalClientId = null;
let missionCanonicalSyncObserversAttached = false;
let missionCanonicalSyncIntervalMs = 0;
let sharePointAuthClient = null;
let sharePointAuthConfigSignature = '';
let sharePointActiveAccount = null;
let sharePointResolvedListId = '';
let sharePointResolvedListSignature = '';
let sharePointWriteQueue = Promise.resolve();
let sharePointWriteInFlight = false;
let sharePointReloadInProgress = false;
let sharePointSyncTimer = null;
let oneDriveAuthClient = null;
let oneDriveAuthConfigSignature = '';
let oneDriveActiveAccount = null;
let oneDriveResolvedItemId = '';
let oneDriveResolvedItemSignature = '';
let oneDriveWriteQueue = Promise.resolve();
let oneDriveWriteInFlight = false;
let oneDriveReloadInProgress = false;
let oneDriveSyncTimer = null;
let googleCloudAuthClient = null;
let googleCloudAuthConfigSignature = '';
let googleCloudAccessToken = '';
let googleCloudAccessTokenExpiresAt = 0;
let googleCloudTokenRequest = null;
let googleCloudTokenRequestPromise = null;
let googleCloudResolvedFileId = '';
let googleCloudResolvedFileSignature = '';
let googleCloudWriteQueue = Promise.resolve();
let googleCloudWriteInFlight = false;
let googleCloudReloadInProgress = false;
let googleCloudSyncTimer = null;
let currentTimeBubbleStrip = null;
let currentTimeIndicatorRefreshTimer = null;
let missionCardFocusTimeout = null;
let focusedMissionCardId = null;
let touchTimelineBarLongPressTimer = null;
let touchTimelineBarLongPressMissionId = null;
let touchMoved = false;
let touchTapMissionId = null;
let touchPanStartX = 0;
let touchPanStartViewStart = 0;
let touchPanStartViewEnd = 0;
let touchPinchStartDist = 0;
let touchPinchStartMidPct = 0;
let touchPinchStartViewStart = 0;
let touchPinchStartViewEnd = 0;
let touchActiveCount = 0;

function createEmptyMissionDefaults() {
    return {
        homeField: '',
        missionDataPath: '',
        availableTailNumbers: '',
        syncBackend: 'local-file',
        oneDriveTenantId: '',
        oneDriveClientId: '',
        oneDriveFilePath: 'themis_operations_missions.json',
        oneDriveRedirectUri: '',
        oneDriveItemId: '',
        oneDriveItemETag: '',
        googleCloudClientId: '164691062517-j5flaa234p213pdspid9p045jtkh5pp3.apps.googleusercontent.com',
        googleCloudFileName: 'themis_operations_missions.json',
        googleCloudFolderId: '1RBGAjIzZSBomJwMTNipayG5_2AyAng8F',
        googleCloudFileId: '',
        sharePointTenantId: '',
        sharePointClientId: '',
        sharePointSiteId: '',
        sharePointListIdentifier: '',
        sharePointItemTitle: SHAREPOINT_CANONICAL_ITEM_TITLE,
        sharePointPayloadField: SHAREPOINT_DEFAULT_PAYLOAD_FIELD,
        sharePointRedirectUri: '',
        sharePointItemId: '',
        sharePointItemETag: ''
    };
}

function normalizeTailNumbersField(value) {
    if (Array.isArray(value)) {
        return value
            .map(item => normalizeMissionText(item))
            .filter(Boolean)
            .join('\n');
    }

    return typeof value === 'string' ? value.trim() : '';
}

function normalizeMissionSyncBackend(value) {
    if (value === 'google-drive-file' || value === 'google-cloud-file') {
        return 'google-cloud-file';
    }

    return value === 'local-file' || value === 'sharepoint-list' || value === 'onedrive-file'
        ? value
        : 'local-file';
}

function loadMissionDefaults() {
    try {
        const raw = localStorage.getItem(DEFAULTS_STORAGE_KEY);
        if (!raw) return createEmptyMissionDefaults();
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== 'object') return createEmptyMissionDefaults();
        const homeField = typeof parsed.homeField === 'string' ? parsed.homeField : '';
        const missionDataPath = typeof parsed.missionDataPath === 'string' ? parsed.missionDataPath : '';
        const availableTailNumbers = normalizeTailNumbersField(parsed.availableTailNumbers ?? parsed.tailNumbers ?? '');
        const legacyHomeField = typeof parsed.tailNum === 'string' ? parsed.tailNum : '';
        const syncBackend = normalizeMissionSyncBackend(parsed.syncBackend);
        const legacyOneDriveTenantId = typeof parsed.sharePointTenantId === 'string' ? parsed.sharePointTenantId.trim() : '';
        const legacyOneDriveClientId = typeof parsed.sharePointClientId === 'string' ? parsed.sharePointClientId.trim() : '';
        const legacyOneDriveFilePath = '';
        const legacyOneDriveRedirectUri = typeof parsed.sharePointRedirectUri === 'string' ? parsed.sharePointRedirectUri.trim() : '';
        const legacyOneDriveItemId = '';
        const legacyOneDriveItemETag = '';
        const legacyGoogleCloudClientId = typeof parsed.googleDriveClientId === 'string' ? parsed.googleDriveClientId.trim() : '';
        const legacyGoogleCloudFileName = typeof parsed.googleDriveFileName === 'string' && parsed.googleDriveFileName.trim()
            ? parsed.googleDriveFileName.trim()
            : 'themis_operations_missions.json';
        const legacyGoogleCloudFolderId = typeof parsed.googleDriveFolderId === 'string' ? parsed.googleDriveFolderId.trim() : '';
        const legacyGoogleCloudFileId = typeof parsed.googleDriveFileId === 'string' ? parsed.googleDriveFileId.trim() : '';
        const oneDriveTenantId = typeof parsed.oneDriveTenantId === 'string' && parsed.oneDriveTenantId.trim()
            ? parsed.oneDriveTenantId.trim()
            : legacyOneDriveTenantId;
        const oneDriveClientId = typeof parsed.oneDriveClientId === 'string' && parsed.oneDriveClientId.trim()
            ? parsed.oneDriveClientId.trim()
            : legacyOneDriveClientId;
        const oneDriveFilePath = typeof parsed.oneDriveFilePath === 'string' && parsed.oneDriveFilePath.trim()
            ? parsed.oneDriveFilePath.trim()
            : (legacyOneDriveFilePath || 'themis_operations_missions.json');
        const oneDriveRedirectUri = typeof parsed.oneDriveRedirectUri === 'string' && parsed.oneDriveRedirectUri.trim()
            ? parsed.oneDriveRedirectUri.trim()
            : legacyOneDriveRedirectUri;
        const oneDriveItemId = typeof parsed.oneDriveItemId === 'string' && parsed.oneDriveItemId.trim()
            ? parsed.oneDriveItemId.trim()
            : legacyOneDriveItemId;
        const oneDriveItemETag = typeof parsed.oneDriveItemETag === 'string' && parsed.oneDriveItemETag.trim()
            ? parsed.oneDriveItemETag.trim()
            : legacyOneDriveItemETag;
        const googleCloudClientId = typeof parsed.googleCloudClientId === 'string' && parsed.googleCloudClientId.trim()
            ? parsed.googleCloudClientId.trim()
            : legacyGoogleCloudClientId;
        const googleCloudFileName = typeof parsed.googleCloudFileName === 'string' && parsed.googleCloudFileName.trim()
            ? parsed.googleCloudFileName.trim()
            : legacyGoogleCloudFileName;
        const googleCloudFolderId = typeof parsed.googleCloudFolderId === 'string' && parsed.googleCloudFolderId.trim()
            ? parsed.googleCloudFolderId.trim()
            : legacyGoogleCloudFolderId;
        const googleCloudFileId = typeof parsed.googleCloudFileId === 'string' && parsed.googleCloudFileId.trim()
            ? parsed.googleCloudFileId.trim()
            : legacyGoogleCloudFileId;
        const sharePointTenantId = typeof parsed.sharePointTenantId === 'string' ? parsed.sharePointTenantId.trim() : '';
        const sharePointClientId = typeof parsed.sharePointClientId === 'string' ? parsed.sharePointClientId.trim() : '';
        const sharePointSiteId = typeof parsed.sharePointSiteId === 'string' ? parsed.sharePointSiteId.trim() : '';
        const sharePointListIdentifier = typeof parsed.sharePointListIdentifier === 'string' ? parsed.sharePointListIdentifier.trim() : '';
        const sharePointItemTitle = typeof parsed.sharePointItemTitle === 'string' && parsed.sharePointItemTitle.trim()
            ? parsed.sharePointItemTitle.trim()
            : SHAREPOINT_CANONICAL_ITEM_TITLE;
        const sharePointPayloadField = typeof parsed.sharePointPayloadField === 'string' && parsed.sharePointPayloadField.trim()
            ? parsed.sharePointPayloadField.trim()
            : SHAREPOINT_DEFAULT_PAYLOAD_FIELD;
        const sharePointRedirectUri = typeof parsed.sharePointRedirectUri === 'string' ? parsed.sharePointRedirectUri.trim() : '';
        const sharePointItemId = typeof parsed.sharePointItemId === 'string' ? parsed.sharePointItemId.trim() : '';
        const sharePointItemETag = typeof parsed.sharePointItemETag === 'string' ? parsed.sharePointItemETag.trim() : '';
        return {
            homeField: (homeField || legacyHomeField).trim(),
            missionDataPath: missionDataPath.trim(),
            availableTailNumbers,
            syncBackend,
            oneDriveTenantId,
            oneDriveClientId,
            oneDriveFilePath,
            oneDriveRedirectUri,
            oneDriveItemId,
            oneDriveItemETag,
            googleCloudClientId,
            googleCloudFileName,
            googleCloudFolderId,
            googleCloudFileId,
            sharePointTenantId,
            sharePointClientId,
            sharePointSiteId,
            sharePointListIdentifier,
            sharePointItemTitle,
            sharePointPayloadField,
            sharePointRedirectUri,
            sharePointItemId,
            sharePointItemETag
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

function nowMissionTimestamp() {
    return new Date().toISOString();
}

function getMissionSyncClientId() {
    if (missionCanonicalClientId) return missionCanonicalClientId;

    try {
        const stored = localStorage.getItem(MISSION_CANONICAL_CLIENT_ID_KEY);
        if (stored) {
            missionCanonicalClientId = stored;
            return missionCanonicalClientId;
        }

        const generated = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
            ? `ops-${crypto.randomUUID()}`
            : `ops-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;

        localStorage.setItem(MISSION_CANONICAL_CLIENT_ID_KEY, generated);
        missionCanonicalClientId = generated;
        return missionCanonicalClientId;
    } catch {
        missionCanonicalClientId = `ops-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
        return missionCanonicalClientId;
    }
}

function createMissionId(prefix = 'mission') {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}-${crypto.randomUUID()}`;
    }

    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function createEmptyMissionSyncMeta() {
    return {
        updatedAt: '',
        updatedBy: '',
        deletedAt: '',
        deletedBy: ''
    };
}

function normalizeMissionSyncMeta(meta, fallback = {}) {
    const source = meta && typeof meta === 'object' ? meta : {};
    return {
        updatedAt: typeof source.updatedAt === 'string' ? source.updatedAt : (fallback.updatedAt || ''),
        updatedBy: typeof source.updatedBy === 'string' ? source.updatedBy : (fallback.updatedBy || ''),
        deletedAt: typeof source.deletedAt === 'string' ? source.deletedAt : (fallback.deletedAt || ''),
        deletedBy: typeof source.deletedBy === 'string' ? source.deletedBy : (fallback.deletedBy || '')
    };
}

function cloneMissionSyncMetaMap(metaById) {
    const cloned = Object.create(null);
    if (!metaById || typeof metaById !== 'object') return cloned;

    Object.keys(metaById).sort().forEach(id => {
        cloned[String(id)] = normalizeMissionSyncMeta(metaById[id]);
    });

    return cloned;
}

function parseMissionTimestampMs(value) {
    if (!value) return 0;
    const timestamp = new Date(value).getTime();
    return Number.isNaN(timestamp) ? 0 : timestamp;
}

function buildComparableMissionViewRecord(mission) {
    const lifts = getMissionLiftsForRecord(mission);
    const primaryLift = lifts[0] || {
        customer: '',
        pax: '',
        cargo: '',
        hazmat: ''
    };

    return {
        id: String(mission.id),
        missionNum: mission.missionNum || '',
        tailNum: mission.tailNum || '',
        pilot: mission.pilot || '',
        copilot: mission.copilot || '',
        crewChief: mission.crewChief || '',
        loadmaster: mission.loadmaster || '',
        liftCustomer: primaryLift.customer || '',
        liftPax: primaryLift.pax || '',
        liftCargo: primaryLift.cargo || '',
        liftHazmat: primaryLift.hazmat || '',
        lifts: lifts.map(lift => ({
            customer: lift.customer || '',
            pax: lift.pax || '',
            cargo: lift.cargo || '',
            hazmat: lift.hazmat || ''
        })),
        legs: Array.isArray(mission.legs)
            ? mission.legs.map(leg => ({
                takeoffIcao: leg.takeoffIcao || '',
                takeoffTime: getDateTimestamp(leg.takeoffTime),
                landIcao: leg.landIcao || '',
                landTime: getDateTimestamp(leg.landTime)
            }))
            : []
    };
}

function getComparableMissionViewState(document) {
    const normalized = normalizeMissionCanonicalDocument(document);
    const stateById = Object.create(null);

    normalized.missions.forEach(mission => {
        stateById[String(mission.id)] = JSON.stringify(buildComparableMissionViewRecord(mission));
    });

    return stateById;
}

function getMissionSyncMeta(id) {
    const key = String(id);
    return normalizeMissionSyncMeta(missionSyncMetaById[key]);
}

function setMissionSyncMeta(id, updates) {
    const key = String(id);
    const next = normalizeMissionSyncMeta(missionSyncMetaById[key]);
    const source = updates && typeof updates === 'object' ? updates : {};
    missionSyncMetaById[key] = normalizeMissionSyncMeta({
        ...next,
        ...source
    });
    return missionSyncMetaById[key];
}

function markMissionUpdated(id, timestamp = nowMissionTimestamp()) {
    return setMissionSyncMeta(id, {
        updatedAt: timestamp,
        updatedBy: getMissionSyncClientId(),
        deletedAt: '',
        deletedBy: ''
    });
}

function markMissionDeleted(id, timestamp = nowMissionTimestamp()) {
    const current = getMissionSyncMeta(id);
    return setMissionSyncMeta(id, {
        updatedAt: current.updatedAt || timestamp,
        updatedBy: current.updatedBy || getMissionSyncClientId(),
        deletedAt: timestamp,
        deletedBy: getMissionSyncClientId()
    });
}

function normalizeMissionCanonicalDocument(raw, runtime = {}) {
    const source = raw && typeof raw === 'object' ? raw : {};
    const isLegacyArray = Array.isArray(raw);
    const rawMissions = isLegacyArray
        ? raw
        : (Array.isArray(source.missions) ? source.missions : []);
    const rawMeta = !isLegacyArray && source.missionSyncMetaById && typeof source.missionSyncMetaById === 'object'
        ? source.missionSyncMetaById
        : {};
    const fallbackUpdatedAt = typeof source.updatedAt === 'string' && source.updatedAt
        ? source.updatedAt
        : '';
    const fallbackUpdatedBy = typeof source.updatedBy === 'string' && source.updatedBy
        ? source.updatedBy
        : '';

    const normalizedMetaById = Object.create(null);
    Object.keys(rawMeta).sort().forEach(id => {
        normalizedMetaById[String(id)] = normalizeMissionSyncMeta(rawMeta[id], {
            updatedAt: fallbackUpdatedAt,
            updatedBy: fallbackUpdatedBy
        });
    });

    const missionsById = new Map();
    rawMissions.forEach((mission, index) => {
        const normalizedMission = normalizeMissionRecord(mission, index);
        if (!normalizedMission) return;

        const id = String(normalizedMission.id);
        const meta = normalizedMetaById[id] = normalizeMissionSyncMeta(normalizedMetaById[id], {
            updatedAt: fallbackUpdatedAt,
            updatedBy: fallbackUpdatedBy
        });

        if (!meta.updatedAt && !meta.deletedAt) {
            meta.updatedAt = fallbackUpdatedAt;
            meta.updatedBy = fallbackUpdatedBy;
        }

        if (meta.deletedAt) return;
        missionsById.set(id, normalizedMission);
    });

    return {
        schemaVersion: Number(source.schemaVersion) || MISSION_CANONICAL_SCHEMA_VERSION,
        revision: Number(source.revision) || 0,
        updatedAt: typeof source.updatedAt === 'string' ? source.updatedAt : fallbackUpdatedAt,
        updatedBy: typeof source.updatedBy === 'string' ? source.updatedBy : fallbackUpdatedBy,
        missions: [...missionsById.values()].sort((a, b) => String(a.id).localeCompare(String(b.id))),
        missionSyncMetaById: cloneMissionSyncMetaMap(normalizedMetaById),
        sourceSignature: typeof runtime.sourceSignature === 'string' ? runtime.sourceSignature : '',
        sourceLastModified: Number(runtime.sourceLastModified) || 0,
        sourceSize: Number(runtime.sourceSize) || 0,
        sourceWasEmpty: Boolean(runtime.sourceWasEmpty)
    };
}

function buildMissionCanonicalDocument() {
    const now = nowMissionTimestamp();
    const normalizedMissions = [];
    const normalizedMetaById = cloneMissionSyncMetaMap(missionSyncMetaById);
    const activeMissions = Array.isArray(missions) ? missions : [];

    activeMissions.forEach((mission, index) => {
        const normalizedMission = normalizeMissionRecord(mission, index);
        if (!normalizedMission) return;

        const id = String(normalizedMission.id);
        const meta = normalizeMissionSyncMeta(normalizedMetaById[id], {
            updatedAt: now,
            updatedBy: getMissionSyncClientId()
        });
        meta.updatedAt = meta.updatedAt || now;
        meta.updatedBy = meta.updatedBy || getMissionSyncClientId();
        meta.deletedAt = '';
        meta.deletedBy = '';
        normalizedMetaById[id] = meta;
        normalizedMissions.push(normalizedMission);
    });

    normalizedMissions.sort((a, b) => String(a.id).localeCompare(String(b.id)));

    const baseRevision = Number(missionCanonicalDocument && missionCanonicalDocument.revision) || 0;

    return {
        schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
        revision: baseRevision + 1,
        updatedAt: now,
        updatedBy: getMissionSyncClientId(),
        missions: normalizedMissions,
        missionSyncMetaById: normalizedMetaById,
        sourceSignature: missionCanonicalDocument && typeof missionCanonicalDocument.sourceSignature === 'string'
            ? missionCanonicalDocument.sourceSignature
            : '',
        sourceLastModified: missionCanonicalDocument && Number(missionCanonicalDocument.sourceLastModified) ? Number(missionCanonicalDocument.sourceLastModified) : 0,
        sourceSize: missionCanonicalDocument && Number(missionCanonicalDocument.sourceSize) ? Number(missionCanonicalDocument.sourceSize) : 0
    };
}

function serializeMissionCanonicalDocument(document) {
    const source = document && typeof document === 'object' ? document : {};
    return JSON.stringify({
        schemaVersion: Number(source.schemaVersion) || MISSION_CANONICAL_SCHEMA_VERSION,
        revision: Number(source.revision) || 0,
        updatedAt: typeof source.updatedAt === 'string' ? source.updatedAt : nowMissionTimestamp(),
        updatedBy: typeof source.updatedBy === 'string' ? source.updatedBy : getMissionSyncClientId(),
        missions: Array.isArray(source.missions) ? source.missions : [],
        missionSyncMetaById: cloneMissionSyncMetaMap(source.missionSyncMetaById)
    }, null, 2);
}

function getComparableMissionCanonicalPayload(document) {
    const normalized = normalizeMissionCanonicalDocument(document);
    return JSON.stringify({
        missions: normalized.missions,
        missionSyncMetaById: cloneMissionSyncMetaMap(normalized.missionSyncMetaById)
    });
}

function getComparableMissionViewPayload(document) {
    const normalized = normalizeMissionCanonicalDocument(document);
    return JSON.stringify(normalized.missions.map(buildComparableMissionViewRecord));
}

function getMissionSyncAnimationState(previousDocument, nextDocument) {
    const previousViewState = getComparableMissionViewState(previousDocument);
    const nextViewState = getComparableMissionViewState(nextDocument);
    const previousIds = new Set(Object.keys(previousViewState));
    const nextIds = new Set(Object.keys(nextViewState));
    const addedMissionIds = [];
    const removedMissionIds = [];
    const changedMissionIds = [];

    nextIds.forEach(id => {
        if (!previousIds.has(id)) {
            addedMissionIds.push(id);
            return;
        }

        if (previousViewState[id] !== nextViewState[id]) {
            changedMissionIds.push(id);
        }
    });

    previousIds.forEach(id => {
        if (!nextIds.has(id)) {
            removedMissionIds.push(id);
        }
    });

    return {
        shouldAnimateRefresh: addedMissionIds.length > 0 || removedMissionIds.length > 0 || changedMissionIds.length > 0,
        addedMissionIds,
        removedMissionIds,
        changedMissionIds
    };
}

function shouldAnimateMissionSyncRefresh(previousDocument, nextDocument) {
    return getMissionSyncAnimationState(previousDocument, nextDocument).shouldAnimateRefresh;
}

function storeMissionCanonicalDocumentLocally(document) {
    try {
        localStorage.setItem(MISSIONS_STORAGE_KEY, serializeMissionCanonicalDocument(document));
    } catch {
        // Ignore storage failures in file:// or restricted browser contexts.
    }
}

function loadMissionCanonicalDocumentFromCache() {
    try {
        const raw = localStorage.getItem(MISSIONS_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return normalizeMissionCanonicalDocument(parsed);
    } catch {
        return null;
    }
}

function applyMissionCanonicalDocumentToRuntime(document, options = {}) {
    const normalizedDocument = normalizeMissionCanonicalDocument(document, {
        sourceSignature: document && typeof document.sourceSignature === 'string' ? document.sourceSignature : '',
        sourceLastModified: document && Number(document.sourceLastModified) ? Number(document.sourceLastModified) : 0,
        sourceSize: document && Number(document.sourceSize) ? Number(document.sourceSize) : 0
    });

    missions = normalizedDocument.missions.map(cloneMissionRecord);
    missionSyncMetaById = cloneMissionSyncMetaMap(normalizedDocument.missionSyncMetaById);
    missionCanonicalDocument = {
        ...normalizedDocument,
        missions: normalizedDocument.missions.map(cloneMissionRecord),
        missionSyncMetaById: cloneMissionSyncMetaMap(normalizedDocument.missionSyncMetaById)
    };

    if (options.persistLocalCache !== false) {
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
    }

    return missionCanonicalDocument;
}

function mergeMissionCanonicalDocuments(localDocument, remoteDocument) {
    const localDoc = normalizeMissionCanonicalDocument(localDocument);
    const remoteDoc = normalizeMissionCanonicalDocument(remoteDocument);
    const localMissionMap = new Map(localDoc.missions.map(mission => [String(mission.id), mission]));
    const remoteMissionMap = new Map(remoteDoc.missions.map(mission => [String(mission.id), mission]));
    const mergedMetaById = Object.create(null);
    const mergedMissions = [];
    const ids = new Set([
        ...Object.keys(localDoc.missionSyncMetaById || {}),
        ...Object.keys(remoteDoc.missionSyncMetaById || {}),
        ...localMissionMap.keys(),
        ...remoteMissionMap.keys()
    ]);

    ids.forEach(id => {
        const localMeta = normalizeMissionSyncMeta(localDoc.missionSyncMetaById[id]);
        const remoteMeta = normalizeMissionSyncMeta(remoteDoc.missionSyncMetaById[id]);
        const localMission = localMissionMap.get(id) || null;
        const remoteMission = remoteMissionMap.get(id) || null;

        if (localMeta.deletedAt || remoteMeta.deletedAt) {
            const localDeleted = parseMissionTimestampMs(localMeta.deletedAt);
            const remoteDeleted = parseMissionTimestampMs(remoteMeta.deletedAt);
            mergedMetaById[id] = normalizeMissionSyncMeta(
                remoteDeleted >= localDeleted ? remoteMeta : localMeta,
                {
                    updatedAt: remoteDoc.updatedAt || localDoc.updatedAt,
                    updatedBy: remoteDoc.updatedBy || localDoc.updatedBy
                }
            );
            if (!mergedMetaById[id].deletedAt) {
                mergedMetaById[id].deletedAt = nowMissionTimestamp();
            }
            return;
        }

        if (localMission && remoteMission) {
            const localUpdated = parseMissionTimestampMs(localMeta.updatedAt || localDoc.updatedAt);
            const remoteUpdated = parseMissionTimestampMs(remoteMeta.updatedAt || remoteDoc.updatedAt);
            if (remoteUpdated >= localUpdated) {
                mergedMissions.push(cloneMissionRecord(remoteMission));
                mergedMetaById[id] = normalizeMissionSyncMeta(remoteMeta, {
                    updatedAt: remoteDoc.updatedAt,
                    updatedBy: remoteDoc.updatedBy
                });
            } else {
                mergedMissions.push(cloneMissionRecord(localMission));
                mergedMetaById[id] = normalizeMissionSyncMeta(localMeta, {
                    updatedAt: localDoc.updatedAt,
                    updatedBy: localDoc.updatedBy
                });
            }
            return;
        }

        if (remoteMission) {
            mergedMissions.push(cloneMissionRecord(remoteMission));
            mergedMetaById[id] = normalizeMissionSyncMeta(remoteMeta, {
                updatedAt: remoteDoc.updatedAt,
                updatedBy: remoteDoc.updatedBy
            });
            return;
        }

        if (localMission) {
            mergedMissions.push(cloneMissionRecord(localMission));
            mergedMetaById[id] = normalizeMissionSyncMeta(localMeta, {
                updatedAt: localDoc.updatedAt,
                updatedBy: localDoc.updatedBy
            });
            return;
        }

        if (localMeta.updatedAt || localMeta.deletedAt || remoteMeta.updatedAt || remoteMeta.deletedAt) {
            mergedMetaById[id] = normalizeMissionSyncMeta(remoteMeta.updatedAt || remoteMeta.deletedAt ? remoteMeta : localMeta, {
                updatedAt: remoteDoc.updatedAt || localDoc.updatedAt,
                updatedBy: remoteDoc.updatedBy || localDoc.updatedBy
            });
        }
    });

    mergedMissions.sort((a, b) => String(a.id).localeCompare(String(b.id)));

    return {
        schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
        revision: Math.max(Number(localDoc.revision) || 0, Number(remoteDoc.revision) || 0),
        updatedAt: remoteDoc.updatedAt || localDoc.updatedAt || nowMissionTimestamp(),
        updatedBy: remoteDoc.updatedBy || localDoc.updatedBy || getMissionSyncClientId(),
        missions: mergedMissions,
        missionSyncMetaById: mergedMetaById,
        sourceSignature: remoteDoc.sourceSignature || localDoc.sourceSignature || '',
        sourceLastModified: remoteDoc.sourceLastModified || localDoc.sourceLastModified || 0,
        sourceSize: remoteDoc.sourceSize || localDoc.sourceSize || 0
    };
}

async function readMissionCanonicalDocumentFromHandle(handle) {
    if (!handle) return null;

    try {
        const file = await handle.getFile();
        const signature = `${file.lastModified || 0}:${file.size || 0}`;
        const text = await file.text();
        if (!text || !text.trim()) {
            return {
                schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
                revision: 0,
                updatedAt: '',
                updatedBy: '',
                missions: [],
                missionSyncMetaById: Object.create(null),
                sourceSignature: signature,
                sourceLastModified: file.lastModified || 0,
                sourceSize: file.size || 0,
                sourceWasEmpty: true
            };
        }

        const parsed = JSON.parse(text);
        return normalizeMissionCanonicalDocument(parsed, {
            sourceSignature: signature,
            sourceLastModified: file.lastModified || 0,
            sourceSize: file.size || 0,
            sourceWasEmpty: false
        });
    } catch (error) {
        console.warn('Failed to read canonical mission data.', error);
        return null;
    }
}

async function syncMissionCanonicalDocumentFromHandle(options = {}) {
    const { force = false, initializeIfEmpty = false } = options;
    if (missionCanonicalReloadInProgress) return null;

    const handle = await loadMissionDataHandle();
    if (!handle) return null;
    if (!force && (missionCanonicalWriteInFlight || missionCanonicalLocalDirty || editingMissionId != null)) {
        return null;
    }

    missionCanonicalReloadInProgress = true;
    try {
        const remoteDocument = await readMissionCanonicalDocumentFromHandle(handle);
        const remoteWasEmpty = Boolean(remoteDocument && remoteDocument.sourceWasEmpty);
        const localDocument = missionCanonicalDocument || loadMissionCanonicalDocumentFromCache() || buildMissionCanonicalDocument();

        if (!remoteDocument) {
            if (initializeIfEmpty) {
                const initialDocument = localDocument || buildMissionCanonicalDocument();
                if (initialDocument) {
                    await queueMissionDataDiskWrite(initialDocument);
                    applyMissionCanonicalDocumentToRuntime(initialDocument);
                    return initialDocument;
                }
            }
            return null;
        }

        const mergedDocument = localDocument
            ? mergeMissionCanonicalDocuments(localDocument, remoteDocument)
            : remoteDocument;

        const syncAnimationState = getMissionSyncAnimationState(localDocument, mergedDocument);
        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderMissionViewsAfterSync(syncAnimationState);

        if (remoteWasEmpty || (localDocument && getComparableMissionCanonicalPayload(mergedDocument) !== getComparableMissionCanonicalPayload(remoteDocument))) {
            void queueMissionDataDiskWrite(mergedDocument);
        }

        return mergedDocument;
    } finally {
        missionCanonicalReloadInProgress = false;
    }
}

function getMissionSyncMode() {
    return normalizeMissionSyncBackend(missionDefaults.syncBackend);
}

function getMissionSyncPollIntervalMs() {
    if (isSharePointSyncMode()) {
        return MISSION_SHAREPOINT_SYNC_POLL_MS;
    }

    if (isOneDriveSyncMode()) {
        return MISSION_ONEDRIVE_SYNC_POLL_MS;
    }

    if (isGoogleCloudSyncMode()) {
        return MISSION_GOOGLECLOUD_SYNC_POLL_MS;
    }

    return MISSION_CANONICAL_SYNC_POLL_MS;
}

function isSharePointSyncMode() {
    return getMissionSyncMode() === 'sharepoint-list';
}

function isOneDriveSyncMode() {
    return getMissionSyncMode() === 'onedrive-file';
}

function isGoogleCloudSyncMode() {
    return getMissionSyncMode() === 'google-cloud-file';
}

function isLocalFileSyncMode() {
    return getMissionSyncMode() === 'local-file';
}

async function syncMissionCanonicalDocumentFromActiveBackend(options = {}) {
    if (isSharePointSyncMode()) {
        return syncMissionCanonicalDocumentFromSharePointList(options);
    }

    if (isOneDriveSyncMode()) {
        return syncMissionCanonicalDocumentFromOneDriveFile(options);
    }

    if (isGoogleCloudSyncMode()) {
        return syncMissionCanonicalDocumentFromGoogleCloudFile(options);
    }

    return syncMissionCanonicalDocumentFromHandle(options);
}

function startMissionCanonicalSyncLoop() {
    attachMissionCanonicalSyncObservers();

    const intervalMs = getMissionSyncPollIntervalMs();
    if (missionCanonicalSyncTimer && missionCanonicalSyncIntervalMs === intervalMs) return;

    if (missionCanonicalSyncTimer) {
        window.clearInterval(missionCanonicalSyncTimer);
        missionCanonicalSyncTimer = null;
    }

    missionCanonicalSyncIntervalMs = intervalMs;
    missionCanonicalSyncTimer = window.setInterval(() => {
        void syncMissionCanonicalDocumentFromActiveBackend();
    }, intervalMs);
}

function stopMissionCanonicalSyncLoop() {
    if (!missionCanonicalSyncTimer) return;
    window.clearInterval(missionCanonicalSyncTimer);
    missionCanonicalSyncTimer = null;
    missionCanonicalSyncIntervalMs = 0;
}

function attachMissionCanonicalSyncObservers() {
    if (missionCanonicalSyncObserversAttached) return;
    missionCanonicalSyncObserversAttached = true;

    window.addEventListener('focus', () => {
        void syncMissionCanonicalDocumentFromActiveBackend();
    });

    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) {
            void syncMissionCanonicalDocumentFromActiveBackend();
        }
    });

    window.addEventListener('online', () => {
        void syncMissionCanonicalDocumentFromActiveBackend();
    });

    window.addEventListener('pagehide', () => {
        if (missionCanonicalLocalDirty || missionCanonicalWriteInFlight) {
            void queueMissionCanonicalDocumentWrite();
        }
    });
}

function syncMissionDefaultsForm() {
    if (missionHomeFieldInput) missionHomeFieldInput.value = missionDefaults.homeField ?? '';
    if (missionTailNumbersInput) missionTailNumbersInput.value = missionDefaults.availableTailNumbers ?? '';
    if (missionDataPathInput) missionDataPathInput.value = missionDefaults.missionDataPath ?? '';
    if (missionSyncBackendInput) missionSyncBackendInput.value = getMissionSyncMode();
    if (missionOneDriveTenantIdInput) missionOneDriveTenantIdInput.value = missionDefaults.oneDriveTenantId ?? '';
    if (missionOneDriveClientIdInput) missionOneDriveClientIdInput.value = missionDefaults.oneDriveClientId ?? '';
    if (missionOneDriveFilePathInput) missionOneDriveFilePathInput.value = missionDefaults.oneDriveFilePath || 'themis_operations_missions.json';
    if (missionOneDriveRedirectUriInput) missionOneDriveRedirectUriInput.value = missionDefaults.oneDriveRedirectUri ?? '';
    if (missionOneDriveItemIdInput) missionOneDriveItemIdInput.value = missionDefaults.oneDriveItemId ?? '';
    if (missionOneDriveItemETagInput) missionOneDriveItemETagInput.value = missionDefaults.oneDriveItemETag ?? '';
    if (missionSharePointTenantIdInput) missionSharePointTenantIdInput.value = missionDefaults.sharePointTenantId ?? '';
    if (missionSharePointClientIdInput) missionSharePointClientIdInput.value = missionDefaults.sharePointClientId ?? '';
    if (missionSharePointSiteIdInput) missionSharePointSiteIdInput.value = missionDefaults.sharePointSiteId ?? '';
    if (missionSharePointListIdentifierInput) missionSharePointListIdentifierInput.value = missionDefaults.sharePointListIdentifier ?? '';
    if (missionSharePointItemTitleInput) missionSharePointItemTitleInput.value = missionDefaults.sharePointItemTitle ?? SHAREPOINT_CANONICAL_ITEM_TITLE;
    if (missionSharePointPayloadFieldInput) missionSharePointPayloadFieldInput.value = missionDefaults.sharePointPayloadField ?? SHAREPOINT_DEFAULT_PAYLOAD_FIELD;
    if (missionSharePointRedirectUriInput) missionSharePointRedirectUriInput.value = missionDefaults.sharePointRedirectUri ?? '';
    if (missionSharePointItemIdInput) missionSharePointItemIdInput.value = missionDefaults.sharePointItemId ?? '';
    if (missionSharePointItemETagInput) missionSharePointItemETagInput.value = missionDefaults.sharePointItemETag ?? '';
    if (missionGoogleCloudClientIdInput) missionGoogleCloudClientIdInput.value = missionDefaults.googleCloudClientId ?? '';
    if (missionGoogleCloudFileNameInput) missionGoogleCloudFileNameInput.value = missionDefaults.googleCloudFileName || 'themis_operations_missions.json';
    if (missionGoogleCloudFolderIdInput) missionGoogleCloudFolderIdInput.value = missionDefaults.googleCloudFolderId ?? '';
    if (missionGoogleCloudFileIdInput) missionGoogleCloudFileIdInput.value = missionDefaults.googleCloudFileId ?? '';
    toggleMissionSyncSettingsVisibility();
    updateSharePointStatusFromState();
    updateOneDriveStatusFromState();
    updateGoogleCloudStatusFromState();
}

function readMissionDefaultsForm() {
    return {
        ...missionDefaults,
        homeField: missionHomeFieldInput ? missionHomeFieldInput.value.trim() : '',
        missionDataPath: missionDataPathInput ? missionDataPathInput.value.trim() : '',
        availableTailNumbers: missionTailNumbersInput ? missionTailNumbersInput.value.trim() : '',
        syncBackend: missionSyncBackendInput ? normalizeMissionSyncBackend(missionSyncBackendInput.value) : getMissionSyncMode(),
        oneDriveTenantId: missionOneDriveTenantIdInput ? missionOneDriveTenantIdInput.value.trim() : '',
        oneDriveClientId: missionOneDriveClientIdInput ? missionOneDriveClientIdInput.value.trim() : '',
        oneDriveFilePath: missionOneDriveFilePathInput ? (missionOneDriveFilePathInput.value.trim() || 'themis_operations_missions.json') : 'themis_operations_missions.json',
        oneDriveRedirectUri: missionOneDriveRedirectUriInput ? missionOneDriveRedirectUriInput.value.trim() : '',
        oneDriveItemId: missionOneDriveItemIdInput ? missionOneDriveItemIdInput.value.trim() : '',
        oneDriveItemETag: missionOneDriveItemETagInput ? missionOneDriveItemETagInput.value.trim() : '',
        googleCloudClientId: missionGoogleCloudClientIdInput ? missionGoogleCloudClientIdInput.value.trim() : '',
        googleCloudFileName: missionGoogleCloudFileNameInput ? (missionGoogleCloudFileNameInput.value.trim() || 'themis_operations_missions.json') : 'themis_operations_missions.json',
        googleCloudFolderId: missionGoogleCloudFolderIdInput ? missionGoogleCloudFolderIdInput.value.trim() : '',
        googleCloudFileId: missionGoogleCloudFileIdInput ? missionGoogleCloudFileIdInput.value.trim() : '',
        sharePointTenantId: missionSharePointTenantIdInput ? missionSharePointTenantIdInput.value.trim() : '',
        sharePointClientId: missionSharePointClientIdInput ? missionSharePointClientIdInput.value.trim() : '',
        sharePointSiteId: missionSharePointSiteIdInput ? missionSharePointSiteIdInput.value.trim() : '',
        sharePointListIdentifier: missionSharePointListIdentifierInput ? missionSharePointListIdentifierInput.value.trim() : '',
        sharePointItemTitle: missionSharePointItemTitleInput ? missionSharePointItemTitleInput.value.trim() : SHAREPOINT_CANONICAL_ITEM_TITLE,
        sharePointPayloadField: missionSharePointPayloadFieldInput ? missionSharePointPayloadFieldInput.value.trim() : SHAREPOINT_DEFAULT_PAYLOAD_FIELD,
        sharePointRedirectUri: missionSharePointRedirectUriInput ? missionSharePointRedirectUriInput.value.trim() : '',
        sharePointItemId: missionSharePointItemIdInput ? missionSharePointItemIdInput.value.trim() : '',
        sharePointItemETag: missionSharePointItemETagInput ? missionSharePointItemETagInput.value.trim() : ''
    };
}

function toggleMissionSyncSettingsVisibility() {
    const localFileMode = isLocalFileSyncMode();
    const sharePointMode = isSharePointSyncMode();
    const oneDriveMode = isOneDriveSyncMode();
    const googleCloudMode = isGoogleCloudSyncMode();

    if (localSyncSettingsContainer) {
        localSyncSettingsContainer.hidden = !localFileMode;
    }

    if (sharePointSyncSettingsContainer) {
        sharePointSyncSettingsContainer.hidden = !sharePointMode;
    }

    if (oneDriveSyncSettingsContainer) {
        oneDriveSyncSettingsContainer.hidden = !oneDriveMode;
    }

    if (sharePointStatusPanel) {
        sharePointStatusPanel.hidden = !sharePointMode;
    }

    if (oneDriveStatusPanel) {
        oneDriveStatusPanel.hidden = !oneDriveMode;
    }

    if (googleCloudSyncSettingsContainer) {
        googleCloudSyncSettingsContainer.hidden = !googleCloudMode;
    }

    if (googleCloudStatusPanel) {
        googleCloudStatusPanel.hidden = !googleCloudMode;
    }

    const missionDataButton = document.getElementById('btn-choose-mission-data-file');
    if (missionDataPathInput) missionDataPathInput.disabled = !localFileMode;
    if (missionDataButton) missionDataButton.disabled = !localFileMode;
    if (sharePointConnectButton) sharePointConnectButton.disabled = !sharePointMode;
    if (sharePointSyncNowButton) sharePointSyncNowButton.disabled = !sharePointMode;
    if (oneDriveConnectButton) oneDriveConnectButton.disabled = !oneDriveMode;
    if (oneDriveSyncNowButton) oneDriveSyncNowButton.disabled = !oneDriveMode;
    if (googleCloudConnectButton) googleCloudConnectButton.disabled = !googleCloudMode;
    if (googleCloudSyncNowButton) googleCloudSyncNowButton.disabled = !googleCloudMode;
}

function activateMissionSyncBackend() {
    stopMissionCanonicalSyncLoop();

    if (isSharePointSyncMode()) {
        updateSharePointStatusFromState();
        if (hasSharePointSyncConfiguration()) {
            void syncSharePointNow();
        }
        return;
    }

    if (isOneDriveSyncMode()) {
        updateOneDriveStatusFromState();
        if (hasOneDriveSyncConfiguration()) {
            void syncOneDriveNow();
        }
        return;
    }

    if (isGoogleCloudSyncMode()) {
        updateGoogleCloudStatusFromState();
        if (hasGoogleCloudSyncConfiguration()) {
            void syncGoogleCloudNow();
        }
        return;
    }

    updateSharePointStatusFromState();
    updateOneDriveStatusFromState();
    updateGoogleCloudStatusFromState();

    void loadMissionDataHandle().then(handle => {
        if (handle) {
            void syncMissionCanonicalDocumentFromHandle({ force: true, initializeIfEmpty: true })
                .then(() => startMissionCanonicalSyncLoop());
        }
    });
}

function setSharePointStatus(message, state = 'idle') {
    if (!missionSharePointStatusInput) return;
    missionSharePointStatusInput.textContent = message;
    missionSharePointStatusInput.classList.remove('connected', 'error');
    if (state === 'connected') {
        missionSharePointStatusInput.classList.add('connected');
    } else if (state === 'error') {
        missionSharePointStatusInput.classList.add('error');
    }
}

function updateSharePointStatusFromState(message = null, state = 'idle') {
    if (message) {
        setSharePointStatus(message, state);
        return;
    }

    if (!isSharePointSyncMode()) {
        setSharePointStatus('Local file sync');
        return;
    }

    if (missionDefaults.sharePointItemId) {
        setSharePointStatus(`Connected to SharePoint item ${missionDefaults.sharePointItemId}`, 'connected');
        return;
    }

    if (missionDefaults.sharePointSiteId && missionDefaults.sharePointListIdentifier) {
        setSharePointStatus('SharePoint configured, not connected yet');
        return;
    }

    setSharePointStatus('SharePoint config incomplete');
}

function setOneDriveStatus(message, state = 'idle') {
    if (!missionOneDriveStatusInput) return;
    missionOneDriveStatusInput.textContent = message;
    missionOneDriveStatusInput.classList.remove('connected', 'error');
    if (state === 'connected') {
        missionOneDriveStatusInput.classList.add('connected');
    } else if (state === 'error') {
        missionOneDriveStatusInput.classList.add('error');
    }
}

function updateOneDriveStatusFromState(message = null, state = 'idle') {
    if (message) {
        setOneDriveStatus(message, state);
        return;
    }

    if (!isOneDriveSyncMode()) {
        setOneDriveStatus('Local file sync');
        return;
    }

    if (missionDefaults.oneDriveItemId) {
        setOneDriveStatus(`Connected to OneDrive file ${missionDefaults.oneDriveFilePath || missionDefaults.oneDriveItemId}`, 'connected');
        return;
    }

    if (missionDefaults.oneDriveClientId && missionDefaults.oneDriveFilePath) {
        setOneDriveStatus('OneDrive configured, not connected yet');
        return;
    }

    setOneDriveStatus('OneDrive config incomplete');
}

function setGoogleCloudStatus(message, state = 'idle') {
    if (!missionGoogleCloudStatusInput) return;
    missionGoogleCloudStatusInput.textContent = message;
    missionGoogleCloudStatusInput.classList.remove('connected', 'error');
    if (state === 'connected') {
        missionGoogleCloudStatusInput.classList.add('connected');
    } else if (state === 'error') {
        missionGoogleCloudStatusInput.classList.add('error');
    }
}

function updateGoogleCloudStatusFromState(message = null, state = 'idle') {
    if (message) {
        setGoogleCloudStatus(message, state);
        return;
    }

    if (!isGoogleCloudSyncMode()) {
        setGoogleCloudStatus('Local file sync');
        return;
    }

    if (!supportsGoogleCloudSyncTransport()) {
        setGoogleCloudStatus('Google Cloud sync requires a secure browser context.', 'error');
        return;
    }

    if (!getGoogleCloudFolderId()) {
        setGoogleCloudStatus('Google Cloud shared folder ID required.', 'error');
        return;
    }

    if (missionDefaults.googleCloudFileId) {
        setGoogleCloudStatus(`Connected to shared Google Cloud file ${missionDefaults.googleCloudFileName || missionDefaults.googleCloudFileId}`, 'connected');
        return;
    }

    if (missionDefaults.googleCloudClientId && missionDefaults.googleCloudFileName) {
        setGoogleCloudStatus('Google Cloud shared file configured, not connected yet');
        return;
    }

    setGoogleCloudStatus('Google Cloud config incomplete');
}

