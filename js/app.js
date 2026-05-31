﻿let missions = [];
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
        googleCloudFolderId: '',
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

        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderTimeline();
        renderMissionCards();
        if (activeTooltipMissionId != null) {
            refreshTooltipForMission(activeTooltipMissionId);
        }

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

    if (missionDefaults.googleCloudFileId) {
        setGoogleCloudStatus(`Connected to Google Cloud file ${missionDefaults.googleCloudFileName || missionDefaults.googleCloudFileId}`, 'connected');
        return;
    }

    if (missionDefaults.googleCloudClientId && missionDefaults.googleCloudFileName) {
        setGoogleCloudStatus('Google Cloud configured, not connected yet');
        return;
    }

    setGoogleCloudStatus('Google Cloud config incomplete');
}

function getMissionDataFileName(pathValue) {
    const trimmed = (pathValue || '').trim();
    if (!trimmed) return 'themis_operations_missions.json';

    const parts = trimmed.split(/[\\/]/).filter(Boolean);
    return parts.length > 0 ? parts[parts.length - 1] : 'themis_operations_missions.json';
}

function supportsMissionDataFileAccess() {
    return window.isSecureContext && (
        typeof window.showOpenFilePicker === 'function' ||
        typeof window.showSaveFilePicker === 'function'
    );
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
    stopMissionCanonicalSyncLoop();

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
    const pickerTypes = [
        {
            description: 'Mission JSON',
            accept: { 'application/json': ['.json'] }
        }
    ];

    try {
        let handle = null;

        if (typeof window.showOpenFilePicker === 'function') {
            const handles = await window.showOpenFilePicker({
                mode: 'readwrite',
                multiple: false,
                startIn: 'documents',
                types: pickerTypes
            });
            handle = Array.isArray(handles) ? handles[0] : handles;
        } else {
            handle = await window.showSaveFilePicker({
                suggestedName,
                startIn: 'documents',
                types: pickerTypes
            });
        }

        if (!handle) return null;

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
        await syncMissionCanonicalDocumentFromHandle({ force: true, initializeIfEmpty: true });
        startMissionCanonicalSyncLoop();

        return handle;
    } catch (error) {
        if (error && error.name === 'AbortError') return null;
        console.warn('Mission data file selection failed.', error);
        return null;
    }
}

async function writeMissionDataToDisk(document = null) {
    const handle = await loadMissionDataHandle();
    if (!handle) return false;
    const snapshot = normalizeMissionCanonicalDocument(document || missionCanonicalDocument || buildMissionCanonicalDocument());
    const writeToken = missionDataFileHandleLoadToken;
    missionCanonicalWriteInFlight = true;

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

        const latestRemote = await readMissionCanonicalDocumentFromHandle(handle);
        let docToWrite = snapshot;
        if (latestRemote) {
            docToWrite = mergeMissionCanonicalDocuments(latestRemote, snapshot);
        }

        docToWrite = {
            ...docToWrite,
            revision: Math.max(
                Number(docToWrite.revision) || 0,
                Number(snapshot.revision) || 0,
                latestRemote ? Number(latestRemote.revision) || 0 : 0
            ),
            updatedAt: nowMissionTimestamp(),
            updatedBy: getMissionSyncClientId()
        };

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

        await writable.write(serializeMissionCanonicalDocument(docToWrite));
        await writable.close();
        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        return true;
    } catch (error) {
        console.warn('Failed to write mission data to disk.', error);
        return false;
    } finally {
        missionCanonicalWriteInFlight = false;
    }
}

function queueMissionDataDiskWrite(document = null) {
    missionDataFileWriteQueue = missionDataFileWriteQueue
        .then(() => writeMissionDataToDisk(document))
        .catch(error => {
            console.warn('Mission data disk write queue failed.', error);
        });

    return missionDataFileWriteQueue;
}

function queueMissionCanonicalDocumentWrite(document = null) {
    if (isSharePointSyncMode()) {
        if (!hasSharePointSyncConfiguration()) {
            return Promise.resolve(false);
        }
        return queueSharePointListWrite(document);
    }

    if (isOneDriveSyncMode()) {
        if (!hasOneDriveSyncConfiguration()) {
            return Promise.resolve(false);
        }
        return queueOneDriveFileWrite(document);
    }

    if (isGoogleCloudSyncMode()) {
        if (!hasGoogleCloudSyncConfiguration()) {
            return Promise.resolve(false);
        }
        return queueGoogleCloudFileWrite(document);
    }

    return queueMissionDataDiskWrite(document);
}

function supportsOneDriveSyncTransport() {
    return window.isSecureContext && typeof fetch === 'function' && typeof msal !== 'undefined';
}

function getOneDriveTenantId() {
    return missionDefaults.oneDriveTenantId ? missionDefaults.oneDriveTenantId.trim() : '';
}

function getOneDriveClientId() {
    return missionDefaults.oneDriveClientId ? missionDefaults.oneDriveClientId.trim() : '';
}

function getOneDriveFilePath() {
    const configured = missionDefaults.oneDriveFilePath ? missionDefaults.oneDriveFilePath.trim() : '';
    return normalizeOneDriveFilePath(configured || 'themis_operations_missions.json');
}

function getOneDriveRedirectUri() {
    const configured = missionDefaults.oneDriveRedirectUri ? missionDefaults.oneDriveRedirectUri.trim() : '';
    if (configured) return configured;
    return window.location.href.split('#')[0];
}

function getOneDriveConfigSignature() {
    return [
        getOneDriveTenantId(),
        getOneDriveClientId(),
        getOneDriveRedirectUri()
    ].join('|');
}

function hasOneDriveSyncConfiguration() {
    return Boolean(getOneDriveClientId() && getOneDriveFilePath());
}

function getOneDriveAuthority() {
    const tenantId = getOneDriveTenantId();
    return `https://login.microsoftonline.com/${tenantId || 'common'}`;
}

function getOneDriveScopes() {
    return ['Files.ReadWrite'];
}

function resetOneDriveAuthClient() {
    oneDriveAuthClient = null;
    oneDriveAuthConfigSignature = '';
    oneDriveActiveAccount = null;
}

function getOneDriveResolutionSignature() {
    return [getOneDriveFilePath()].join('|');
}

function normalizeOneDriveFilePath(pathValue) {
    if (typeof pathValue !== 'string') return '';
    return pathValue
        .trim()
        .replace(/\\/g, '/')
        .replace(/^\/+/, '')
        .replace(/\/{2,}/g, '/')
        .replace(/\/+$/, '');
}

function encodeOneDrivePathSegments(pathValue) {
    const normalized = normalizeOneDriveFilePath(pathValue);
    if (!normalized) return '';
    return normalized
        .split('/')
        .filter(Boolean)
        .map(segment => encodeURIComponent(segment))
        .join('/');
}

function buildOneDriveItemPathByPath(pathValue) {
    const encodedPath = encodeOneDrivePathSegments(pathValue);
    return encodedPath ? `/me/drive/root:/${encodedPath}` : null;
}

function buildOneDriveItemContentPathByPath(pathValue) {
    const itemPath = buildOneDriveItemPathByPath(pathValue);
    return itemPath ? `${itemPath}:/content` : null;
}

function buildOneDriveItemPathById(itemId) {
    const encodedItemId = encodeURIComponent(String(itemId));
    return `/me/drive/items/${encodedItemId}`;
}

function buildOneDriveItemContentPathById(itemId) {
    return `${buildOneDriveItemPathById(itemId)}/content`;
}

function updateOneDriveItemReference(record) {
    const itemId = record && record.id != null ? String(record.id) : '';
    const etag = record && typeof (record.eTag || record['@odata.etag'] || record.etag) === 'string'
        ? String(record.eTag || record['@odata.etag'] || record.etag)
        : '';

    oneDriveResolvedItemId = itemId;
    oneDriveResolvedItemSignature = getOneDriveResolutionSignature();

    if ((missionDefaults.oneDriveItemId || '') !== itemId) {
        missionDefaults.oneDriveItemId = itemId;
    }

    if ((missionDefaults.oneDriveItemETag || '') !== etag) {
        missionDefaults.oneDriveItemETag = etag;
    }

    persistMissionDefaults();
    if (missionOneDriveItemIdInput) missionOneDriveItemIdInput.value = missionDefaults.oneDriveItemId ?? '';
    if (missionOneDriveItemETagInput) missionOneDriveItemETagInput.value = missionDefaults.oneDriveItemETag ?? '';
    updateOneDriveStatusFromState();
}

function clearOneDriveItemReference() {
    oneDriveResolvedItemId = '';
    oneDriveResolvedItemSignature = '';

    if (missionDefaults.oneDriveItemId) {
        missionDefaults.oneDriveItemId = '';
    }

    if (missionDefaults.oneDriveItemETag) {
        missionDefaults.oneDriveItemETag = '';
    }
}

async function getOneDriveAuthClient() {
    const signature = getOneDriveConfigSignature();
    if (oneDriveAuthClient && oneDriveAuthConfigSignature === signature) {
        return oneDriveAuthClient;
    }

    if (!supportsOneDriveSyncTransport()) {
        throw new Error('OneDrive sync requires Microsoft auth support in a secure browser context.');
    }

    const clientId = getOneDriveClientId();
    if (!clientId) {
        throw new Error('OneDrive client ID is required.');
    }

    oneDriveAuthClient = new msal.PublicClientApplication({
        auth: {
            clientId,
            authority: getOneDriveAuthority(),
            redirectUri: getOneDriveRedirectUri(),
            navigateToLoginRequestUrl: false
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false
        }
    });
    oneDriveAuthConfigSignature = signature;
    oneDriveActiveAccount = oneDriveAuthClient.getAllAccounts()[0] || null;
    return oneDriveAuthClient;
}

async function getOneDriveAccessToken(options = {}) {
    const interactive = Boolean(options.interactive);
    const client = await getOneDriveAuthClient();
    const scopes = getOneDriveScopes();
    const account = oneDriveActiveAccount || client.getAllAccounts()[0] || null;

    if (account) {
        try {
            const tokenResult = await client.acquireTokenSilent({
                scopes,
                account
            });
            oneDriveActiveAccount = tokenResult.account || account;
            return tokenResult.accessToken;
        } catch (error) {
            if (!interactive) return null;
        }
    }

    if (!interactive) return null;

    const loginResult = await client.loginPopup({
        scopes,
        prompt: 'select_account'
    });
    oneDriveActiveAccount = loginResult.account || null;
    if (loginResult.accessToken) return loginResult.accessToken;

    const retryResult = await client.acquireTokenSilent({
        scopes,
        account: oneDriveActiveAccount
    });
    return retryResult.accessToken;
}

async function graphOneDriveRequest(path, options = {}) {
    const token = await getOneDriveAccessToken({ interactive: Boolean(options.interactive) });
    if (!token) return null;

    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    if (options.body != null && !headers.has('Content-Type') && options.method && options.method !== 'GET' && options.method !== 'HEAD') {
        headers.set('Content-Type', 'application/json');
    }

    const response = await fetch(`${SHAREPOINT_GRAPH_BASE_URL}${path}`, {
        method: options.method || 'GET',
        headers,
        body: options.body,
        cache: 'no-store'
    });

    const responseText = await response.text();
    let responseBody = null;
    if (responseText && options.responseType !== 'text') {
        try {
            responseBody = JSON.parse(responseText);
        } catch {
            responseBody = responseText;
        }
    }

    if (!response.ok) {
        const error = new Error(graphErrorMessage({ body: responseBody, message: response.statusText }, `OneDrive request failed with status ${response.status}.`));
        error.status = response.status;
        error.body = responseBody;
        throw error;
    }

    if (options.responseType === 'text') {
        return responseText;
    }

    if (responseBody == null && response.status === 204) {
        return {};
    }

    if (responseBody == null && responseText) {
        try {
            return JSON.parse(responseText);
        } catch {
            return responseText;
        }
    }

    return responseBody;
}

async function findOneDriveCanonicalItemRecord(options = {}) {
    const interactive = Boolean(options.interactive);
    const storedItemId = missionDefaults.oneDriveItemId ? missionDefaults.oneDriveItemId.trim() : '';
    const resolutionSignature = getOneDriveResolutionSignature();
    const filePath = getOneDriveFilePath();

    if (!hasOneDriveSyncConfiguration()) return null;

    if (storedItemId && oneDriveResolvedItemId === storedItemId && oneDriveResolvedItemSignature === resolutionSignature) {
        try {
            const itemPath = buildOneDriveItemPathById(storedItemId);
            const record = await graphOneDriveRequest(itemPath, {
                interactive
            });
            if (record) return record;
        } catch (error) {
            if (error && error.status !== 404) {
                throw error;
            }
        }
    }

    if (storedItemId) {
        try {
            const itemPath = buildOneDriveItemPathById(storedItemId);
            const record = await graphOneDriveRequest(itemPath, {
                interactive
            });
            if (record) {
                oneDriveResolvedItemId = String(record.id || storedItemId);
                oneDriveResolvedItemSignature = resolutionSignature;
                return record;
            }
        } catch (error) {
            if (error && error.status !== 404) {
                throw error;
            }
        }
    }

    if (!filePath) return null;

    const itemPath = buildOneDriveItemPathByPath(filePath);
    if (!itemPath) return null;

    const record = await graphOneDriveRequest(itemPath, {
        interactive
    });

    if (record) {
        oneDriveResolvedItemId = String(record.id || '');
        oneDriveResolvedItemSignature = resolutionSignature;
    }

    return record || null;
}

async function parseOneDriveCanonicalDocumentFromRecord(record) {
    const itemId = record && record.id != null ? String(record.id) : '';
    const fileName = record && typeof record.name === 'string' && record.name.trim()
        ? record.name.trim()
        : getOneDriveFilePath();

    if (!itemId) {
        return normalizeMissionCanonicalDocument({
            schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
            revision: 0,
            updatedAt: '',
            updatedBy: '',
            missions: [],
            missionSyncMetaById: Object.create(null)
        }, {
            sourceSignature: `onedrive:${getOneDriveFilePath()}:unknown`,
            sourceWasEmpty: true
        });
    }

    const payload = await graphOneDriveRequest(buildOneDriveItemContentPathById(itemId), {
        interactive: false,
        responseType: 'text'
    });

    if (payload == null) {
        throw new Error(`OneDrive file content could not be read for file ${fileName || itemId}.`);
    }

    if (typeof payload !== 'string') {
        throw new Error(`OneDrive file content could not be read for file ${fileName || itemId}.`);
    }

    if (!payload.trim()) {
        return normalizeMissionCanonicalDocument({
            schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
            revision: 0,
            updatedAt: '',
            updatedBy: '',
            missions: [],
            missionSyncMetaById: Object.create(null)
        }, {
            sourceSignature: `onedrive:${getOneDriveFilePath()}:${itemId}`,
            sourceWasEmpty: true
        });
    }

    try {
        const parsed = JSON.parse(payload);
        return normalizeMissionCanonicalDocument(parsed, {
            sourceSignature: `onedrive:${getOneDriveFilePath()}:${itemId}`,
            sourceLastModified: 0,
            sourceSize: payload.length,
            sourceWasEmpty: false
        });
    } catch {
        throw new Error(`OneDrive payload is not valid JSON for file ${fileName || itemId}.`);
    }
}

async function createOneDriveCanonicalItemRecord(document) {
    const contentPath = buildOneDriveItemContentPathByPath(getOneDriveFilePath());
    if (!contentPath) return null;

    return graphOneDriveRequest(contentPath, {
        method: 'PUT',
        interactive: false,
        body: serializeMissionCanonicalDocument(document)
    });
}

async function updateOneDriveCanonicalItemRecord(itemId, document, etag = '') {
    const headers = {};
    headers['If-Match'] = etag || '*';

    const contentPath = buildOneDriveItemContentPathById(itemId);
    if (!contentPath) return null;

    const updated = await graphOneDriveRequest(contentPath, {
        method: 'PUT',
        interactive: false,
        headers,
        body: serializeMissionCanonicalDocument(document)
    });

    return updated;
}

async function writeOneDriveCanonicalDocument(document = null) {
    if (!hasOneDriveSyncConfiguration()) return false;
    if (oneDriveWriteInFlight) return false;

    const snapshot = normalizeMissionCanonicalDocument(document || missionCanonicalDocument || buildMissionCanonicalDocument());
    oneDriveWriteInFlight = true;

    try {
        const latestRemoteRecord = await findOneDriveCanonicalItemRecord({ interactive: false });
        let docToWrite = snapshot;
        let remoteRecord = latestRemoteRecord || null;

        if (remoteRecord) {
            const remoteDocument = await parseOneDriveCanonicalDocumentFromRecord(remoteRecord);
            docToWrite = mergeMissionCanonicalDocuments(remoteDocument, snapshot);
        }

        docToWrite = {
            ...docToWrite,
            revision: Math.max(Number(docToWrite.revision) || 0, Number(snapshot.revision) || 0),
            updatedAt: nowMissionTimestamp(),
            updatedBy: getMissionSyncClientId()
        };

        let updatedRecord = null;
        if (remoteRecord && remoteRecord.id != null) {
            try {
                updatedRecord = await updateOneDriveCanonicalItemRecord(remoteRecord.id, docToWrite, remoteRecord.eTag || remoteRecord['@odata.etag'] || remoteRecord.etag || missionDefaults.oneDriveItemETag || '');
            } catch (error) {
                if (error && (error.status === 409 || error.status === 412)) {
                    const retryRecord = await findOneDriveCanonicalItemRecord({ interactive: false });
                    if (retryRecord && retryRecord.id != null) {
                        const retryDocument = await parseOneDriveCanonicalDocumentFromRecord(retryRecord);
                        docToWrite = mergeMissionCanonicalDocuments(retryDocument, snapshot);
                        updatedRecord = await updateOneDriveCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord.eTag || retryRecord['@odata.etag'] || retryRecord.etag || '');
                    } else {
                        throw error;
                    }
                } else {
                    throw error;
                }
            }
        } else {
            try {
                updatedRecord = await createOneDriveCanonicalItemRecord(docToWrite);
            } catch (error) {
                if (error && (error.status === 409 || error.status === 412)) {
                    const retryRecord = await findOneDriveCanonicalItemRecord({ interactive: false });
                    if (retryRecord && retryRecord.id != null) {
                        const retryDocument = await parseOneDriveCanonicalDocumentFromRecord(retryRecord);
                        docToWrite = mergeMissionCanonicalDocuments(retryDocument, snapshot);
                        updatedRecord = await updateOneDriveCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord.eTag || retryRecord['@odata.etag'] || retryRecord.etag || '');
                    } else {
                        throw error;
                    }
                } else {
                    throw error;
                }
            }
        }

        if (!updatedRecord) {
            const retryRecord = await findOneDriveCanonicalItemRecord({ interactive: false });
            if (retryRecord && retryRecord.id != null) {
                updatedRecord = await updateOneDriveCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord.eTag || retryRecord['@odata.etag'] || retryRecord.etag || '');
            }
        }

        if (!updatedRecord) {
            throw new Error('OneDrive file could not be created or updated.');
        }

        updateOneDriveItemReference(updatedRecord);

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        updateOneDriveStatusFromState();
        return true;
    } catch (error) {
        console.warn('Failed to write mission data to OneDrive.', error);
        setOneDriveStatus(graphErrorMessage(error, 'OneDrive write failed.'), 'error');
        return false;
    } finally {
        oneDriveWriteInFlight = false;
    }
}

function queueOneDriveFileWrite(document = null) {
    oneDriveWriteQueue = oneDriveWriteQueue
        .then(() => writeOneDriveCanonicalDocument(document))
        .catch(error => {
            console.warn('OneDrive write queue failed.', error);
        });

    return oneDriveWriteQueue;
}

async function syncMissionCanonicalDocumentFromOneDriveFile(options = {}) {
    const { force = false, initializeIfEmpty = false, interactive = false, skipAccessTokenCheck = false } = options;
    if (oneDriveReloadInProgress) return null;
    if (!hasOneDriveSyncConfiguration()) {
        updateOneDriveStatusFromState();
        return null;
    }

    if (!skipAccessTokenCheck) {
        const accessToken = await getOneDriveAccessToken({ interactive });
        if (!accessToken) {
            updateOneDriveStatusFromState();
            return null;
        }
    }

    if (!force && (oneDriveWriteInFlight || missionCanonicalLocalDirty || editingMissionId != null)) {
        return null;
    }

    oneDriveReloadInProgress = true;
    try {
        const remoteRecord = await findOneDriveCanonicalItemRecord({ interactive });
        const remoteWasEmpty = Boolean(remoteRecord && remoteRecord.size === 0);
        const localDocument = missionCanonicalDocument || loadMissionCanonicalDocumentFromCache() || buildMissionCanonicalDocument();

        if (!remoteRecord) {
            if (initializeIfEmpty) {
                const initialDocument = localDocument || buildMissionCanonicalDocument();
                if (initialDocument) {
                    const initialized = await queueOneDriveFileWrite(initialDocument);
                    if (initialized) {
                        applyMissionCanonicalDocumentToRuntime(initialDocument);
                        updateOneDriveStatusFromState();
                        return initialDocument;
                    }
                }
            }
            return null;
        }

        updateOneDriveItemReference(remoteRecord);
        const remoteDocument = await parseOneDriveCanonicalDocumentFromRecord(remoteRecord);
        const mergedDocument = localDocument
            ? mergeMissionCanonicalDocuments(localDocument, remoteDocument)
            : remoteDocument;

        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderTimeline();
        renderMissionCards();
        if (activeTooltipMissionId != null) {
            refreshTooltipForMission(activeTooltipMissionId);
        }

        if (remoteWasEmpty || (localDocument && getComparableMissionCanonicalPayload(mergedDocument) !== getComparableMissionCanonicalPayload(remoteDocument))) {
            void queueOneDriveFileWrite(mergedDocument);
        }

        updateOneDriveStatusFromState();
        return mergedDocument;
    } catch (error) {
        console.warn('Failed to sync mission data from OneDrive.', error);
        setOneDriveStatus(graphErrorMessage(error, 'OneDrive sync failed.'), 'error');
        return null;
    } finally {
        oneDriveReloadInProgress = false;
    }
}

async function requestOneDriveConnection() {
    if (!isOneDriveSyncMode()) {
        setOneDriveStatus('Switch to OneDrive sync first.');
        return null;
    }

    if (!hasOneDriveSyncConfiguration()) {
        setOneDriveStatus('OneDrive config is incomplete.', 'error');
        return null;
    }

    try {
        setOneDriveStatus('Connecting to OneDrive...');
        await getOneDriveAccessToken({ interactive: true });
        const synced = await syncMissionCanonicalDocumentFromOneDriveFile({
            force: true,
            initializeIfEmpty: true,
            interactive: true,
            skipAccessTokenCheck: true
        });
        updateOneDriveStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('OneDrive connection failed.', error);
        setOneDriveStatus(graphErrorMessage(error, 'OneDrive connection failed.'), 'error');
        return null;
    }
}

async function syncOneDriveNow() {
    if (!isOneDriveSyncMode()) return null;
    if (!hasOneDriveSyncConfiguration()) {
        setOneDriveStatus('OneDrive config is incomplete.', 'error');
        return null;
    }

    try {
        const accessToken = await getOneDriveAccessToken({ interactive: false });
        if (!accessToken) {
            updateOneDriveStatusFromState();
            return null;
        }
        const synced = await syncMissionCanonicalDocumentFromOneDriveFile({
            force: true,
            initializeIfEmpty: true,
            skipAccessTokenCheck: true
        });
        updateOneDriveStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('OneDrive sync failed.', error);
        setOneDriveStatus(graphErrorMessage(error, 'OneDrive sync failed.'), 'error');
        return null;
    }
}

function supportsGoogleCloudSyncTransport() {
    return window.isSecureContext
        && typeof fetch === 'function'
        && typeof google !== 'undefined'
        && google.accounts
        && google.accounts.oauth2
        && typeof google.accounts.oauth2.initTokenClient === 'function';
}

function getGoogleCloudClientId() {
    return missionDefaults.googleCloudClientId ? missionDefaults.googleCloudClientId.trim() : '';
}

function getGoogleCloudFileName() {
    const configured = missionDefaults.googleCloudFileName ? missionDefaults.googleCloudFileName.trim() : '';
    return configured || 'themis_operations_missions.json';
}

function getGoogleCloudFolderId() {
    return missionDefaults.googleCloudFolderId ? missionDefaults.googleCloudFolderId.trim() : '';
}

function getGoogleCloudAuthConfigSignature() {
    return getGoogleCloudClientId();
}

function getGoogleCloudResolutionSignature() {
    return [getGoogleCloudFileName(), getGoogleCloudFolderId()].join('|');
}

function hasGoogleCloudSyncConfiguration() {
    return Boolean(getGoogleCloudClientId() && getGoogleCloudFileName());
}

function getGoogleCloudScopes() {
    return ['https://www.googleapis.com/auth/drive'];
}

function getGoogleCloudAuthErrorMessage(response, fallback = 'Google Cloud authorization failed.') {
    const errorText = [
        response && typeof response.error === 'string' ? response.error : '',
        response && typeof response.error_description === 'string' ? response.error_description : '',
        response && typeof response.message === 'string' ? response.message : ''
    ].join(' ').toLowerCase();

    if (errorText.includes('invalid_client') || errorText.includes('origin_mismatch')) {
        return 'Google rejected the OAuth client. Use a Google OAuth 2.0 Web browser client ID and add this page origin to Authorized JavaScript origins in Google Cloud.';
    }

    if (errorText.includes('deleted_client')) {
        return 'The Google OAuth client was deleted. Create or restore the client ID in Google Cloud.';
    }

    if (errorText.includes('org_internal')) {
        return 'This Google OAuth client is restricted to a Google Workspace organization.';
    }

    return response && typeof response.error_description === 'string' && response.error_description.trim()
        ? response.error_description.trim()
        : (response && typeof response.error === 'string' && response.error.trim()
            ? response.error.trim()
            : fallback);
}

function resetGoogleCloudAuthClient() {
    googleCloudAuthClient = null;
    googleCloudAuthConfigSignature = '';
    googleCloudAccessToken = '';
    googleCloudAccessTokenExpiresAt = 0;
    if (googleCloudTokenRequest) {
        const pending = googleCloudTokenRequest;
        googleCloudTokenRequest = null;
        googleCloudTokenRequestPromise = null;
        try {
            pending.reject(new Error('Google Cloud authentication was reset.'));
        } catch {
            // Ignore promise rejection handling failures.
        }
    }
}

function buildGoogleCloudFileResourceFields() {
    return 'id,name,modifiedTime,version,md5Checksum,size,parents';
}

function buildGoogleCloudFilesPath() {
    return '/files';
}

function buildGoogleCloudFileMetadataPathById(fileId) {
    const query = new URLSearchParams();
    query.set('supportsAllDrives', 'true');
    query.set('fields', buildGoogleCloudFileResourceFields());
    return `/files/${encodeURIComponent(String(fileId))}?${query.toString()}`;
}

function buildGoogleCloudFileDownloadPathById(fileId) {
    const query = new URLSearchParams();
    query.set('supportsAllDrives', 'true');
    query.set('alt', 'media');
    return `/files/${encodeURIComponent(String(fileId))}?${query.toString()}`;
}

function buildGoogleCloudFileUploadPathById(fileId) {
    const query = new URLSearchParams();
    query.set('supportsAllDrives', 'true');
    query.set('uploadType', 'media');
    query.set('fields', buildGoogleCloudFileResourceFields());
    return `/files/${encodeURIComponent(String(fileId))}?${query.toString()}`;
}

function buildGoogleCloudFileCreatePath() {
    const query = new URLSearchParams();
    query.set('supportsAllDrives', 'true');
    query.set('fields', buildGoogleCloudFileResourceFields());
    return `/files?${query.toString()}`;
}

function escapeGoogleCloudQueryString(value) {
    return String(value)
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'");
}

function buildGoogleCloudFileSearchQuery() {
    const clauses = [
        `name = '${escapeGoogleCloudQueryString(getGoogleCloudFileName())}'`,
        'trashed = false'
    ];

    const folderId = getGoogleCloudFolderId();
    if (folderId) {
        clauses.push(`'${escapeGoogleCloudQueryString(folderId)}' in parents`);
    }

    return clauses.join(' and ');
}

async function getGoogleCloudAuthClient() {
    const signature = getGoogleCloudAuthConfigSignature();
    if (googleCloudAuthClient && googleCloudAuthConfigSignature === signature) {
        return googleCloudAuthClient;
    }

    if (!supportsGoogleCloudSyncTransport()) {
        throw new Error('Google Cloud sync requires Google Identity Services in a secure browser context.');
    }

    const clientId = getGoogleCloudClientId();
    if (!clientId) {
        throw new Error('Google Cloud client ID is required.');
    }

    googleCloudAuthClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: getGoogleCloudScopes().join(' '),
        callback: response => {
            const pending = googleCloudTokenRequest;
            googleCloudTokenRequest = null;
            googleCloudTokenRequestPromise = null;

            if (response && response.access_token) {
                googleCloudAccessToken = response.access_token;
                googleCloudAccessTokenExpiresAt = Date.now() + ((Number(response.expires_in) || 0) * 1000);
                if (pending) {
                    pending.resolve(response.access_token);
                }
                return;
            }

            googleCloudAccessToken = '';
            googleCloudAccessTokenExpiresAt = 0;
            const message = getGoogleCloudAuthErrorMessage(response, 'Google Cloud access token request failed.');
            if (pending) {
                pending.reject(new Error(message));
            }
        }
    });
    googleCloudAuthConfigSignature = signature;
    googleCloudAccessToken = '';
    googleCloudAccessTokenExpiresAt = 0;
    return googleCloudAuthClient;
}

async function getGoogleCloudAccessToken(options = {}) {
    const interactive = Boolean(options.interactive);
    const cachedTokenIsValid = googleCloudAccessToken && googleCloudAccessTokenExpiresAt > Date.now() + 60000;
    if (cachedTokenIsValid) {
        return googleCloudAccessToken;
    }

    if (googleCloudTokenRequestPromise) {
        return googleCloudTokenRequestPromise;
    }

    const client = await getGoogleCloudAuthClient();

    let resolveToken;
    let rejectToken;
    googleCloudTokenRequestPromise = new Promise((resolve, reject) => {
        resolveToken = resolve;
        rejectToken = reject;
    });
    googleCloudTokenRequest = { resolve: resolveToken, reject: rejectToken };

    try {
        client.requestAccessToken({
            prompt: interactive ? 'consent' : ''
        });
    } catch (error) {
        googleCloudTokenRequest = null;
        googleCloudTokenRequestPromise = null;
        rejectToken(error);
    }

    return googleCloudTokenRequestPromise;
}

async function googleCloudRequest(path, options = {}) {
    const token = await getGoogleCloudAccessToken({ interactive: Boolean(options.interactive) });
    if (!token) return null;

    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    if (options.body != null && !headers.has('Content-Type') && options.method && options.method !== 'GET' && options.method !== 'HEAD') {
        headers.set('Content-Type', 'application/json; charset=utf-8');
    }

    const response = await fetch(`${options.baseUrl || GOOGLE_CLOUD_API_BASE_URL}${path}`, {
        method: options.method || 'GET',
        headers,
        body: options.body,
        cache: 'no-store'
    });

    const responseText = await response.text();
    let responseBody = null;
    if (responseText && options.responseType !== 'text') {
        try {
            responseBody = JSON.parse(responseText);
        } catch {
            responseBody = responseText;
        }
    }

    if (!response.ok) {
        if (response.status === 401 || response.status === 403) {
            googleCloudAccessToken = '';
            googleCloudAccessTokenExpiresAt = 0;
        }
        const error = new Error(graphErrorMessage({ body: responseBody, message: response.statusText }, `Google Cloud request failed with status ${response.status}.`));
        error.status = response.status;
        error.body = responseBody;
        throw error;
    }

    if (options.responseType === 'text') {
        return responseText;
    }

    if (responseBody == null && response.status === 204) {
        return {};
    }

    if (responseBody == null && responseText) {
        try {
            return JSON.parse(responseText);
        } catch {
            return responseText;
        }
    }

    return responseBody;
}

function updateGoogleCloudItemReference(record) {
    const fileId = record && record.id != null ? String(record.id) : '';
    let changed = false;

    if ((missionDefaults.googleCloudFileId || '') !== fileId) {
        missionDefaults.googleCloudFileId = fileId;
        changed = true;
    }

    if (fileId) {
        googleCloudResolvedFileId = fileId;
        googleCloudResolvedFileSignature = getGoogleCloudResolutionSignature();
    }

    if (changed) {
        syncMissionDefaultsForm();
        persistMissionDefaults();
    }
}

function clearGoogleCloudItemReference() {
    let changed = false;
    if (missionDefaults.googleCloudFileId) {
        missionDefaults.googleCloudFileId = '';
        changed = true;
    }

    googleCloudResolvedFileId = '';
    googleCloudResolvedFileSignature = '';

    if (changed) {
        syncMissionDefaultsForm();
        persistMissionDefaults();
    }
}

async function findGoogleCloudCanonicalItemRecord(options = {}) {
    const interactive = Boolean(options.interactive);
    const targetName = getGoogleCloudFileName();
    const normalizedTargetName = targetName.trim().toLowerCase();
    const storedFileId = missionDefaults.googleCloudFileId ? missionDefaults.googleCloudFileId.trim() : '';
    const resolutionSignature = getGoogleCloudResolutionSignature();

    if (!hasGoogleCloudSyncConfiguration()) return null;

    if (storedFileId) {
        try {
            const record = await googleCloudRequest(buildGoogleCloudFileMetadataPathById(storedFileId), {
                interactive
            });
            if (record) {
                googleCloudResolvedFileId = String(record.id || storedFileId);
                googleCloudResolvedFileSignature = resolutionSignature;
                return record;
            }
        } catch (error) {
            if (!error || error.status !== 404) {
                throw error;
            }
        }
    }

    const query = new URLSearchParams();
    query.set('q', buildGoogleCloudFileSearchQuery());
    query.set('pageSize', '10');
    query.set('fields', `files(${buildGoogleCloudFileResourceFields()}),nextPageToken`);
    query.set('orderBy', 'modifiedTime desc');
    // Search across My Drive and shared drives so every device can resolve the same canonical file.
    query.set('corpora', 'allDrives');
    query.set('spaces', 'drive');
    query.set('supportsAllDrives', 'true');
    query.set('includeItemsFromAllDrives', 'true');

    const listing = await googleCloudRequest(`${buildGoogleCloudFilesPath()}?${query.toString()}`, {
        interactive
    });
    const files = listing && Array.isArray(listing.files) ? listing.files : [];
    const record = files.find(item => item && item.id != null && String(item.name || '').trim().toLowerCase() === normalizedTargetName) || files.find(item => item && item.id != null) || null;

    if (record && record.id != null) {
        googleCloudResolvedFileId = String(record.id);
        googleCloudResolvedFileSignature = resolutionSignature;
    }

    return record;
}

async function parseGoogleCloudCanonicalDocumentFromRecord(record) {
    const fileId = record && record.id != null ? String(record.id) : '';
    const fileName = record && typeof record.name === 'string' && record.name.trim()
        ? record.name.trim()
        : getGoogleCloudFileName();

    if (!fileId) {
        return normalizeMissionCanonicalDocument({
            schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
            revision: 0,
            updatedAt: '',
            updatedBy: '',
            missions: [],
            missionSyncMetaById: Object.create(null)
        }, {
            sourceSignature: `googlecloud:${getGoogleCloudFileName()}:unknown`,
            sourceWasEmpty: true
        });
    }

    const payload = await googleCloudRequest(buildGoogleCloudFileDownloadPathById(fileId), {
        interactive: false,
        responseType: 'text'
    });

    if (payload == null) {
        throw new Error(`Google Cloud file content could not be read for file ${fileName || fileId}.`);
    }

    if (typeof payload !== 'string') {
        throw new Error(`Google Cloud file content could not be read for file ${fileName || fileId}.`);
    }

    if (!payload.trim()) {
        return normalizeMissionCanonicalDocument({
            schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
            revision: 0,
            updatedAt: '',
            updatedBy: '',
            missions: [],
            missionSyncMetaById: Object.create(null)
        }, {
            sourceSignature: `googlecloud:${getGoogleCloudFileName()}:${fileId}`,
            sourceWasEmpty: true
        });
    }

    try {
        const parsed = JSON.parse(payload);
        return normalizeMissionCanonicalDocument(parsed, {
            sourceSignature: `googlecloud:${getGoogleCloudFileName()}:${fileId}`,
            sourceLastModified: record && record.modifiedTime ? new Date(record.modifiedTime).getTime() : 0,
            sourceSize: record && Number(record.size) ? Number(record.size) : payload.length,
            sourceWasEmpty: false
        });
    } catch {
        throw new Error(`Google Cloud payload is not valid JSON for file ${fileName || fileId}.`);
    }
}

async function createGoogleCloudCanonicalItemRecord(document) {
    const folderId = getGoogleCloudFolderId();
    const body = {
        name: getGoogleCloudFileName(),
        mimeType: GOOGLE_CLOUD_CANONICAL_FILE_MIME_TYPE
    };

    if (folderId) {
        body.parents = [folderId];
    }

    return googleCloudRequest(buildGoogleCloudFileCreatePath(), {
        method: 'POST',
        interactive: false,
        body: JSON.stringify(body)
    });
}

async function updateGoogleCloudCanonicalItemRecord(fileId, document) {
    const uploadPath = buildGoogleCloudFileUploadPathById(fileId);
    if (!uploadPath) return null;

    return googleCloudRequest(uploadPath, {
        method: 'PATCH',
        baseUrl: GOOGLE_CLOUD_UPLOAD_BASE_URL,
        interactive: false,
        body: serializeMissionCanonicalDocument(document)
    });
}

async function writeGoogleCloudCanonicalDocument(document = null) {
    if (!hasGoogleCloudSyncConfiguration()) return false;
    if (googleCloudWriteInFlight) return false;

    const snapshot = normalizeMissionCanonicalDocument(document || missionCanonicalDocument || buildMissionCanonicalDocument());
    googleCloudWriteInFlight = true;

    try {
        const latestRemoteRecord = await findGoogleCloudCanonicalItemRecord({ interactive: false });
        let docToWrite = snapshot;
        let remoteRecord = latestRemoteRecord || null;

        if (remoteRecord) {
            const remoteDocument = await parseGoogleCloudCanonicalDocumentFromRecord(remoteRecord);
            docToWrite = mergeMissionCanonicalDocuments(remoteDocument, snapshot);
        }

        docToWrite = {
            ...docToWrite,
            revision: Math.max(Number(docToWrite.revision) || 0, Number(snapshot.revision) || 0),
            updatedAt: nowMissionTimestamp(),
            updatedBy: getMissionSyncClientId()
        };

        let updatedRecord = null;
        if (remoteRecord && remoteRecord.id != null) {
            try {
                updatedRecord = await updateGoogleCloudCanonicalItemRecord(remoteRecord.id, docToWrite);
            } catch (error) {
                if (error && error.status === 404) {
                    remoteRecord = null;
                } else {
                    throw error;
                }
            }
        }

        if (!updatedRecord) {
            const createdRecord = await createGoogleCloudCanonicalItemRecord(docToWrite);
            if (createdRecord && createdRecord.id != null) {
                updatedRecord = await updateGoogleCloudCanonicalItemRecord(createdRecord.id, docToWrite);
            } else {
                updatedRecord = createdRecord;
            }
        }

        if (!updatedRecord || updatedRecord.id == null) {
            throw new Error('Google Cloud file could not be created or updated.');
        }

        updateGoogleCloudItemReference(updatedRecord);

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        updateGoogleCloudStatusFromState();
        return true;
    } catch (error) {
        console.warn('Failed to write mission data to Google Cloud.', error);
        setGoogleCloudStatus(graphErrorMessage(error, 'Google Cloud write failed.'), 'error');
        return false;
    } finally {
        googleCloudWriteInFlight = false;
    }
}

function queueGoogleCloudFileWrite(document = null) {
    googleCloudWriteQueue = googleCloudWriteQueue
        .then(() => writeGoogleCloudCanonicalDocument(document))
        .catch(error => {
            console.warn('Google Cloud write queue failed.', error);
        });

    return googleCloudWriteQueue;
}

async function syncMissionCanonicalDocumentFromGoogleCloudFile(options = {}) {
    const { force = false, initializeIfEmpty = false, interactive = false, skipAccessTokenCheck = false } = options;
    if (googleCloudReloadInProgress) return null;
    if (!hasGoogleCloudSyncConfiguration()) {
        updateGoogleCloudStatusFromState();
        return null;
    }

    if (!supportsGoogleCloudSyncTransport()) {
        setGoogleCloudStatus('Google Cloud sync requires a secure browser context.', 'error');
        return null;
    }

    if (!skipAccessTokenCheck) {
        const accessToken = await getGoogleCloudAccessToken({ interactive });
        if (!accessToken) {
            updateGoogleCloudStatusFromState();
            return null;
        }
    }

    if (!force && (googleCloudWriteInFlight || missionCanonicalLocalDirty || editingMissionId != null)) {
        return null;
    }

    googleCloudReloadInProgress = true;
    try {
        const remoteRecord = await findGoogleCloudCanonicalItemRecord({ interactive });
        const remoteWasEmpty = Boolean(remoteRecord && Number(remoteRecord.size) === 0);
        const localDocument = missionCanonicalDocument || loadMissionCanonicalDocumentFromCache() || buildMissionCanonicalDocument();

        if (!remoteRecord) {
            if (initializeIfEmpty) {
                const initialDocument = localDocument || buildMissionCanonicalDocument();
                if (initialDocument) {
                    const initialized = await queueGoogleCloudFileWrite(initialDocument);
                    if (initialized) {
                        applyMissionCanonicalDocumentToRuntime(initialDocument);
                        updateGoogleCloudStatusFromState();
                        return initialDocument;
                    }
                }
            }
            return null;
        }

        updateGoogleCloudItemReference(remoteRecord);
        const remoteDocument = await parseGoogleCloudCanonicalDocumentFromRecord(remoteRecord);
        const mergedDocument = localDocument
            ? mergeMissionCanonicalDocuments(localDocument, remoteDocument)
            : remoteDocument;

        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderTimeline();
        renderMissionCards();
        if (activeTooltipMissionId != null) {
            refreshTooltipForMission(activeTooltipMissionId);
        }

        if (remoteWasEmpty || (localDocument && getComparableMissionCanonicalPayload(mergedDocument) !== getComparableMissionCanonicalPayload(remoteDocument))) {
            void queueGoogleCloudFileWrite(mergedDocument);
        }

        updateGoogleCloudStatusFromState();
        return mergedDocument;
    } catch (error) {
        console.warn('Failed to sync mission data from Google Cloud.', error);
        setGoogleCloudStatus(graphErrorMessage(error, 'Google Cloud sync failed.'), 'error');
        return null;
    } finally {
        googleCloudReloadInProgress = false;
    }
}

async function requestGoogleCloudConnection() {
    if (!isGoogleCloudSyncMode()) {
        setGoogleCloudStatus('Switch to Google Cloud sync first.');
        return null;
    }

    if (!hasGoogleCloudSyncConfiguration()) {
        setGoogleCloudStatus('Google Cloud config is incomplete.', 'error');
        return null;
    }

    try {
        setGoogleCloudStatus('Connecting to Google Cloud...');
        await getGoogleCloudAccessToken({ interactive: true });
        const synced = await syncMissionCanonicalDocumentFromGoogleCloudFile({
            force: true,
            initializeIfEmpty: true,
            interactive: true,
            skipAccessTokenCheck: true
        });
        updateGoogleCloudStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('Google Cloud connection failed.', error);
        setGoogleCloudStatus(graphErrorMessage(error, 'Google Cloud connection failed.'), 'error');
        return null;
    }
}

async function syncGoogleCloudNow() {
    if (!isGoogleCloudSyncMode()) return null;
    if (!hasGoogleCloudSyncConfiguration()) {
        setGoogleCloudStatus('Google Cloud config is incomplete.', 'error');
        return null;
    }

    try {
        const accessToken = await getGoogleCloudAccessToken({ interactive: false });
        if (!accessToken) {
            updateGoogleCloudStatusFromState();
            return null;
        }
        const synced = await syncMissionCanonicalDocumentFromGoogleCloudFile({
            force: true,
            initializeIfEmpty: true,
            skipAccessTokenCheck: true
        });
        updateGoogleCloudStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('Google Cloud sync failed.', error);
        setGoogleCloudStatus(graphErrorMessage(error, 'Google Cloud sync failed.'), 'error');
        return null;
    }
}

function supportsSharePointSyncTransport() {
    return window.isSecureContext && typeof fetch === 'function' && typeof msal !== 'undefined';
}

function getSharePointTenantId() {
    return missionDefaults.sharePointTenantId ? missionDefaults.sharePointTenantId.trim() : '';
}

function getSharePointClientId() {
    return missionDefaults.sharePointClientId ? missionDefaults.sharePointClientId.trim() : '';
}

function getSharePointSiteId() {
    return missionDefaults.sharePointSiteId ? missionDefaults.sharePointSiteId.trim() : '';
}

function getSharePointListIdentifier() {
    return missionDefaults.sharePointListIdentifier ? missionDefaults.sharePointListIdentifier.trim() : '';
}

function getSharePointItemTitle() {
    return missionDefaults.sharePointItemTitle ? missionDefaults.sharePointItemTitle.trim() : '';
}

function getSharePointPayloadField() {
    return missionDefaults.sharePointPayloadField ? missionDefaults.sharePointPayloadField.trim() : '';
}

function getSharePointConfigSignature() {
    return [
        getSharePointTenantId(),
        getSharePointClientId(),
        getSharePointRedirectUri(),
        getSharePointSiteId(),
        getSharePointListIdentifier(),
        getSharePointItemTitle(),
        getSharePointPayloadField()
    ].join('|');
}

function hasSharePointSyncConfiguration() {
    return Boolean(getSharePointClientId() && getSharePointSiteId() && getSharePointListIdentifier());
}

function getSharePointAuthority() {
    const tenantId = getSharePointTenantId();
    return `https://login.microsoftonline.com/${tenantId || 'common'}`;
}

function getSharePointRedirectUri() {
    const configured = missionDefaults.sharePointRedirectUri ? missionDefaults.sharePointRedirectUri.trim() : '';
    if (configured) return configured;
    return window.location.href.split('#')[0];
}

function getSharePointScopes() {
    return ['Sites.ReadWrite.All'];
}

function resetSharePointAuthClient() {
    sharePointAuthClient = null;
    sharePointAuthConfigSignature = '';
    sharePointActiveAccount = null;
}

function getSharePointListResolutionSignature() {
    return [
        getSharePointSiteId(),
        getSharePointListIdentifier()
    ].join('|');
}

function buildSharePointListsPath() {
    const siteId = encodeURIComponent(getSharePointSiteId());
    return `/sites/${siteId}/lists`;
}

function buildSharePointListPath(listId) {
    const siteId = encodeURIComponent(getSharePointSiteId());
    const encodedListId = encodeURIComponent(String(listId));
    return `/sites/${siteId}/lists/${encodedListId}`;
}

function toSharePointGraphRelativePath(nextLink) {
    if (typeof nextLink !== 'string' || !nextLink.trim()) return '';
    return nextLink.startsWith(SHAREPOINT_GRAPH_BASE_URL)
        ? nextLink.slice(SHAREPOINT_GRAPH_BASE_URL.length)
        : nextLink;
}

function escapeSharePointODataString(value) {
    return String(value).replace(/'/g, "''");
}

async function resolveSharePointListId(options = {}) {
    const signature = getSharePointListResolutionSignature();
    if (sharePointResolvedListId && sharePointResolvedListSignature === signature) {
        return sharePointResolvedListId;
    }

    const interactive = Boolean(options.interactive);
    const listIdentifier = getSharePointListIdentifier();
    if (!listIdentifier) {
        throw new Error('SharePoint list identifier is required.');
    }

    try {
        const directRecord = await graphSharePointRequest(`${buildSharePointListPath(listIdentifier)}?$select=id,displayName`, {
            interactive
        });
        if (!directRecord) {
            return null;
        }
        if (directRecord.id != null) {
            sharePointResolvedListId = String(directRecord.id);
            sharePointResolvedListSignature = signature;
            return sharePointResolvedListId;
        }
    } catch (error) {
        if (!error || error.status !== 404) {
            throw error;
        }
    }

    let nextPath = `${buildSharePointListsPath()}?$select=id,displayName&$top=999`;
    const normalizedIdentifier = listIdentifier.trim().toLowerCase();

    while (nextPath) {
        const listing = await graphSharePointRequest(nextPath, { interactive });
        if (!listing) {
            return null;
        }
        const lists = listing && Array.isArray(listing.value) ? listing.value : [];
        const match = lists.find(item => {
            if (!item || item.id == null) return false;
            if (String(item.id) === listIdentifier) return true;
            const displayName = typeof item.displayName === 'string' ? item.displayName.trim().toLowerCase() : '';
            return displayName === normalizedIdentifier;
        }) || null;

        if (match) {
            sharePointResolvedListId = String(match.id);
            sharePointResolvedListSignature = signature;
            return sharePointResolvedListId;
        }

        nextPath = listing && typeof listing['@odata.nextLink'] === 'string'
            ? toSharePointGraphRelativePath(listing['@odata.nextLink'])
            : '';
    }

    throw new Error(`SharePoint list "${listIdentifier}" was not found on site ${getSharePointSiteId()}.`);
}

function updateSharePointItemReference(record) {
    const itemId = record && record.id != null ? String(record.id) : '';
    const etag = record && typeof record['@odata.etag'] === 'string'
        ? record['@odata.etag']
        : (record && typeof record.etag === 'string' ? record.etag : '');

    let changed = false;
    if ((missionDefaults.sharePointItemId || '') !== itemId) {
        missionDefaults.sharePointItemId = itemId;
        changed = true;
    }
    if ((missionDefaults.sharePointItemETag || '') !== etag) {
        missionDefaults.sharePointItemETag = etag;
        changed = true;
    }

    if (changed) {
        syncMissionDefaultsForm();
        persistMissionDefaults();
    }
}

function clearSharePointItemReference() {
    let changed = false;
    if (missionDefaults.sharePointItemId) {
        missionDefaults.sharePointItemId = '';
        changed = true;
    }
    if (missionDefaults.sharePointItemETag) {
        missionDefaults.sharePointItemETag = '';
        changed = true;
    }

    if (changed) {
        syncMissionDefaultsForm();
        persistMissionDefaults();
    }
}

function buildSharePointItemFields(document) {
    const payloadField = getSharePointPayloadField() || SHAREPOINT_DEFAULT_PAYLOAD_FIELD;
    const fields = {
        Title: getSharePointItemTitle() || SHAREPOINT_CANONICAL_ITEM_TITLE
    };
    fields[payloadField] = serializeMissionCanonicalDocument(document);
    return fields;
}

function buildSharePointItemExpandQuery() {
    const payloadField = getSharePointPayloadField() || SHAREPOINT_DEFAULT_PAYLOAD_FIELD;
    return `fields($select=Title,${payloadField})`;
}

async function buildSharePointItemsPath(options = {}) {
    const siteId = encodeURIComponent(getSharePointSiteId());
    const resolvedListId = await resolveSharePointListId(options);
    if (!resolvedListId) return null;
    const listId = encodeURIComponent(resolvedListId);
    return `/sites/${siteId}/lists/${listId}/items`;
}

async function buildSharePointItemPath(itemId, options = {}) {
    const siteId = encodeURIComponent(getSharePointSiteId());
    const resolvedListId = await resolveSharePointListId(options);
    if (!resolvedListId) return null;
    const listId = encodeURIComponent(resolvedListId);
    const encodedItemId = encodeURIComponent(String(itemId));
    return `/sites/${siteId}/lists/${listId}/items/${encodedItemId}`;
}

async function buildSharePointItemFieldsPath(itemId, options = {}) {
    const itemPath = await buildSharePointItemPath(itemId, options);
    return itemPath ? `${itemPath}/fields` : null;
}

function graphErrorMessage(error, fallback = 'SharePoint request failed.') {
    if (!error) return fallback;
    if (error.body && error.body.error && typeof error.body.error.message === 'string') {
        return error.body.error.message;
    }
    if (typeof error.message === 'string' && error.message.trim()) return error.message;
    return fallback;
}

async function getSharePointAuthClient() {
    const signature = getSharePointConfigSignature();
    if (sharePointAuthClient && sharePointAuthConfigSignature === signature) {
        return sharePointAuthClient;
    }

    if (!supportsSharePointSyncTransport()) {
        throw new Error('SharePoint sync requires Microsoft auth support in a secure browser context.');
    }

    const clientId = getSharePointClientId();
    if (!clientId) {
        throw new Error('SharePoint client ID is required.');
    }

    sharePointAuthClient = new msal.PublicClientApplication({
        auth: {
            clientId,
            authority: getSharePointAuthority(),
            redirectUri: getSharePointRedirectUri(),
            navigateToLoginRequestUrl: false
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false
        }
    });
    sharePointAuthConfigSignature = signature;
    sharePointActiveAccount = sharePointAuthClient.getAllAccounts()[0] || null;
    return sharePointAuthClient;
}

async function getSharePointAccessToken(options = {}) {
    const interactive = Boolean(options.interactive);
    const client = await getSharePointAuthClient();
    const scopes = getSharePointScopes();
    const account = sharePointActiveAccount || client.getAllAccounts()[0] || null;

    if (account) {
        try {
            const tokenResult = await client.acquireTokenSilent({
                scopes,
                account
            });
            sharePointActiveAccount = tokenResult.account || account;
            return tokenResult.accessToken;
        } catch (error) {
            if (!interactive) return null;
        }
    }

    if (!interactive) return null;

    const loginResult = await client.loginPopup({
        scopes,
        prompt: 'select_account'
    });
    sharePointActiveAccount = loginResult.account || null;
    if (loginResult.accessToken) return loginResult.accessToken;

    const retryResult = await client.acquireTokenSilent({
        scopes,
        account: sharePointActiveAccount
    });
    return retryResult.accessToken;
}

async function graphSharePointRequest(path, options = {}) {
    const token = await getSharePointAccessToken({ interactive: Boolean(options.interactive) });
    if (!token) return null;

    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    if (options.body != null && !headers.has('Content-Type') && options.method && options.method !== 'GET' && options.method !== 'HEAD') {
        headers.set('Content-Type', 'application/json');
    }

    const response = await fetch(`${SHAREPOINT_GRAPH_BASE_URL}${path}`, {
        method: options.method || 'GET',
        headers,
        body: options.body,
        cache: 'no-store'
    });

    const responseText = await response.text();
    let responseBody = null;
    if (responseText) {
        try {
            responseBody = JSON.parse(responseText);
        } catch {
            responseBody = responseText;
        }
    }

    if (!response.ok) {
        const error = new Error(graphErrorMessage({ body: responseBody, message: response.statusText }, `SharePoint request failed with status ${response.status}.`));
        error.status = response.status;
        error.body = responseBody;
        throw error;
    }

    if (responseBody == null && response.status === 204) {
        return {};
    }

    return responseBody;
}

async function findSharePointCanonicalItemRecord(options = {}) {
    const interactive = Boolean(options.interactive);
    const targetTitle = getSharePointItemTitle() || SHAREPOINT_CANONICAL_ITEM_TITLE;
    const storedItemId = missionDefaults.sharePointItemId ? missionDefaults.sharePointItemId.trim() : '';

    if (!hasSharePointSyncConfiguration()) return null;

    if (storedItemId) {
        try {
            const expandQuery = new URLSearchParams();
            expandQuery.set('$expand', buildSharePointItemExpandQuery());
            const itemPath = await buildSharePointItemPath(storedItemId, { interactive });
            if (!itemPath) return null;
            const record = await graphSharePointRequest(`${itemPath}?${expandQuery.toString()}`, {
                interactive
            });
            if (record) return record;
        } catch (error) {
            if (error && error.status !== 404) {
                throw error;
            }
        }
    }

    const query = new URLSearchParams();
    query.set('$expand', buildSharePointItemExpandQuery());
    query.set('$filter', `fields/Title eq '${escapeSharePointODataString(targetTitle)}'`);
    query.set('$top', '5');

    const itemsPath = await buildSharePointItemsPath({ interactive });
    if (!itemsPath) return null;
    const listing = await graphSharePointRequest(`${itemsPath}?${query.toString()}`, {
        interactive
    });

    const items = listing && Array.isArray(listing.value) ? listing.value : [];
    const byTitle = items.find(item => {
        const title = item && item.fields && typeof item.fields.Title === 'string' ? item.fields.Title.trim() : '';
        return title.toLowerCase() === targetTitle.toLowerCase();
    }) || null;
    return byTitle;
}

function parseSharePointCanonicalDocumentFromRecord(record) {
    const payloadField = getSharePointPayloadField() || SHAREPOINT_DEFAULT_PAYLOAD_FIELD;
    const fields = record && record.fields && typeof record.fields === 'object' ? record.fields : {};
    const payload = fields[payloadField];
    const sourceSignature = `sharepoint:${getSharePointSiteId()}:${sharePointResolvedListId || getSharePointListIdentifier()}:${record && record.id ? String(record.id) : 'unknown'}`;

    if (typeof payload !== 'string' || !payload.trim()) {
        return normalizeMissionCanonicalDocument({
            schemaVersion: MISSION_CANONICAL_SCHEMA_VERSION,
            revision: 0,
            updatedAt: '',
            updatedBy: '',
            missions: [],
            missionSyncMetaById: Object.create(null)
        }, {
            sourceSignature,
            sourceWasEmpty: true
        });
    }

    try {
        const parsed = JSON.parse(payload);
        return normalizeMissionCanonicalDocument(parsed, {
            sourceSignature,
            sourceLastModified: 0,
            sourceSize: payload.length,
            sourceWasEmpty: false
        });
    } catch (error) {
        throw new Error(`SharePoint payload is not valid JSON for item ${record && record.id ? String(record.id) : 'unknown'}.`);
    }
}

async function createSharePointCanonicalItemRecord(document) {
    const itemsPath = await buildSharePointItemsPath();
    if (!itemsPath) return null;
    const created = await graphSharePointRequest(itemsPath, {
        method: 'POST',
        interactive: false,
        body: JSON.stringify({
            fields: buildSharePointItemFields(document)
        })
    });

    return created;
}

async function updateSharePointCanonicalItemRecord(itemId, document, etag = '') {
    const headers = {};
    headers['If-Match'] = etag || '*';

    const fieldsPath = await buildSharePointItemFieldsPath(itemId);
    if (!fieldsPath) return null;
    const updated = await graphSharePointRequest(fieldsPath, {
        method: 'PATCH',
        interactive: false,
        headers,
        body: JSON.stringify(buildSharePointItemFields(document))
    });

    if (updated && Object.keys(updated).length === 0) {
        return {
            id: String(itemId),
            '@odata.etag': etag
        };
    }

    return updated;
}

async function writeSharePointCanonicalDocument(document = null) {
    if (!hasSharePointSyncConfiguration()) return false;
    if (sharePointWriteInFlight) return false;

    const snapshot = normalizeMissionCanonicalDocument(document || missionCanonicalDocument || buildMissionCanonicalDocument());
    sharePointWriteInFlight = true;

    try {
        const latestRemoteRecord = await findSharePointCanonicalItemRecord({ interactive: false });
        let docToWrite = snapshot;
        let remoteRecord = latestRemoteRecord || null;

        if (remoteRecord) {
            const remoteDocument = parseSharePointCanonicalDocumentFromRecord(remoteRecord);
            docToWrite = mergeMissionCanonicalDocuments(remoteDocument, snapshot);
        }

        docToWrite = {
            ...docToWrite,
            revision: Math.max(Number(docToWrite.revision) || 0, Number(snapshot.revision) || 0),
            updatedAt: nowMissionTimestamp(),
            updatedBy: getMissionSyncClientId()
        };

        let updatedRecord = null;
        if (remoteRecord && remoteRecord.id != null) {
            try {
                updatedRecord = await updateSharePointCanonicalItemRecord(remoteRecord.id, docToWrite, remoteRecord['@odata.etag'] || remoteRecord.etag || missionDefaults.sharePointItemETag || '');
            } catch (error) {
                if (error && (error.status === 409 || error.status === 412)) {
                    const retryRecord = await findSharePointCanonicalItemRecord({ interactive: false });
                    if (retryRecord && retryRecord.id != null) {
                        const retryDocument = parseSharePointCanonicalDocumentFromRecord(retryRecord);
                        docToWrite = mergeMissionCanonicalDocuments(retryDocument, snapshot);
                        updatedRecord = await updateSharePointCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord['@odata.etag'] || retryRecord.etag || '');
                    } else {
                        throw error;
                    }
                } else {
                    throw error;
                }
            }
        } else {
            try {
                updatedRecord = await createSharePointCanonicalItemRecord(docToWrite);
            } catch (error) {
                if (error && (error.status === 409 || error.status === 412)) {
                    const retryRecord = await findSharePointCanonicalItemRecord({ interactive: false });
                    if (retryRecord && retryRecord.id != null) {
                        const retryDocument = parseSharePointCanonicalDocumentFromRecord(retryRecord);
                        docToWrite = mergeMissionCanonicalDocuments(retryDocument, snapshot);
                        updatedRecord = await updateSharePointCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord['@odata.etag'] || retryRecord.etag || '');
                    } else {
                        throw error;
                    }
                } else {
                    throw error;
                }
            }
        }

        if (!updatedRecord) {
            const retryRecord = await findSharePointCanonicalItemRecord({ interactive: false });
            if (retryRecord && retryRecord.id != null) {
                updatedRecord = await updateSharePointCanonicalItemRecord(retryRecord.id, docToWrite, retryRecord['@odata.etag'] || retryRecord.etag || '');
            }
        }

        if (!updatedRecord) {
            throw new Error('SharePoint item could not be created or updated.');
        }

        updateSharePointItemReference(updatedRecord);

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        updateSharePointStatusFromState();
        return true;
    } catch (error) {
        console.warn('Failed to write mission data to SharePoint.', error);
        setSharePointStatus(graphErrorMessage(error, 'SharePoint write failed.'), 'error');
        return false;
    } finally {
        sharePointWriteInFlight = false;
    }
}

function queueSharePointListWrite(document = null) {
    sharePointWriteQueue = sharePointWriteQueue
        .then(() => writeSharePointCanonicalDocument(document))
        .catch(error => {
            console.warn('SharePoint write queue failed.', error);
        });

    return sharePointWriteQueue;
}

async function syncMissionCanonicalDocumentFromSharePointList(options = {}) {
    const { force = false, initializeIfEmpty = false, interactive = false, skipAccessTokenCheck = false } = options;
    if (sharePointReloadInProgress) return null;
    if (!hasSharePointSyncConfiguration()) {
        updateSharePointStatusFromState();
        return null;
    }

    if (!skipAccessTokenCheck) {
        const accessToken = await getSharePointAccessToken({ interactive });
        if (!accessToken) {
            updateSharePointStatusFromState();
            return null;
        }
    }

    if (!force && (sharePointWriteInFlight || missionCanonicalLocalDirty || editingMissionId != null)) {
        return null;
    }

    sharePointReloadInProgress = true;
    try {
        const remoteRecord = await findSharePointCanonicalItemRecord({ interactive });
        const remoteWasEmpty = Boolean(remoteRecord && remoteRecord.fields && !(remoteRecord.fields[getSharePointPayloadField() || SHAREPOINT_DEFAULT_PAYLOAD_FIELD]));
        const localDocument = missionCanonicalDocument || loadMissionCanonicalDocumentFromCache() || buildMissionCanonicalDocument();

        if (!remoteRecord) {
            if (initializeIfEmpty) {
                const initialDocument = localDocument || buildMissionCanonicalDocument();
                if (initialDocument) {
                    const initialized = await queueSharePointListWrite(initialDocument);
                    if (initialized) {
                        applyMissionCanonicalDocumentToRuntime(initialDocument);
                        updateSharePointStatusFromState();
                        return initialDocument;
                    }
                }
            }
            return null;
        }

        updateSharePointItemReference(remoteRecord);
        const remoteDocument = parseSharePointCanonicalDocumentFromRecord(remoteRecord);
        const mergedDocument = localDocument
            ? mergeMissionCanonicalDocuments(localDocument, remoteDocument)
            : remoteDocument;

        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderTimeline();
        renderMissionCards();
        if (activeTooltipMissionId != null) {
            refreshTooltipForMission(activeTooltipMissionId);
        }

        if (remoteWasEmpty || (localDocument && getComparableMissionCanonicalPayload(mergedDocument) !== getComparableMissionCanonicalPayload(remoteDocument))) {
            void queueSharePointListWrite(mergedDocument);
        }

        updateSharePointStatusFromState();
        return mergedDocument;
    } catch (error) {
        console.warn('Failed to sync mission data from SharePoint.', error);
        setSharePointStatus(graphErrorMessage(error, 'SharePoint sync failed.'), 'error');
        return null;
    } finally {
        sharePointReloadInProgress = false;
    }
}

async function requestSharePointConnection() {
    if (!isSharePointSyncMode()) {
        setSharePointStatus('Switch to SharePoint sync first.');
        return null;
    }

    if (!hasSharePointSyncConfiguration()) {
        setSharePointStatus('SharePoint config is incomplete.', 'error');
        return null;
    }

    try {
        setSharePointStatus('Connecting to SharePoint...');
        await getSharePointAccessToken({ interactive: true });
        const synced = await syncMissionCanonicalDocumentFromSharePointList({
            force: true,
            initializeIfEmpty: true,
            interactive: true,
            skipAccessTokenCheck: true
        });
        updateSharePointStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('SharePoint connection failed.', error);
        setSharePointStatus(graphErrorMessage(error, 'SharePoint connection failed.'), 'error');
        return null;
    }
}

async function syncSharePointNow() {
    if (!isSharePointSyncMode()) return null;
    if (!hasSharePointSyncConfiguration()) {
        setSharePointStatus('SharePoint config is incomplete.', 'error');
        return null;
    }

    try {
        const accessToken = await getSharePointAccessToken({ interactive: false });
        if (!accessToken) {
            updateSharePointStatusFromState();
            return null;
        }
        const synced = await syncMissionCanonicalDocumentFromSharePointList({
            force: true,
            initializeIfEmpty: true,
            skipAccessTokenCheck: true
        });
        updateSharePointStatusFromState();
        startMissionCanonicalSyncLoop();
        return synced;
    } catch (error) {
        console.warn('SharePoint sync failed.', error);
        setSharePointStatus(graphErrorMessage(error, 'SharePoint sync failed.'), 'error');
        return null;
    }
}

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

function readMissionDataFromForm() {
    const legNodes = document.querySelectorAll('.leg-container');
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
            liftCustomer: document.getElementById('liftCustomer').value,
            liftPax: document.getElementById('liftPax').value,
            liftCargo: document.getElementById('liftCargo').value,
            liftHazmat: document.getElementById('liftHazmat').value,
            legs
        }
    };
}

async function buildMissionDataFromForm() {
    await loadAirportData().catch(() => {});

    const legNodes = document.querySelectorAll('.leg-container');
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
    tooltip.style.display = 'none';
    clearHoverRouteMap();
    touchMoved = false;
    touchActiveCount = e.touches.length;

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
    if (e.touches.length === 1) {
        // Dropped from 2 fingers to 1: restart pan tracking
        touchMoved = false;
        touchPanStartX = e.touches[0].clientX;
        touchPanStartViewStart = viewStart;
        touchPanStartViewEnd = viewEnd;
    }
}, { passive: true });

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

        const clampedLeft = Math.max(0, leftPct);
        const clampedRight = Math.min(100, leftPct + widthPct);
        bar.style.left = `${clampedLeft}%`;
        bar.style.width = `${Math.max(0, clampedRight - clampedLeft)}%`;
        
        // Assign top position based on matching tail row
        const tailStr = mission.tailNum ? mission.tailNum.toUpperCase() : 'TBD';
        const tailIndex = uniqueTails.indexOf(tailStr);
        bar.style.top = `${rowOffset + (tailIndex * rowHeight) + 8}px`; 
        
        const firstIcao = escapeHtml(mission.legs[0].takeoffIcao);
        const lastIcao = escapeHtml(mission.legs[mission.legs.length - 1].landIcao);
        bar.innerHTML = `<span>${escapeHtml(mission.missionNum)} (${firstIcao}&rarr;${lastIcao})</span>`;

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
        bar.addEventListener('touchend', (e) => {
            if (touchMoved) return;
            if (touchTapMissionId != null && String(touchTapMissionId) === String(mission.id)) return;
            const touch = e.changedTouches[0];
            if (!touch) return;
            showMissionTooltip(mission, { pageX: touch.pageX, pageY: touch.pageY });
        });
        missionsContainer.appendChild(bar);
    });
}

function renderMissionCards() {
    const list = document.getElementById('mission-list');
    clearMissionCardFocusHighlight();
    list.innerHTML = '';
    const sorted = [...missions].sort((a, b) => getMissionTimes(b).start - getMissionTimes(a).start);

    sorted.forEach(mission => {
        const card = document.createElement('div');
        card.className = 'mission-card';
        card.id = `card-${mission.id}`;

        const hasAnyCrew = !!(mission.pilot || mission.copilot || mission.crewChief || mission.loadmaster);
        const isComplete = !!(mission.missionNum && mission.tailNum && mission.legs.length > 0 && 
                              mission.pilot && mission.copilot && mission.crewChief && mission.loadmaster);

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
                    <p><strong>Customer:</strong> ${mission.liftCustomer ? escapeHtml(mission.liftCustomer) : '0'}</p>
                    <p><strong>Pax:</strong> ${mission.liftPax ? escapeHtml(String(mission.liftPax)) : '0'}</p>
                    <p><strong>Cargo:</strong> ${mission.liftCargo ? escapeHtml(mission.liftCargo) : '<em>None</em>'}</p>
                    <p><strong>Hazmat:</strong> ${mission.liftHazmat ? escapeHtml(mission.liftHazmat) : '<em>None</em>'}</p>
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

        for(let i=1; i<lines.length; i++) {
            const cols = parseCsvLine(lines[i]).map(c => c.trim());
            let row = {};
            headers.forEach((h, idx) => row[h] = cols[idx]);
            
            const mNum = row.mission;
            if(!mNum) continue;

            if(!missionMap.has(mNum)) {
                missionMap.set(mNum, {
                    id: createMissionId('csv'),
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
                if (normalizedMission) {
                    missions.push(normalizedMission);
                    markMissionUpdated(normalizedMission.id);
                }
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
        }
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

init();
