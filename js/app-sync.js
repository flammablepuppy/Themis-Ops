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
        const syncAnimationState = remoteRecord ? getMissionSyncAnimationState(snapshot, docToWrite) : null;

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        if (syncAnimationState && syncAnimationState.shouldAnimateRefresh) {
            renderMissionViewsAfterSync(syncAnimationState);
        }
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

        const syncAnimationState = getMissionSyncAnimationState(localDocument, mergedDocument);
        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderMissionViewsAfterSync(syncAnimationState);

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
    return Boolean(getGoogleCloudClientId() && getGoogleCloudFileName() && getGoogleCloudFolderId());
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
    const folderId = getGoogleCloudFolderId();
    const storedFileId = missionDefaults.googleCloudFileId ? missionDefaults.googleCloudFileId.trim() : '';
    const resolutionSignature = getGoogleCloudResolutionSignature();

    if (!hasGoogleCloudSyncConfiguration()) return null;
    if (!folderId) return null;

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
    if (!folderId) {
        throw new Error('Google Cloud shared folder ID is required.');
    }
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
        const syncAnimationState = remoteRecord ? getMissionSyncAnimationState(snapshot, docToWrite) : null;

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        if (syncAnimationState && syncAnimationState.shouldAnimateRefresh) {
            renderMissionViewsAfterSync(syncAnimationState);
        }
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

        const syncAnimationState = getMissionSyncAnimationState(localDocument, mergedDocument);
        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderMissionViewsAfterSync(syncAnimationState);

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
        setGoogleCloudStatus('Google Cloud shared folder ID required.', 'error');
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
        setGoogleCloudStatus('Google Cloud shared folder ID required.', 'error');
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
        const syncAnimationState = remoteRecord ? getMissionSyncAnimationState(snapshot, docToWrite) : null;

        missionCanonicalDocument = {
            ...docToWrite,
            missions: docToWrite.missions.map(cloneMissionRecord),
            missionSyncMetaById: cloneMissionSyncMetaMap(docToWrite.missionSyncMetaById)
        };
        missions = missionCanonicalDocument.missions.map(cloneMissionRecord);
        missionSyncMetaById = cloneMissionSyncMetaMap(missionCanonicalDocument.missionSyncMetaById);
        storeMissionCanonicalDocumentLocally(missionCanonicalDocument);
        missionCanonicalLocalDirty = false;
        if (syncAnimationState && syncAnimationState.shouldAnimateRefresh) {
            renderMissionViewsAfterSync(syncAnimationState);
        }
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

        const syncAnimationState = getMissionSyncAnimationState(localDocument, mergedDocument);
        applyMissionCanonicalDocumentToRuntime(mergedDocument);
        renderMissionViewsAfterSync(syncAnimationState);

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

