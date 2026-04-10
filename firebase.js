// =========================================================================
// CONEXÃO COM O BANCO DE DADOS (FIREBASE CLOUD) — v164.1
// Fix: throttle de escrita para evitar "Write stream exhausted"
// =========================================================================
const firebaseConfig = {
    apiKey:            "AIzaSyDMs67hSbMC3zMJt7qYct1zuvQhu7Rk_t4",
    authDomain:        "sistema-logistica-f03ef.firebaseapp.com",
    projectId:         "sistema-logistica-f03ef",
    storageBucket:     "sistema-logistica-f03ef.firebasestorage.app",
    messagingSenderId: "874583974046",
    appId:             "1:874583974046:web:c12cb71a831ea649a89b8f"
};

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}
const dbCloud = firebase.firestore();

// ── Throttle: evita escrever a mesma chave mais de 1x por minuto ──────────
const _saveThrottle = {};
const _saveQueue    = {};
let   _saveTimer    = null;

function _flushSaveQueue() {
    const keys = Object.keys(_saveQueue);
    if (!keys.length) return;
    keys.forEach(async k => {
        const v = _saveQueue[k];
        delete _saveQueue[k];
        await _doSaveToDB(k, v);
    });
}

async function _doSaveToDB(k, v) {
    try {
        const strData  = JSON.stringify(v);
        const chunkSize = 800000;
        if (strData.length > chunkSize) {
            const totalChunks = Math.ceil(strData.length / chunkSize);
            await dbCloud.collection('logistica_dados').doc(k).set({ isChunked: true, chunks: totalChunks });
            for (let i = 0; i < totalChunks; i++) {
                await dbCloud.collection('logistica_dados').doc(`${k}_chunk_${i}`)
                    .set({ payload: strData.substring(i * chunkSize, (i + 1) * chunkSize) });
            }
        } else {
            await dbCloud.collection('logistica_dados').doc(k).set({ isChunked: false, payload: strData });
        }
        _saveThrottle[k] = Date.now();
    } catch (e) {
        // Se resource-exhausted, agenda retry em 2 minutos
        if (e && e.code === 'resource-exhausted') {
            console.warn('[Firebase] Cota atingida para "' + k + '". Retry em 2 min.');
            setTimeout(() => _doSaveToDB(k, v), 120000);
        } else {
            console.error('[Firebase] Erro ao salvar "' + k + '":', e);
            throw e;
        }
    }
}

window.openDB = async function () { return true; };

window.saveToDB = async function (k, v) {
    const now      = Date.now();
    const lastSave = _saveThrottle[k] || 0;
    const MIN_INTERVAL = 60000; // mínimo 60s entre escritas da mesma chave

    if (now - lastSave < MIN_INTERVAL) {
        // Enfileira para escrita futura (debounce de 5s)
        _saveQueue[k] = v;
        clearTimeout(_saveTimer);
        _saveTimer = setTimeout(_flushSaveQueue, 5000);
        return;
    }
    await _doSaveToDB(k, v);
};

window.getFromDB = async function (k) {
    try {
        const doc = await dbCloud.collection('logistica_dados').doc(k).get();
        if (doc.exists) {
            const data = doc.data();
            if (data.isChunked) {
                let fullStr = '';
                for (let i = 0; i < data.chunks; i++) {
                    const chunkDoc = await dbCloud.collection('logistica_dados').doc(`${k}_chunk_${i}`).get();
                    if (chunkDoc.exists) fullStr += chunkDoc.data().payload;
                }
                return JSON.parse(fullStr);
            } else {
                return JSON.parse(data.payload || '{}');
            }
        }
        return null;
    } catch (e) {
        console.warn('[Firebase] Erro ao ler "' + k + '":', e);
        return null;
    }
};
