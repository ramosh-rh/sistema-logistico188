// =========================================================================
// CONEXÃO COM O BANCO DE DADOS (FIREBASE CLOUD)
// =========================================================================
const firebaseConfig = { 
    apiKey: "AIzaSyDMs67hSbMC3zMJt7qYct1zuvQhu7Rk_t4", 
    authDomain: "sistema-logistica-f03ef.firebaseapp.com", 
    projectId: "sistema-logistica-f03ef", 
    storageBucket: "sistema-logistica-f03ef.firebasestorage.app", 
    messagingSenderId: "874583974046", 
    appId: "1:874583974046:web:c12cb71a831ea649a89b8f" 
};

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}
const dbCloud = firebase.firestore();

window.openDB = async function () { return true; };

window.saveToDB = async function (k, v) {
    try {
        const strData = JSON.stringify(v); const chunkSize = 800000;
        if (strData.length > chunkSize) {
            const totalChunks = Math.ceil(strData.length / chunkSize);
            await dbCloud.collection("logistica_dados").doc(k).set({ isChunked: true, chunks: totalChunks });
            for (let i = 0; i < totalChunks; i++) {
                const chunkStr = strData.substring(i * chunkSize, (i + 1) * chunkSize);
                await dbCloud.collection("logistica_dados").doc(`${k}_chunk_${i}`).set({ payload: chunkStr });
            }
        } else { await dbCloud.collection("logistica_dados").doc(k).set({ isChunked: false, payload: strData }); }
    } catch (e) { console.error(e); throw e; }
};

window.getFromDB = async function (k) {
    try {
        const doc = await dbCloud.collection("logistica_dados").doc(k).get();
        if (doc.exists) {
            const data = doc.data();
            if (data.isChunked) {
                let fullStr = "";
                for (let i = 0; i < data.chunks; i++) {
                    const chunkDoc = await dbCloud.collection("logistica_dados").doc(`${k}_chunk_${i}`).get();
                    if (chunkDoc.exists) fullStr += chunkDoc.data().payload;
                }
                return JSON.parse(fullStr);
            } else { return JSON.parse(data.payload || "{}"); }
        }
        return null;
    } catch (e) { return null; }
};