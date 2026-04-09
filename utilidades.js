// =========================================================================
// FERRAMENTAS E UTILIDADES (FORMATADORES)
// =========================================================================
function parseNum(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return val;
    let str = String(val).trim();
    if (str.includes(',') && str.includes('.')) { str = str.replace(/\./g, '').replace(',', '.'); }
    else if (str.includes(',')) { str = str.replace(',', '.'); }
    str = str.replace(/[^\d\.-]/g, '');
    let n = parseFloat(str);
    return isNaN(n) ? 0 : n;
}

function fmtInt(n) { return Math.round(n || 0).toString(); }
function fmtBig(n) { return (n || 0).toLocaleString('pt-BR'); }
function safeUpdate(id, val) { try { const el = document.getElementById(id); if (el) { if (el.tagName === 'INPUT') el.value = val; else el.innerText = val; } } catch (e) {} }
function getVal(id) { try { const el = document.getElementById(id); if (!el) return 0; return parseInt(el.innerText) || 0; } catch (e) { return 0; } }
function sanitize(str) { if (!str) return ""; return String(str).replace(/[\.\-\/\s]/g, "").toUpperCase(); }
function formatDateBR(isoDateStr) { if(!isoDateStr) return '-'; const parts = String(isoDateStr).split('-'); if(parts.length===3) return `${parts[2]}/${parts[1]}/${parts[0]}`; return isoDateStr; }
function formatDecimalToTime(decimalHours) {
    if (!decimalHours || decimalHours === 0) return "00:00";
    let sign = decimalHours < 0 ? '-' : '+';
    let absH = Math.abs(decimalHours);
    let h = Math.floor(absH);
    let m = Math.round((absH - h) * 60);
    if (m === 60) { h += 1; m = 0; }
    return `${sign}${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}