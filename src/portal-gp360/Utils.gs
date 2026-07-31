// =============================================================================
// HELPERS E MÉTODOS DE EXTRAÇÃO SEGURA
// =============================================================================

function extrairIdDrive(linkOuId) {
  if (!linkOuId) return "";
  const match = linkOuId.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : linkOuId;
}

function formatarDataSegura(dataValor) {
  if (!dataValor) return "";
  try { 
    if (dataValor instanceof Date) {
        return Utilities.formatDate(dataValor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) return ds;

    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); 
    }
    return ds; 
  } catch(e) { return String(dataValor).split(' ')[0]; }
}

function obterDataRawSegura(dataValor) {
  if (!dataValor) return 0;
  try { 
    if (dataValor instanceof Date) return dataValor.getTime();

    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) {
        let p = ds.split('/');
        if(p.length === 3) return new Date(p[2], p[1]-1, p[0]).getTime();
    } else if (ds.includes('-')) {
        let p = ds.split('-');
        if(p.length === 3) return new Date(p[0], p[1]-1, p[2]).getTime();
    }
    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) return dt.getTime(); 
    return 0; 
  } catch(e) { return 0; }
}
