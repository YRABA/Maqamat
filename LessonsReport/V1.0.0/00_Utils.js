/** Utils: המרות/עזר */
const Utils = (() => {

  function str(v) {
    return (v == null) ? '' : String(v).trim();
  }

  function toDateKey(d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    const y = d.getFullYear();
    const m = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    return `${y}-${m}-${day}`;
  }

  function coerceToDate(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate(), 12, 0, 0);

    let d = null;
    if (typeof v === 'number') {
      // Excel serial number conversion
      d = new Date(Math.round((v - 25569) * 86400 * 1000));
    } else if (typeof v === 'string') {
      const s = v.trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        const [Y, M, D] = s.split('-').map(Number);
        d = new Date(Y, M - 1, D);
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
        const [D, M, Y] = s.split('/').map(Number);
        d = new Date(Y, M - 1, D);
      } else {
        d = new Date(s);
      }
    }

    if (!d || isNaN(d)) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
  }

  function indexMap(headers) {
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);
    return idx;
  }

  return {
    str,
    toDateKey,
    coerceToDate,
    indexMap
  };
})();
