// Cliente HTTP de la API con token Bearer (guardado en sessionStorage: se borra al cerrar el navegador).
const BASE = '/api';
let onUnauthorized = () => {};

export const session = {
  get token() { try { return sessionStorage.getItem('mv_token'); } catch { return null; } },
  get user() { try { return JSON.parse(sessionStorage.getItem('mv_user') || 'null'); } catch { return null; } },
  save(token, user) { sessionStorage.setItem('mv_token', token); sessionStorage.setItem('mv_user', JSON.stringify(user)); },
  clear() { try { sessionStorage.removeItem('mv_token'); sessionStorage.removeItem('mv_user'); } catch { /* */ } },
  can(perm) { const u = this.user; return !!(u && u.permisos && u.permisos.includes(perm)); },
};

export function setUnauthorizedHandler(fn) { onUnauthorized = fn; }

export class ApiError extends Error {
  constructor(status, message, data) { super(message); this.status = status; this.data = data; }
}

export async function api(path, { method = 'GET', body, form, blob = false, query } = {}) {
  const headers = {};
  if (session.token) headers.Authorization = `Bearer ${session.token}`;
  let payload;
  if (form) payload = form;
  else if (body !== undefined) { headers['Content-Type'] = 'application/json'; payload = JSON.stringify(body); }
  let url = BASE + path;
  if (query) {
    const qs = new URLSearchParams();
    Object.entries(query).forEach(([k, v]) => {
      if (v === undefined || v === null || v === '') return;
      (Array.isArray(v) ? v : [v]).forEach((x) => qs.append(k, x));
    });
    if ([...qs].length) url += `?${qs}`;
  }
  const res = await fetch(url, { method, headers, body: payload });
  if (res.status === 401 && !path.startsWith('/auth/login')) { session.clear(); onUnauthorized(); throw new ApiError(401, 'Sesión expirada'); }
  if (blob && res.ok) return { blob: await res.blob(), filename: filenameFrom(res) };
  const ct = res.headers.get('content-type') || '';
  const data = ct.includes('json') ? await res.json() : await res.text();
  if (!res.ok) {
    let msg = (data && data.detail) || res.statusText;
    if (Array.isArray(msg)) msg = msg.map((d) => `${(d.loc || []).slice(-1)[0]}: ${d.msg}`).join('; ');
    throw new ApiError(res.status, typeof msg === 'string' ? msg : 'Error', data);
  }
  return data;
}

function filenameFrom(res) {
  const cd = res.headers.get('content-disposition') || '';
  const m = cd.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i);
  return m ? decodeURIComponent(m[1] || m[2]) : 'descarga';
}

export async function download(path, query) {
  const { blob, filename } = await api(path, { blob: true, query });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

const imgCache = new Map();
export async function imageUrl(path) {
  if (imgCache.has(path)) return imgCache.get(path);
  const { blob } = await api(path, { blob: true });
  const url = URL.createObjectURL(blob);
  imgCache.set(path, url);
  return url;
}
