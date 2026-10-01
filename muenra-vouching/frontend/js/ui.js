// Utilidades de interfaz: construcción segura del DOM (sin innerHTML con datos), formato y avisos.

export function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    }
  }
  append(el, children);
  return el;
}

function append(el, children) {
  for (const c of children.flat(Infinity)) {
    if (c === null || c === undefined || c === false) continue;
    el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  }
}

export function clear(el, ...children) { el.replaceChildren(); append(el, children); return el; }

const fmtMoney = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
const fmtPct = new Intl.NumberFormat('es-CO', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 2 });
const fmtInt = new Intl.NumberFormat('es-CO');

export const fmt = {
  money: (v) => (v === null || v === undefined || v === '' ? '—' : fmtMoney.format(Number(v))),
  pct: (v) => (v === null || v === undefined ? '—' : fmtPct.format(Number(v))),
  int: (v) => (v === null || v === undefined ? '—' : fmtInt.format(Number(v))),
  date: (v) => (v ? String(v).slice(0, 10) : '—'),
  datetime: (v) => (v ? new Date(v.endsWith && !/Z|[+-]\d\d:\d\d$/.test(v) ? `${v}Z` : v).toLocaleString('es-CO') : '—'),
  text: (v) => (v === null || v === undefined || v === '' ? '—' : String(v)),
};

export const STATUSES = ['COINCIDE', 'COINCIDE CON TOLERANCIA', 'EXCEPCIÓN', 'REVISIÓN MANUAL', 'SIN SOPORTE',
  'SOPORTE NO REFERENCIADO', 'POSIBLE DUPLICADO', 'DOCUMENTO ILEGIBLE', 'ERROR DE LECTURA', 'PENDIENTE'];
export const STATUS_COLORS = {
  'COINCIDE': '#1e7b4b', 'COINCIDE CON TOLERANCIA': '#7cb342', 'EXCEPCIÓN': '#c62f3b', 'REVISIÓN MANUAL': '#e2a400',
  'SIN SOPORTE': '#e07b1a', 'SOPORTE NO REFERENCIADO': '#2a5ea8', 'POSIBLE DUPLICADO': '#b03a78',
  'DOCUMENTO ILEGIBLE': '#7d8693', 'ERROR DE LECTURA': '#4b5361', 'PENDIENTE': '#c3cad5',
};

export function statusClass(s) {
  return `st-${String(s || 'PENDIENTE').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '-')}`;
}
export function badge(s) { return h('span', { class: `badge ${statusClass(s)}` }, s || 'PENDIENTE'); }

export function scoreTag(v) {
  if (v === null || v === undefined) return h('span', { class: 'score na' }, 'N/E');
  const cls = v >= 0.85 ? 'hi' : v >= 0.6 ? 'mid' : 'lo';
  return h('span', { class: `score ${cls}` }, `${Math.round(v * 100)}`);
}

export function toast(msg, kind = '') {
  const el = h('div', { class: `toast ${kind}`, role: 'status' }, msg);
  document.getElementById('toasts').appendChild(el);
  setTimeout(() => el.remove(), kind === 'bad' ? 7000 : 4000);
}

export function modal(title, body, actions = []) {
  return new Promise((resolve) => {
    const close = (v) => { back.remove(); resolve(v); };
    const back = h('div', { class: 'modal-back', onclick: (e) => { if (e.target === back) close(null); } },
      h('div', { class: 'modal', role: 'dialog', 'aria-modal': 'true' },
        h('h2', null, title), body,
        h('div', { class: 'btn-row', style: { marginTop: '16px', justifyContent: 'flex-end' } },
          h('button', { class: 'btn ghost', onclick: () => close(null) }, 'Cancelar'),
          actions.map((a) => h('button', { class: `btn ${a.class || 'primary'}`, onclick: async () => {
            const v = a.value ? await a.value() : true;
            if (v !== false) close(v);
          } }, a.label)))));
    document.body.appendChild(back);
    const first = back.querySelector('input, textarea, select');
    if (first) first.focus();
  });
}

export function loading(text = 'Cargando…') { return h('div', { class: 'empty' }, h('span', { class: 'spinner' }), ' ', text); }

export function errorBox(err) { return h('div', { class: 'alert bad' }, err.message || String(err)); }

export function kpi(label, value, sub, accent) {
  return h('div', { class: 'kpi', style: accent ? { '--accent': accent } : null },
    h('div', { class: 'label' }, label), h('div', { class: 'value' }, value), sub ? h('div', { class: 'sub' }, sub) : null);
}

export function field(label, input, hint) {
  return h('div', { class: 'field' }, h('label', null, label), input, hint ? h('div', { class: 'small muted' }, hint) : null);
}

export function sortableTable(columns, rows, { onRowClick, pageSize = 100, empty = 'Sin registros' } = {}) {
  // columns: [{key, label, render?, num?, sort?}]
  let sortKey = null; let asc = true; let page = 0;
  const wrap = h('div');
  const draw = () => {
    let data = rows.slice();
    if (sortKey) {
      const col = columns.find((c) => c.key === sortKey);
      const get = col.sort || ((r) => r[sortKey]);
      data.sort((a, b) => {
        const x = get(a); const y = get(b);
        if (x === y) return 0;
        if (x === null || x === undefined) return 1;
        if (y === null || y === undefined) return -1;
        return (x > y ? 1 : -1) * (asc ? 1 : -1);
      });
    }
    const pages = Math.max(1, Math.ceil(data.length / pageSize));
    page = Math.min(page, pages - 1);
    const slice = data.slice(page * pageSize, (page + 1) * pageSize);
    const table = h('table', { class: 'table' },
      h('thead', null, h('tr', null, columns.map((c) => h('th', {
        class: `sortable ${c.num ? 'num' : ''}`,
        onclick: () => { if (sortKey === c.key) asc = !asc; else { sortKey = c.key; asc = true; } draw(); },
      }, c.label, sortKey === c.key ? (asc ? ' ▲' : ' ▼') : '')))),
      h('tbody', null, slice.length ? slice.map((r) => h('tr', { class: onRowClick ? 'clickable' : '', onclick: onRowClick ? () => onRowClick(r) : null },
        columns.map((c) => h('td', { class: c.num ? 'num' : '' }, c.render ? c.render(r) : fmt.text(r[c.key])))))
        : h('tr', null, h('td', { colspan: columns.length }, h('div', { class: 'empty' }, empty)))));
    clear(wrap, h('div', { class: 'table-wrap' }, table),
      pages > 1 ? h('div', { class: 'btn-row', style: { marginTop: '8px' } },
        h('button', { class: 'btn sm', disabled: page === 0, onclick: () => { page -= 1; draw(); } }, '‹ Anterior'),
        h('span', { class: 'small muted' }, `Página ${page + 1} de ${pages} · ${data.length} registros`),
        h('button', { class: 'btn sm', disabled: page >= pages - 1, onclick: () => { page += 1; draw(); } }, 'Siguiente ›'))
        : h('div', { class: 'small muted', style: { marginTop: '6px' } }, `${data.length} registros`));
  };
  draw();
  return wrap;
}
