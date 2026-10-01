// Pantalla 7: tabla de resultados con filtros por estado y búsqueda.
import { api } from '../api.js';
import { badge, clear, fmt, h, sortableTable, STATUSES } from '../ui.js';

export function columns(pid, ctx) {
  return [
    { key: 'id_muestra', label: 'ID muestra' },
    { key: 'estado', label: 'Estado', render: (r) => h('div', null, badge(r.estado), r.decision_revisor ? h('div', { class: 'small muted' }, `✔ ${r.decision_revisor}`) : null) },
    { key: 'tercero_esperado', label: 'Tercero', render: (r) => h('div', null, r.tercero_esperado, h('div', { class: 'small muted' }, r.nit_esperado)) },
    { key: 'numero_esperado', label: 'Número esp. / ext.', render: (r) => h('div', null, fmt.text(r.numero_esperado), h('div', { class: 'small muted' }, fmt.text(r.numero_extraido))) },
    { key: 'fecha_esperada', label: 'Fecha esp. / ext.', render: (r) => h('div', null, fmt.date(r.fecha_esperada), h('div', { class: 'small muted' }, fmt.date(r.fecha_extraida))) },
    { key: 'valor_esperado', label: 'Valor esperado', num: true, render: (r) => fmt.money(r.valor_esperado) },
    { key: 'valor_extraido', label: 'Valor extraído', num: true, render: (r) => fmt.money(r.valor_extraido) },
    { key: 'diferencia_porcentual', label: 'Dif. %', num: true, render: (r) => fmt.pct(r.diferencia_porcentual) },
    { key: 'coincidencia_nit', label: 'NIT', render: (r) => fmt.text(r.coincidencia_nit) },
    { key: 'similitud_nombre', label: 'Sim. nombre', num: true, render: (r) => (r.similitud_nombre === null ? '—' : `${fmt.int(r.similitud_nombre)} %`) },
    { key: 'puntaje', label: 'Puntaje', num: true, render: (r) => (r.puntaje === null ? '—' : fmt.int(r.puntaje * 100)) },
    { key: 'archivos', label: 'Soportes', render: (r) => h('div', { class: 'small' }, r.documentos.map((d) => h('div', null, d.rol === 'COMPLEMENTARIO' ? '＋ ' : '', d.archivo))) },
    { key: 'motivos', label: 'Motivo', sort: (r) => (r.motivos || []).join(' '), render: (r) => h('div', { class: 'small', style: { maxWidth: '320px' } }, (r.motivos || []).slice(0, 2).join(' · ')) },
  ];
}

export async function render(ctx) {
  const pid = ctx.project.id;
  const all = await api(`/projects/${pid}/results`);
  const selected = new Set(ctx.query.getAll('estado'));
  let q = ctx.query.get('q') || '';
  const chips = h('div', { class: 'chips' });
  const table = h('div');
  const counts = {};
  all.forEach((r) => { counts[r.estado] = (counts[r.estado] || 0) + 1; });
  const draw = () => {
    clear(chips, h('button', { class: `chip ${selected.size ? '' : 'active'}`, onclick: () => { selected.clear(); draw(); } }, 'Todos', h('span', { class: 'count' }, all.length)),
      STATUSES.filter((s) => counts[s]).map((s) => h('button', { class: `chip ${selected.has(s) ? 'active' : ''}`,
        onclick: () => { if (selected.has(s)) selected.delete(s); else selected.add(s); draw(); } }, s, h('span', { class: 'count' }, counts[s]))));
    const ql = q.toLowerCase();
    const rows = all.filter((r) => (!selected.size || selected.has(r.estado)) && (!ql || [r.id_muestra, r.tercero_esperado, r.nit_esperado, r.numero_esperado, r.numero_extraido, r.archivos, r.concepto]
      .some((v) => v && String(v).toLowerCase().includes(ql))));
    clear(table, sortableTable(columns(pid, ctx), rows, { onRowClick: (r) => ctx.navigate(`/p/${pid}/visor/${r.id}`), empty: 'Sin resultados para los filtros aplicados' }));
  };
  draw();
  const search = h('input', { class: 'input', placeholder: 'Buscar por ID, tercero, NIT, número o archivo…', value: q, oninput: (e) => { q = e.target.value; draw(); } });
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Tabla de resultados'),
      h('div', { class: 'muted' }, 'Haga clic en una partida para abrir el visor de evidencia y la explicación por criterio.')),
    h('a', { class: 'btn', href: `#/p/${pid}/exportaciones` }, 'Exportar a Excel')),
    h('div', { class: 'card' }, h('div', { class: 'toolbar' }, search), chips, h('div', { style: { marginTop: '12px' } }, table)));
}
