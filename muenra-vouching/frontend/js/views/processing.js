// Pantalla 6: procesamiento asíncrono con seguimiento de estado.
import { api, session } from '../api.js';
import { badge, clear, fmt, h, sortableTable, toast } from '../ui.js';

export async function render(ctx) {
  const pid = ctx.project.id;
  const summary = h('div');
  const table = h('div');
  let timer = null;
  const refresh = async () => {
    const [jobs, docs] = await Promise.all([api(`/projects/${pid}/jobs`), api(`/projects/${pid}/documents`)]);
    const total = jobs.total_documentos || 0;
    const done = total - (jobs.documentos_por_estado.PENDIENTE || 0) - (jobs.documentos_por_estado.PROCESANDO || 0);
    const pct = total ? Math.round((done / total) * 100) : 0;
    clear(summary,
      h('div', { class: 'kpis' },
        ['PENDIENTE', 'PROCESANDO', 'PROCESADO', 'ILEGIBLE', 'ERROR'].map((s) => h('div', { class: 'kpi' }, h('div', { class: 'label' }, s), h('div', { class: 'value' }, fmt.int(jobs.documentos_por_estado[s] || 0))))),
      h('div', { style: { margin: '14px 0 4px' } }, `Avance de lectura: ${done} de ${total} documentos (${pct} %)`, jobs.en_curso ? h('span', null, ' ', h('span', { class: 'spinner' })) : null),
      h('div', { class: 'progress' }, h('div', { style: { width: `${pct}%` } })),
      h('div', { class: 'small muted', style: { marginTop: '6px' } }, Object.entries(jobs.trabajos).map(([k, v]) => `${k}: ${v}`).join(' · ') || 'Sin trabajos'));
    clear(table, sortableTable([
      { key: 'id', label: '#', num: true },
      { key: 'archivo', label: 'Archivo' },
      { key: 'formato', label: 'Formato' },
      { key: 'estado_tecnico', label: 'Lectura' },
      { key: 'metodo', label: 'Método' },
      { key: 'tipo_nombre', label: 'Tipo documental' },
      { key: 'confianza_ocr', label: 'Conf. OCR', num: true, render: (r) => (r.confianza_ocr ? `${fmt.int(r.confianza_ocr)}` : '—') },
      { key: 'numero', label: 'Número' },
      { key: 'valor_total', label: 'Valor', num: true, render: (r) => fmt.money(r.valor_total) },
      { key: 'estado', label: 'Estado vouching', render: (r) => badge(r.estado) },
      { key: 'observacion', label: 'Observación', render: (r) => h('span', { class: 'small muted' }, r.observacion || '') },
      { key: 'acc', label: '', render: (r) => h('div', { class: 'btn-row' },
        h('a', { class: 'btn sm', href: `#/p/${pid}/documento/${r.id}`, onclick: (e) => e.stopPropagation() }, 'Ver'),
        session.can('procesar') && ['ERROR', 'ILEGIBLE', 'PROCESADO'].includes(r.estado_tecnico)
          ? h('button', { class: 'btn sm ghost', onclick: async (e) => { e.stopPropagation(); await api(`/documents/${r.id}/reprocess`, { method: 'POST' }); toast('Documento encolado'); refresh(); } }, 'Reprocesar') : null) },
    ], docs, { empty: 'Aún no hay documentos cargados' }));
    clearTimeout(timer);
    timer = setTimeout(refresh, jobs.en_curso ? 2000 : 10000);
  };
  ctx.onLeave(() => clearTimeout(timer));
  await refresh();
  const act = async (path, msg) => { try { await api(path, { method: 'POST' }); toast(msg, 'ok'); refresh(); } catch (err) { toast(err.message, 'bad'); } };
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Procesamiento de documentos'),
      h('div', { class: 'muted' }, 'Prioridad: XML determinístico → PDF con texto → OCR → IA (opcional). Se conserva evidencia por campo.')),
    session.can('procesar') ? h('div', { class: 'btn-row' },
      h('button', { class: 'btn', onclick: () => act(`/projects/${pid}/process?reprocess_errors=true`, 'Documentos con error encolados') }, 'Reintentar errores'),
      h('button', { class: 'btn', onclick: () => act(`/projects/${pid}/process`, 'Pendientes encolados') }, 'Procesar pendientes'),
      h('button', { class: 'btn primary', onclick: () => act(`/projects/${pid}/match`, 'Motor de vouching encolado') }, 'Ejecutar vouching')) : null),
    h('div', { class: 'card' }, summary),
    h('div', { class: 'card' }, h('h2', null, 'Documentos'), table));
}
