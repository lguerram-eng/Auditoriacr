// Pantalla 13: historial de auditoría (accesos, cambios y revisiones).
import { api, session } from '../api.js';
import { clear, fmt, h, sortableTable } from '../ui.js';

export async function render(ctx) {
  const global = ctx.mode === 'global';
  const box = h('div');
  let tab = 'eventos';
  const draw = async () => {
    if (tab === 'eventos') {
      const rows = await api(global ? '/audit' : `/projects/${ctx.project.id}/audit`, { query: { limit: 2000 } });
      clear(box, sortableTable([
        { key: 'fecha', label: 'Fecha', render: (r) => fmt.datetime(r.fecha), sort: (r) => r.fecha },
        { key: 'usuario', label: 'Usuario' },
        { key: 'accion', label: 'Acción', render: (r) => h('span', { class: r.exito ? '' : 'no' }, r.accion) },
        { key: 'entidad', label: 'Entidad', render: (r) => (r.entidad ? `${r.entidad} #${r.entidad_id || ''}` : '—') },
        { key: 'ip', label: 'IP' },
        { key: 'detalle', label: 'Detalle', render: (r) => h('div', { class: 'mono', style: { maxWidth: '520px', whiteSpace: 'pre-wrap' } }, Object.keys(r.detalle || {}).length ? JSON.stringify(r.detalle) : '') },
      ], rows, { pageSize: 100 }));
    } else {
      const rows = await api(`/projects/${ctx.project.id}/review-history`);
      clear(box, sortableTable([
        { key: 'fecha', label: 'Fecha', render: (r) => fmt.datetime(r.fecha), sort: (r) => r.fecha },
        { key: 'usuario', label: 'Usuario' },
        { key: 'accion', label: 'Acción' },
        { key: 'id_muestra', label: 'ID muestra' },
        { key: 'campo', label: 'Campo' },
        { key: 'valor_anterior', label: 'Valor anterior' },
        { key: 'valor_nuevo', label: 'Valor nuevo' },
        { key: 'comentario', label: 'Comentario' },
      ], rows, { pageSize: 100, empty: 'Sin revisiones registradas' }));
    }
  };
  await draw();
  const tabs = global ? null : h('div', { class: 'chips', style: { marginBottom: '12px' } },
    [['eventos', 'Accesos y cambios'], ['revisiones', 'Historial de revisiones']].map(([k, l]) => h('button', { class: `chip ${tab === k ? 'active' : ''}`,
      onclick: (e) => { tab = k; [...e.target.parentNode.children].forEach((c) => c.classList.remove('active')); e.target.classList.add('active'); draw(); } }, l)));
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, global ? 'Historial de auditoría — global' : 'Historial de auditoría'),
      h('div', { class: 'muted' }, 'Registro inmutable de accesos, cargas, procesamientos, cambios de configuración y decisiones de revisión.'))),
    !global && !session.can('auditoria.proyecto') ? h('div', { class: 'alert warn' }, 'Sin permiso para ver la auditoría') : h('div', { class: 'card' }, tabs, box));
}
