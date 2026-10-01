// Pantalla 3: proyectos y creación de proyecto de vouching.
import { api, session } from '../api.js';
import { field, fmt, h, toast } from '../ui.js';

export async function render(ctx) {
  if (ctx.mode === 'new') return createForm(ctx);
  const projects = await ctx.refreshProjects();
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Proyectos de vouching')),
      session.can('proyecto.crear') ? h('a', { class: 'btn primary', href: '#/proyectos/nuevo' }, '＋ Crear proyecto') : null),
    h('div', { class: 'grid grid-3' }, projects.length ? projects.map((p) => h('div', { class: 'card' },
      h('div', { class: 'small muted' }, p.code), h('h2', null, p.name),
      h('div', null, p.client_name, p.client_nit ? h('span', { class: 'muted' }, ` · NIT ${p.client_nit}`) : null),
      h('div', { class: 'small muted', style: { margin: '8px 0' } }, `Periodo ${p.period || '—'} · ${fmt.int(p.partidas)} partidas · ${fmt.int(p.documentos)} documentos · ${p.status}`),
      h('div', { class: 'small muted' }, `IA: ${p.allow_ai_processing ? 'autorizada' : 'no autorizada'} · Entrenamiento con documentos: bloqueado · Retención ${p.retention_days} días`),
      h('div', { class: 'btn-row', style: { marginTop: '12px' } }, h('a', { class: 'btn primary sm', href: `#/p/${p.id}/tablero` }, 'Abrir'))))
      : h('div', { class: 'card empty' }, 'No hay proyectos asignados.')));
}

async function createForm(ctx) {
  let users = [];
  try { users = await api('/users'); } catch { /* */ }
  const me = session.user;
  const inp = (name, attrs = {}) => h('input', { class: 'input', name, ...attrs });
  const code = inp('code', { required: true, pattern: '[A-Za-z0-9_.-]+', placeholder: 'AUD-2026-001' });
  const name = inp('name', { required: true, placeholder: 'Vouching de compras 2026' });
  const client = inp('client', { required: true, placeholder: 'Razón social del cliente' });
  const nit = inp('nit', { placeholder: '900.000.000-0' });
  const period = inp('period', { placeholder: '2026-12' });
  const desc = h('textarea', { class: 'input', rows: 3 });
  const ai = h('input', { type: 'checkbox' });
  const retention = inp('retention', { type: 'number', min: 30, max: 36500, value: 3650 });
  const members = users.filter((u) => u.id !== me.id).map((u) => ({ u, cb: h('input', { type: 'checkbox', value: u.id }) }));
  const submit = async (e) => {
    e.preventDefault();
    try {
      const p = await api('/projects', { method: 'POST', body: {
        code: code.value.trim(), name: name.value.trim(), client_name: client.value.trim(), client_nit: nit.value.trim() || null,
        period: period.value.trim() || null, description: desc.value.trim() || null, allow_ai_processing: ai.checked,
        retention_days: Number(retention.value), member_ids: members.filter((m) => m.cb.checked).map((m) => m.u.id) } });
      toast('Proyecto creado', 'ok');
      await ctx.refreshProjects();
      ctx.navigate(`/p/${p.id}/importar`);
    } catch (err) { toast(err.message, 'bad'); }
  };
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Crear proyecto de vouching'), h('div', { class: 'muted' }, 'Cada proyecto aísla la información de un cliente o encargo.'))),
    h('form', { class: 'card', onsubmit: submit, style: { maxWidth: '900px' } },
      h('div', { class: 'form-grid' },
        field('Código del proyecto *', code, 'Letras, números, punto, guion o guion bajo'), field('Nombre *', name),
        field('Cliente / entidad auditada *', client), field('NIT del cliente', nit, 'Permite distinguir facturas de venta y de compra'),
        field('Periodo auditado', period), field('Retención de la información (días)', retention, 'Al cerrar el proyecto, se eliminará tras este plazo')),
      field('Descripción', desc),
      h('div', { class: 'check', style: { marginBottom: '8px' } }, ai, h('span', null, 'Autorizar procesamiento con IA (solo completa campos faltantes con cita literal verificada)')),
      h('div', { class: 'alert info small' }, 'Los documentos de este proyecto nunca se utilizarán para entrenamiento de modelos.'),
      members.length ? h('div', null, h('h3', null, 'Miembros del equipo'),
        h('div', { class: 'grid grid-3' }, members.map((m) => h('label', { class: 'check' }, m.cb, `${m.u.full_name} (${m.u.role})`)))) : null,
      h('div', { class: 'btn-row', style: { marginTop: '16px' } }, h('button', { class: 'btn primary', type: 'submit' }, 'Crear proyecto'))));
}
