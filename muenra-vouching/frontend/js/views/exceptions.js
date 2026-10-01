// Pantalla 9: panel de excepciones agrupado por tipo de incidencia.
import { api } from '../api.js';
import { badge, fmt, h, STATUS_COLORS } from '../ui.js';

const GROUPS = ['EXCEPCIÓN', 'REVISIÓN MANUAL', 'POSIBLE DUPLICADO', 'SIN SOPORTE', 'DOCUMENTO ILEGIBLE', 'ERROR DE LECTURA'];

export async function render(ctx) {
  const pid = ctx.project.id;
  const [rows, docs] = await Promise.all([
    api(`/projects/${pid}/results`, { query: { exceptions_only: true } }),
    api(`/projects/${pid}/documents`),
  ]);
  const unref = docs.filter((d) => ['SOPORTE NO REFERENCIADO', 'POSIBLE DUPLICADO', 'DOCUMENTO ILEGIBLE', 'ERROR DE LECTURA'].includes(d.estado));
  const pending = rows.filter((r) => !r.decision_revisor).length;
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Panel de excepciones'),
      h('div', { class: 'muted' }, `${rows.length} partidas requieren atención · ${pending} sin decisión del revisor. Una excepción es una diferencia por investigar, no una conclusión de fraude.`))),
    GROUPS.map((g) => {
      const items = rows.filter((r) => r.estado === g);
      if (!items.length) return null;
      const value = items.reduce((a, r) => a + Number(r.valor_esperado || 0), 0);
      return h('div', { class: 'card', style: { borderLeft: `4px solid ${STATUS_COLORS[g]}` } },
        h('div', { class: 'page-head', style: { marginBottom: '8px' } }, h('div', { class: 'grow' }, h('h2', null, badge(g), ` ${items.length} partidas`)), h('div', { class: 'muted' }, `Valor esperado ${fmt.money(value)}`)),
        h('table', { class: 'table' },
          h('thead', null, h('tr', null, ['ID', 'Tercero', 'Número', 'Valor esperado', 'Valor extraído', 'Motivos', 'Revisión', ''].map((c) => h('th', null, c)))),
          h('tbody', null, items.map((r) => h('tr', null,
            h('td', null, h('b', null, r.id_muestra)), h('td', null, r.tercero_esperado), h('td', null, r.numero_esperado),
            h('td', { class: 'num' }, fmt.money(r.valor_esperado)), h('td', { class: 'num' }, fmt.money(r.valor_extraido)),
            h('td', null, h('ul', { class: 'reasons small' }, (r.motivos || []).map((m) => h('li', null, m)))),
            h('td', null, r.decision_revisor ? h('span', { class: 'small' }, `${r.decision_revisor} · ${r.revisor || ''}`) : h('span', { class: 'small muted' }, 'Pendiente')),
            h('td', null, h('a', { class: 'btn sm primary', href: `#/p/${pid}/visor/${r.id}` }, 'Revisar')))))));
    }),
    unref.length ? h('div', { class: 'card' }, h('h2', null, 'Documentos con incidencias o no referenciados'),
      h('table', { class: 'table' },
        h('thead', null, h('tr', null, ['Archivo', 'Estado', 'Tipo', 'Número', 'Valor', 'Observación', ''].map((c) => h('th', null, c)))),
        h('tbody', null, unref.map((d) => h('tr', null, h('td', null, d.archivo), h('td', null, badge(d.estado)), h('td', null, d.tipo_nombre || '—'),
          h('td', null, d.numero || '—'), h('td', { class: 'num' }, fmt.money(d.valor_total)), h('td', { class: 'small muted' }, d.observacion || ''),
          h('td', null, h('a', { class: 'btn sm', href: `#/p/${pid}/documento/${d.id}` }, 'Ver')))))))
      : null,
    !rows.length && !unref.length ? h('div', { class: 'card empty' }, '✓ No hay excepciones pendientes') : null);
}
