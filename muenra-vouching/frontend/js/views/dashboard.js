// Pantalla 2: tablero principal (global y por proyecto).
import { api } from '../api.js';
import { badge, fmt, h, kpi, STATUS_COLORS, STATUSES } from '../ui.js';

export async function render(ctx) {
  if (!ctx.project) return globalDashboard(ctx);
  const pid = ctx.project.id;
  const [k, rec] = await Promise.all([api(`/projects/${pid}/dashboard`), api(`/projects/${pid}/reconciliation`)]);
  const total = k.total_partidas || 0;
  const dist = STATUSES.filter((s) => s !== 'SOPORTE NO REFERENCIADO').map((s) => ({ s, n: (k.por_estado[s] || {}).partidas || 0, v: (k.por_estado[s] || {}).valor || 0 }));
  const docDist = Object.entries(k.documentos_por_estado || {});
  const docTotal = docDist.reduce((a, [, n]) => a + n, 0);
  const check = (label, ok, detail) => h('li', null, h('span', null, label, detail ? h('div', { class: 'small muted' }, detail) : null), h('span', { class: ok ? 'yes' : 'no' }, ok ? '✓ Cuadra' : '✗ Revisar'));
  const val = rec.validacion_resultado_esperado;

  return h('div', null,
    h('div', { class: 'page-head' },
      h('div', { class: 'grow' }, h('h1', null, 'Tablero principal'), h('div', { class: 'muted' }, `${ctx.project.client_name} · Periodo ${ctx.project.period || '—'}`)),
      h('div', { class: 'btn-row' },
        h('a', { class: 'btn', href: `#/p/${pid}/resultados` }, 'Ver resultados'),
        h('a', { class: 'btn gold', href: `#/p/${pid}/excepciones` }, `Excepciones (${k.excepciones + k.revision_manual + k.posibles_duplicados})`))),
    k.trabajos_pendientes ? h('div', { class: 'alert info' }, `Procesamiento en curso: ${k.trabajos_pendientes} trabajos en cola. `, h('a', { href: `#/p/${pid}/procesamiento` }, 'Ver avance')) : null,
    h('div', { class: 'kpis' },
      kpi('Total de partidas', fmt.int(k.total_partidas), null, '#1f3a5f'),
      kpi('Total de documentos', fmt.int(k.total_documentos), `${fmt.int(k.documentos_procesados)} procesados`, '#2b4d78'),
      kpi('Partidas con soporte', fmt.int(k.partidas_con_soporte), null, '#2a5ea8'),
      kpi('Partidas sin soporte', fmt.int(k.partidas_sin_soporte), null, STATUS_COLORS['SIN SOPORTE']),
      kpi('Coincidencias completas', fmt.int(k.coincidencias), null, STATUS_COLORS.COINCIDE),
      kpi('Coincidencias con tolerancia', fmt.int(k.coincidencias_tolerancia), null, STATUS_COLORS['COINCIDE CON TOLERANCIA']),
      kpi('Excepciones', fmt.int(k.excepciones), 'Diferencias por investigar', STATUS_COLORS['EXCEPCIÓN']),
      kpi('Pendientes de revisión', fmt.int(k.pendientes_revision), `${fmt.int(k.revisadas)} revisadas`, STATUS_COLORS['REVISIÓN MANUAL']),
      kpi('Documentos ilegibles', fmt.int(k.documentos_ilegibles), `${fmt.int(k.documentos_error)} con error de lectura`, STATUS_COLORS['DOCUMENTO ILEGIBLE']),
      kpi('Valor total de la población', fmt.money(k.valor_poblacion), null, '#1f3a5f'),
      kpi('Valor soportado', fmt.money(k.valor_soportado), null, STATUS_COLORS.COINCIDE),
      kpi('Valor con diferencias', fmt.money(k.valor_con_diferencias), `Σ |diferencias| ${fmt.money(k.suma_diferencias_absolutas)}`, STATUS_COLORS['EXCEPCIÓN']),
      kpi('Porcentaje de cobertura', `${fmt.int(k.porcentaje_cobertura)} %`, 'Valor soportado / población', '#c8963e'),
      kpi('Porcentaje de avance', `${fmt.int(k.porcentaje_avance)} %`, 'Procesamiento y revisión', '#c8963e')),
    h('div', { class: 'grid grid-2', style: { marginTop: '16px' } },
      h('div', { class: 'card' }, h('h2', null, 'Partidas por estado'),
        h('div', { class: 'stackbar' }, dist.filter((d) => d.n).map((d) => h('span', { title: `${d.s}: ${d.n}`, style: { width: `${(d.n / Math.max(total, 1)) * 100}%`, background: STATUS_COLORS[d.s] } }))),
        h('table', { class: 'table', style: { marginTop: '12px' } },
          h('thead', null, h('tr', null, h('th', null, 'Estado'), h('th', { class: 'num' }, 'Partidas'), h('th', { class: 'num' }, 'Valor esperado'))),
          h('tbody', null, dist.filter((d) => d.n).map((d) => h('tr', { class: 'clickable', onclick: () => ctx.navigate(`/p/${pid}/resultados?estado=${encodeURIComponent(d.s)}`) },
            h('td', null, badge(d.s)), h('td', { class: 'num' }, fmt.int(d.n)), h('td', { class: 'num' }, fmt.money(d.v))))))),
      h('div', { class: 'card' }, h('h2', null, 'Conciliación antes y después del procesamiento'),
        h('ul', { class: 'check-list' },
          check('Conteo de partidas', rec.conteo_partidas_cuadra, `Importadas ${rec.partidas_importadas} · en base ${rec.partidas_en_base} · con resultado ${rec.partidas_con_resultado}`),
          check('Suma de VALOR_ESPERADO', rec.valor_cuadra, `Importado ${fmt.money(rec.valor_importado)} · por estados ${fmt.money(rec.valor_por_estados)}`),
          check('Documentos sin pérdida', rec.conteo_documentos_cuadra, `Cargados ${rec.documentos_cargados} · clasificados ${rec.documentos_clasificados}`),
          check('Sin duplicación de valor', rec.sin_duplicacion_de_valor, `Valor asignado a soportes ${fmt.money(rec.valor_asignado_a_soportes)}`),
          val ? check('Validación contra RESULTADO_ESPERADO', val.precision === 1, `${val.coinciden} de ${val.total} (${fmt.pct(val.precision)})`) : null),
        h('h3', { style: { marginTop: '16px' } }, 'Documentos por estado'),
        h('div', { class: 'stackbar' }, docDist.map(([s, n]) => h('span', { title: `${s}: ${n}`, style: { width: `${(n / Math.max(docTotal, 1)) * 100}%`, background: STATUS_COLORS[s] || '#999' } }))),
        h('div', { class: 'legend' }, docDist.map(([s, n]) => h('span', null, h('i', { class: 'dot', style: { background: STATUS_COLORS[s] || '#999' } }), `${s}: ${n}`))),
        h('p', { class: 'small muted', style: { marginTop: '14px' } }, k.ultima_ejecucion ? `Última ejecución del motor: ${fmt.datetime(k.ultima_ejecucion.fecha)}` : 'El motor aún no se ha ejecutado.'),
        h('p', { class: 'small muted' }, 'Una excepción representa una diferencia que debe investigarse; no constituye por sí misma un indicio de fraude.'))));
}

async function globalDashboard(ctx) {
  const rows = await api('/dashboard');
  const sum = (k) => rows.reduce((a, r) => a + Number(r[k] || 0), 0);
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Tablero principal — Muenra Vouching'), h('div', { class: 'muted' }, 'Resumen de todos sus proyectos de vouching'))),
    h('div', { class: 'kpis' },
      kpi('Proyectos', fmt.int(rows.length)),
      kpi('Total de partidas', fmt.int(sum('total_partidas'))),
      kpi('Total de documentos', fmt.int(sum('total_documentos'))),
      kpi('Excepciones', fmt.int(sum('excepciones')), null, STATUS_COLORS['EXCEPCIÓN']),
      kpi('Pendientes de revisión', fmt.int(sum('pendientes_revision')), null, STATUS_COLORS['REVISIÓN MANUAL']),
      kpi('Valor de las poblaciones', fmt.money(sum('valor_poblacion'))),
      kpi('Valor soportado', fmt.money(sum('valor_soportado')), null, STATUS_COLORS.COINCIDE)),
    h('div', { class: 'card', style: { marginTop: '16px' } }, h('h2', null, 'Proyectos'),
      rows.length ? h('table', { class: 'table' },
        h('thead', null, h('tr', null, ['Código', 'Proyecto', 'Cliente', 'Partidas', 'Documentos', 'Coinciden', 'Excepciones', 'Cobertura', 'Avance'].map((c) => h('th', null, c)))),
        h('tbody', null, rows.map((r) => h('tr', { class: 'clickable', onclick: () => ctx.navigate(`/p/${r.proyecto.id}/tablero`) },
          h('td', null, r.proyecto.code), h('td', null, r.proyecto.name), h('td', null, r.proyecto.client_name),
          h('td', { class: 'num' }, fmt.int(r.total_partidas)), h('td', { class: 'num' }, fmt.int(r.total_documentos)),
          h('td', { class: 'num' }, fmt.int(r.coincidencias + r.coincidencias_tolerancia)), h('td', { class: 'num' }, fmt.int(r.excepciones)),
          h('td', { class: 'num' }, `${fmt.int(r.porcentaje_cobertura)} %`), h('td', { class: 'num' }, `${fmt.int(r.porcentaje_avance)} %`)))))
        : h('div', { class: 'empty' }, 'Aún no tiene proyectos. ', h('a', { href: '#/proyectos/nuevo' }, 'Cree el primero'))));
}
