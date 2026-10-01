// Pantalla 4: importación del Excel de referencia con informe de integridad.
import { api } from '../api.js';
import { clear, fmt, h, toast } from '../ui.js';

export async function render(ctx) {
  const pid = ctx.project.id;
  const report = h('div');
  const reportCard = h('div', { class: 'card hidden' }, report);
  const history = h('div');
  let file = null;
  const name = h('div', { class: 'muted' }, 'Ningún archivo seleccionado');
  const input = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', class: 'hidden', onchange: (e) => pick(e.target.files[0]) });
  const pick = (f) => { file = f; clear(name, f ? `${f.name} (${fmt.int(Math.round(f.size / 1024))} KB)` : 'Ningún archivo seleccionado'); clear(report); };
  const drop = h('div', { class: 'dropzone', onclick: () => input.click(),
    ondragover: (e) => { e.preventDefault(); drop.classList.add('over'); }, ondragleave: () => drop.classList.remove('over'),
    ondrop: (e) => { e.preventDefault(); drop.classList.remove('over'); pick(e.dataTransfer.files[0]); } },
  h('div', { class: 'big' }, '📊'), h('div', null, h('b', null, 'Arrastre aquí el Excel de referencia'), ' o haga clic para seleccionarlo'),
  h('div', { class: 'small muted' }, 'Hojas: CONFIGURACION, REFERENCIA_VOUCHING (obligatoria), ENTIDADES_ALIAS, TIPOS_DOCUMENTO, CAMPOS_EXTRACCION, RESULTADO_ESPERADO · También CSV'), name, input);

  const send = async (dry) => {
    if (!file) { toast('Seleccione un archivo', 'bad'); return; }
    const form = new FormData();
    form.append('file', file);
    reportCard.classList.remove('hidden');
    clear(report, h('div', { class: 'empty' }, h('span', { class: 'spinner' }), dry ? ' Validando…' : ' Importando…'));
    try {
      const r = await api(`/projects/${pid}/imports`, { method: 'POST', form, query: { dry_run: dry } });
      clear(report, renderReport(r.informe, r.importado));
      if (r.importado) { toast('Importación completada', 'ok'); loadHistory(); }
    } catch (err) {
      if (err.data && err.data.informe) clear(report, renderReport(err.data.informe, false));
      else clear(report, h('div', { class: 'alert bad' }, err.message));
    }
  };
  const loadHistory = async () => {
    const rows = await api(`/projects/${pid}/imports`);
    clear(history, rows.length ? h('table', { class: 'table' },
      h('thead', null, h('tr', null, ['Fecha', 'Archivo', 'SHA-256', 'Estado', 'Partidas', 'Valor total', 'Activa'].map((c) => h('th', null, c)))),
      h('tbody', null, rows.map((b) => h('tr', null, h('td', null, fmt.datetime(b.fecha)), h('td', null, b.archivo), h('td', { class: 'mono' }, b.sha256.slice(0, 16) + '…'),
        h('td', null, b.estado), h('td', { class: 'num' }, fmt.int(b.partidas)), h('td', { class: 'num' }, fmt.money(b.valor_total)), h('td', null, b.activa ? 'Sí' : 'No')))))
      : h('div', { class: 'empty' }, 'Sin importaciones previas'));
  };
  loadHistory();
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Importar Excel de referencia'),
      h('div', { class: 'muted' }, 'Antes de importar, valide columnas, duplicados, fechas, valores y totales de control.'))),
    h('div', { class: 'card' }, drop,
      h('div', { class: 'btn-row', style: { marginTop: '14px' } },
        h('button', { class: 'btn', onclick: () => send(true) }, '1. Validar integridad'),
        h('button', { class: 'btn primary', onclick: () => send(false) }, '2. Importar'),
        h('span', { class: 'small muted' }, 'Una nueva importación reemplaza la población vigente del proyecto (queda registrada en la auditoría).'))),
    reportCard,
    h('div', { class: 'card' }, h('h2', null, 'Historial de importaciones'), history));
}

export function renderReport(r, imported) {
  if (!r) return h('div');
  const cls = r.estado === 'RECHAZADA' ? 'bad' : r.estado === 'VALIDA' ? 'ok' : 'warn';
  const rec = r.conciliacion || {};
  return h('div', null,
    h('h2', null, 'Informe de integridad de la carga'),
    h('div', { class: `alert ${cls}` }, h('b', null, `Estado: ${r.estado}`), imported ? ' — datos importados' : ' — no se importaron datos'),
    h('div', { class: 'kpis' },
      h('div', { class: 'kpi' }, h('div', { class: 'label' }, 'Partidas válidas'), h('div', { class: 'value' }, fmt.int(r.resumen.partidas))),
      h('div', { class: 'kpi' }, h('div', { class: 'label' }, 'Suma VALOR_ESPERADO'), h('div', { class: 'value' }, fmt.money(r.resumen.valor_total))),
      h('div', { class: 'kpi' }, h('div', { class: 'label' }, 'Registros leídos'), h('div', { class: 'value' }, fmt.int(rec.registros_leidos))),
      rec.control_total_registros !== undefined ? h('div', { class: 'kpi' }, h('div', { class: 'label' }, 'Control de registros'),
        h('div', { class: 'value' }, fmt.int(rec.control_total_registros)), h('div', { class: rec.registros_cuadran ? 'yes' : 'no' }, rec.registros_cuadran ? '✓ Concilia' : '✗ No concilia')) : null,
      rec.control_total_valor !== undefined ? h('div', { class: 'kpi' }, h('div', { class: 'label' }, 'Control de valor'),
        h('div', { class: 'value' }, fmt.money(rec.control_total_valor)), h('div', { class: rec.valor_cuadra ? 'yes' : 'no' }, rec.valor_cuadra ? '✓ Concilia' : '✗ No concilia')) : null),
    h('div', { class: 'grid grid-2', style: { marginTop: '12px' } },
      h('div', null, h('h3', null, 'Hojas'), h('ul', { class: 'check-list' }, Object.entries(r.hojas_reconocidas || {}).map(([s, ok]) =>
        h('li', null, s, h('span', { class: ok ? 'yes' : 'muted' }, ok ? '✓ encontrada' : 'no encontrada'))))),
      h('div', null, h('h3', null, 'Columnas mapeadas'), h('ul', { class: 'check-list' }, Object.entries(r.columnas_mapeadas || {}).map(([c, src]) =>
        h('li', null, c, h('span', { class: 'muted' }, src)))))),
    r.errores.length ? h('div', { class: 'alert bad', style: { marginTop: '12px' } }, h('b', null, 'Errores bloqueantes'), h('ul', null, r.errores.map((e) => h('li', null, e)))) : null,
    r.advertencias.length ? h('div', { class: 'alert warn' }, h('b', null, 'Advertencias'), h('ul', null, r.advertencias.map((e) => h('li', null, e)))) : null,
    (r.informacion || []).length ? h('div', { class: 'alert info' }, h('ul', null, r.informacion.map((e) => h('li', null, e)))) : null,
    r.incidencias_por_fila.length ? h('details', { open: r.incidencias_por_fila.length < 30 },
      h('summary', null, `Incidencias por fila (${r.total_incidencias})`),
      h('div', { class: 'table-wrap', style: { marginTop: '8px' } }, h('table', { class: 'table' },
        h('thead', null, h('tr', null, ['Fila', 'Columna', 'Nivel', 'Detalle'].map((c) => h('th', null, c)))),
        h('tbody', null, r.incidencias_por_fila.map((i) => h('tr', null, h('td', null, i.fila), h('td', null, i.columna),
          h('td', { class: i.nivel === 'ERROR' ? 'no' : '' }, i.nivel), h('td', null, i.detalle))))))) : null,
    Object.keys(r.resumen.parametros || {}).length ? h('details', null, h('summary', null, 'Parámetros leídos de CONFIGURACION'),
      h('pre', { class: 'json' }, JSON.stringify(r.resumen.parametros, null, 2))) : null);
}
