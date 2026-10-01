// Pantalla 14: exportaciones a Excel.
import { api, download } from '../api.js';
import { h, toast } from '../ui.js';

const DESCR = {
  resumen: 'Indicadores, distribución por estado y conciliación', resultado: 'Todas las partidas con todos los campos comparados',
  coincidencias: 'COINCIDE y COINCIDE CON TOLERANCIA', excepciones: 'Excepciones, revisión manual, duplicados, ilegibles',
  sin_soporte: 'Partidas sin documento', no_referenciados: 'Documentos sin partida y documentos con incidencias',
  evidencias: 'Cada campo extraído con página, región, texto, método y confianza', historial: 'Decisiones y correcciones de los revisores',
  parametros: 'Tolerancias, umbrales, pesos y políticas', registro: 'Versiones de OCR, reglas, IA y detalle técnico por documento',
};

export async function render(ctx) {
  const pid = ctx.project.id;
  const sheets = await api('/export/sheets');
  const checks = Object.entries(sheets).map(([k, label]) => ({ k, cb: h('input', { type: 'checkbox', checked: true }), label }));
  const go = async (only) => {
    try {
      toast('Generando archivo Excel…');
      await download(`/projects/${pid}/export`, only ? { sheets: only.join(',') } : undefined);
    } catch (err) { toast(err.message, 'bad'); }
  };
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Exportaciones'),
      h('div', { class: 'muted' }, 'Archivo Excel con la marca Muenra Vouching, protegido contra inyección de fórmulas. Cada exportación queda en el historial de auditoría.'))),
    h('div', { class: 'card' },
      h('div', { class: 'grid grid-2' }, checks.map((c, i) => h('label', { class: 'check', style: { alignItems: 'flex-start' } }, c.cb,
        h('div', null, h('b', null, `${i + 1}. ${c.label}`), h('div', { class: 'small muted' }, DESCR[c.k]))))),
      h('div', { class: 'btn-row', style: { marginTop: '16px' } },
        h('button', { class: 'btn primary', onclick: () => go(null) }, '⤓ Exportar todo'),
        h('button', { class: 'btn', onclick: () => { const sel = checks.filter((c) => c.cb.checked).map((c) => c.k); if (sel.length) go(sel); } }, 'Exportar selección'))));
}
