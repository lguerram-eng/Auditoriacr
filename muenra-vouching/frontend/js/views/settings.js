// Pantalla 11: configuración del proyecto (tolerancias, umbrales, pesos, IA, retención) y cuenta del usuario.
import { api, session } from '../api.js';
import { field, h, toast } from '../ui.js';

export async function render(ctx) {
  const pid = ctx.project.id;
  const s = await api(`/projects/${pid}/settings`);
  const p = s.parametros;
  const canEdit = session.can('configurar');
  const num = (v, attrs = {}) => h('input', { class: 'input', type: 'number', value: v, disabled: !canEdit, ...attrs });
  const tol = num(Math.round(p.TOLERANCIA_VALOR * 10000) / 100, { step: 0.5, min: 0, max: 100 });
  const zero = h('select', { class: 'input', disabled: !canEdit },
    h('option', { value: 'REVISION_MANUAL', selected: p.POLITICA_VALOR_CERO === 'REVISION_MANUAL' }, 'Enviar a REVISIÓN MANUAL'),
    h('option', { value: 'EXACTA', selected: p.POLITICA_VALOR_CERO === 'EXACTA' }, 'Exigir coincidencia exacta (si difiere: EXCEPCIÓN)'));
  const nameThr = num(p.UMBRAL_NOMBRE, { min: 0, max: 100 });
  const days = num(p.TOLERANCIA_DIAS, { min: 0, max: 365 });
  const rel = num(Math.round(p.UMBRAL_RELACION * 100), { min: 0, max: 100 });
  const ocr = num(p.UMBRAL_CONFIANZA_OCR, { min: 0, max: 100 });
  const maxSum = num(p.MAX_SOPORTES_SUMA, { min: 1, max: 8 });
  const dateReq = h('input', { type: 'checkbox', checked: p.EXIGIR_FECHA_EN_TOLERANCIA, disabled: !canEdit });
  const weights = Object.entries(p.PESOS).map(([k, v]) => ({ k, input: h('input', { type: 'range', min: 0, max: 50, value: v, disabled: !canEdit }), out: h('b', null, v) }));
  weights.forEach((w) => w.input.addEventListener('input', () => { w.out.textContent = w.input.value; }));
  const ai = h('input', { type: 'checkbox', checked: s.permite_ia, disabled: !session.can('proyecto.editar') });
  const retention = num(s.retencion_dias, { min: 30, max: 36500, disabled: !session.can('proyecto.editar') });

  const save = async () => {
    try {
      await api(`/projects/${pid}/settings`, { method: 'PUT', body: { parameters: {
        TOLERANCIA_VALOR: Number(tol.value) / 100, POLITICA_VALOR_CERO: zero.value, UMBRAL_NOMBRE: Number(nameThr.value),
        TOLERANCIA_DIAS: Number(days.value), UMBRAL_RELACION: Number(rel.value) / 100, UMBRAL_CONFIANZA_OCR: Number(ocr.value),
        MAX_SOPORTES_SUMA: Number(maxSum.value), EXIGIR_FECHA_EN_TOLERANCIA: dateReq.checked, MONEDA_BASE: p.MONEDA_BASE,
        PESOS: Object.fromEntries(weights.map((w) => [w.k, Number(w.input.value)])) } } });
      if (session.can('proyecto.editar')) await api(`/projects/${pid}`, { method: 'PATCH', body: { allow_ai_processing: ai.checked, retention_days: Number(retention.value) } });
      toast('Configuración guardada. Ejecute el motor de vouching para aplicarla.', 'ok');
    } catch (err) { toast(err.message, 'bad'); }
  };

  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Configuración'), h('div', { class: 'muted' }, 'Parámetros del motor de coincidencias para este proyecto.')),
      canEdit ? h('div', { class: 'btn-row' },
        h('button', { class: 'btn', onclick: async () => { await api(`/projects/${pid}/match`, { method: 'POST' }); toast('Motor encolado', 'ok'); } }, 'Ejecutar vouching'),
        h('button', { class: 'btn primary', onclick: save }, 'Guardar')) : null),
    h('div', { class: 'grid grid-2' },
      h('div', { class: 'card' }, h('h2', null, 'Regla de valor'),
        field('Tolerancia de valor (%)', tol, 'Predeterminada 10 %. DIFERENCIA_PORCENTAJE = |extraído − esperado| / |esperado|'),
        field('Si el valor esperado es cero', zero, 'Nunca se divide entre cero'),
        h('h2', { style: { marginTop: '18px' } }, 'Umbrales'),
        field('Similitud mínima de nombres (%)', nameThr, 'Predeterminado 85 %, tras normalizar y retirar sufijos jurídicos'),
        field('Tolerancia de fechas (días)', days),
        h('label', { class: 'check', style: { marginBottom: '12px' } }, dateReq, 'Fecha fuera de tolerancia genera EXCEPCIÓN'),
        field('Puntaje mínimo de relación (0-100)', rel),
        field('Confianza mínima OCR para aceptar automáticamente', ocr),
        field('Máximo de soportes a combinar para alcanzar el valor', maxSum)),
      h('div', null,
        h('div', { class: 'card' }, h('h2', null, 'Pesos por criterio'),
          h('p', { class: 'small muted' }, 'El puntaje de relación es el promedio ponderado de los criterios evaluables. El NIT tiene prioridad: nunca se aprueba solo por nombre si el NIT lo contradice.'),
          weights.map((w) => h('div', { class: 'field' }, h('label', null, `${s.criterios[w.k] || w.k} · prioridad ${s.prioridades[w.k] || ''}`),
            h('div', { class: 'btn-row' }, w.input, w.out)))),
        h('div', { class: 'card' }, h('h2', null, 'Privacidad y retención'),
          h('label', { class: 'check', style: { marginBottom: '8px' } }, ai, 'Autorizar IA para completar campos faltantes (requiere configuración del servidor)'),
          h('div', { class: 'alert info small' }, 'Uso de documentos para entrenamiento: BLOQUEADO (no configurable).'),
          field('Retención (días tras el cierre del proyecto)', retention)))),
    h('div', { class: 'card' }, h('h2', null, 'Mi cuenta'), passwordForm()));
}

function passwordForm() {
  const cur = h('input', { class: 'input', type: 'password', autocomplete: 'current-password' });
  const nw = h('input', { class: 'input', type: 'password', autocomplete: 'new-password' });
  return h('form', { style: { maxWidth: '420px' }, onsubmit: async (e) => {
    e.preventDefault();
    try { await api('/auth/password/change', { method: 'POST', body: { current_password: cur.value, new_password: nw.value } }); toast('Contraseña actualizada', 'ok'); cur.value = ''; nw.value = ''; } catch (err) { toast(err.message, 'bad'); }
  } }, field('Contraseña actual', cur), field('Nueva contraseña', nw, 'Mínimo 10 caracteres con mayúscula, minúscula, número y símbolo'),
  h('button', { class: 'btn', type: 'submit' }, 'Cambiar contraseña'));
}
