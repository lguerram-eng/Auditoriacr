// Pantalla 8: visor documental con evidencia resaltada, comparación por criterio y revisión humana.
import { api, imageUrl, session } from '../api.js';
import { badge, clear, field, fmt, h, modal, scoreTag, toast } from '../ui.js';
import { decide } from './review.js';

const FIELD_LABELS = {
  tipo_documento: 'Tipo de documento', numero_documento: 'Número del documento', fecha_emision: 'Fecha de emisión',
  fecha_vencimiento: 'Fecha de vencimiento', emisor_nombre: 'Emisor', receptor_nombre: 'Receptor', emisor_nit: 'NIT emisor',
  emisor_dv: 'DV emisor', receptor_nit: 'NIT receptor', receptor_dv: 'DV receptor', subtotal: 'Subtotal', base_gravable: 'Base gravable',
  iva: 'IVA', retenciones: 'Retenciones', descuentos: 'Descuentos', valor_total: 'Valor total', moneda: 'Moneda', cufe: 'CUFE / CUDE',
  cude: 'CUDE', numero_contrato: 'Número de contrato', orden_compra: 'Orden de compra', centro_costo: 'Centro de coste', concepto: 'Concepto',
  forma_pago: 'Forma de pago', cuenta_bancaria: 'Cuenta bancaria', firmantes: 'Firmantes', vigencia_contrato: 'Vigencia del contrato',
  objeto_contractual: 'Objeto contractual', documento_referencia: 'Documento referenciado',
};
const label = (k) => FIELD_LABELS[k] || k.replace(/_/g, ' ');
const IMAGE_TYPES = ['pdf', 'png', 'jpg', 'tiff'];

export async function render(ctx) {
  const pid = ctx.project.id;
  const isResult = ctx.mode === 'result';
  const result = isResult ? await api(`/results/${ctx.params[1]}`) : null;
  const docs = isResult ? result.documentos : [{ id: ctx.params[1], rol: 'DOCUMENTO' }];
  const state = { docId: docs.length ? docs[0].id : null, doc: null, page: 1, activeField: null };

  const docPane = h('div', { class: 'doc-pane' });
  const side = h('div', { class: 'side-pane' });

  const loadDoc = async (id, focus) => {
    state.docId = id;
    state.doc = await api(`/documents/${id}`);
    state.page = 1;
    state.activeField = null;
    if (focus) applyFocus(focus);
    else {
      const vf = state.doc.campos.find((f) => f.campo === 'valor_total');
      if (vf) { state.activeField = vf.id; state.page = vf.pagina || 1; }
    }
    drawDoc();
    drawSide();
  };

  const applyFocus = ({ fieldId, page }) => {
    const f = state.doc && state.doc.campos.find((x) => x.id === fieldId);
    state.activeField = fieldId || null;
    state.page = (f && f.pagina) || page || state.page;
  };

  const focusField = (fieldId, docId) => {
    if (docId && docId !== state.docId) { loadDoc(docId, { fieldId }); return; }
    applyFocus({ fieldId });
    drawDoc();
    drawSide();
  };

  // ------------------------------------------------------------------ documento
  const drawDoc = async () => {
    const d = state.doc;
    const tabs = h('div', { class: 'doc-tabs' }, docs.map((x) => h('button', { class: x.id === state.docId ? 'active' : '', onclick: () => loadDoc(x.id) },
      `${x.rol === 'PRINCIPAL' ? '★ ' : x.rol === 'COMPLEMENTARIO' ? '＋ ' : x.rol === 'REPRESENTACION_GRAFICA' ? '⎘ ' : ''}${x.archivo || `Documento ${x.id}`}`)));
    if (!d) { clear(docPane, tabs, h('div', { class: 'empty', style: { color: '#fff' } }, 'Esta partida no tiene documentos relacionados.')); return; }
    const pages = Math.max(1, d.paginas || 1);
    const nav = h('div', { class: 'page-nav' },
      h('button', { class: 'btn sm', disabled: state.page <= 1, onclick: () => { state.page -= 1; drawDoc(); } }, '‹'),
      h('span', null, `Página ${state.page} de ${pages}`),
      h('button', { class: 'btn sm', disabled: state.page >= pages, onclick: () => { state.page += 1; drawDoc(); } }, '›'),
      h('span', { class: 'small', style: { marginLeft: 'auto', opacity: 0.85 } }, `${d.metodo || ''} · ${d.formato.toUpperCase()}`),
      h('button', { class: 'btn sm', onclick: () => import('../api.js').then((m) => m.download(`/documents/${d.id}/file`)) }, 'Descargar original'));
    const pageFields = d.campos.filter((f) => (f.pagina || 1) === state.page);
    let pageEl;
    if (IMAGE_TYPES.includes(d.formato) && d.estado_tecnico !== 'ERROR') {
      const img = h('img', { alt: `Página ${state.page} de ${d.archivo}` });
      pageEl = h('div', { class: 'page' }, img,
        pageFields.filter((f) => f.region).map((f) => h('div', {
          class: `hl ${f.id === state.activeField ? 'active' : ''}`, title: `${label(f.campo)}: ${f.valor}`,
          style: { left: `${f.region[0] * 100}%`, top: `${f.region[1] * 100}%`, width: `${(f.region[2] - f.region[0]) * 100}%`, height: `${(f.region[3] - f.region[1]) * 100}%` },
        })));
      imageUrl(`/documents/${d.id}/pages/${state.page}/image`).then((u) => {
        img.src = u;
        img.onload = () => { const a = pageEl.querySelector('.hl.active'); if (a) a.scrollIntoView({ block: 'nearest', inline: 'nearest' }); };
      }).catch(() => clear(pageEl, h('div', { class: 'empty' }, 'No fue posible renderizar la página')));
    } else {
      const pg = (d.paginas_detalle || []).find((p) => p.pagina === state.page) || { lineas: [] };
      const active = d.campos.find((f) => f.id === state.activeField);
      pageEl = h('div', { class: 'text-doc' }, pg.lineas.length ? pg.lineas.map((ln) => h('div', { class: active && active.evidencia && (active.evidencia === ln.text || active.evidencia.includes(ln.text)) ? 'active' : '' }, ln.text))
        : h('div', { class: 'muted' }, d.observacion || 'Sin texto disponible'));
      setTimeout(() => { const a = pageEl.querySelector('.active'); if (a) a.scrollIntoView({ block: 'nearest' }); }, 50);
    }
    clear(docPane, tabs, nav, pageEl);
  };

  // ------------------------------------------------------------------ panel lateral
  const drawSide = () => {
    const d = state.doc;
    const blocks = [];
    if (result) blocks.push(resultCard(), comparisonCard(), criteriaCard());
    if (d) blocks.push(docInfoCard(d), fieldsCard(d));
    if (result) blocks.push(decisionCard(), historyCard());
    if (d && d.tablas && d.tablas.length) blocks.push(tablesCard(d));
    clear(side, blocks);
  };

  const resultCard = () => h('div', { class: 'card' },
    h('div', { class: 'btn-row', style: { justifyContent: 'space-between' } }, h('h2', { style: { margin: 0 } }, `Partida ${result.id_muestra}`), badge(result.estado)),
    h('div', { class: 'small muted', style: { margin: '4px 0 8px' } },
      `Estado automático: ${result.estado_automatico}`, result.puntaje !== null ? ` · Puntaje de relación ${Math.round(result.puntaje * 100)}/100` : '',
      result.decision_revisor ? ` · ${result.decision_revisor} por ${result.revisor}` : ''),
    h('div', { class: 'small' }, `${result.tercero_esperado || ''} · NIT ${result.nit_esperado || '—'} · ${result.tipo_esperado || ''} ${result.numero_esperado || ''} · ${fmt.date(result.fecha_esperada)}`),
    result.motivos && result.motivos.length ? h('ul', { class: 'reasons small' }, result.motivos.map((m) => h('li', null, m))) : null);

  const comparisonCard = () => {
    const c = (result.explicacion || {}).comparacion_valor;
    const row = (k, v, strong) => h('tr', null, h('td', { class: 'muted' }, k), h('td', { class: 'num' }, strong ? h('b', null, v) : v));
    return h('div', { class: 'card' }, h('h3', null, 'Comparación de valor'),
      c ? h('table', { class: 'table cmp-table' }, h('tbody', null,
        row('Valor esperado', fmt.money(result.valor_esperado)),
        row('Valor extraído', fmt.money(result.valor_extraido), true),
        row('Diferencia absoluta', fmt.money(result.diferencia_absoluta)),
        row('Diferencia porcentual', fmt.pct(result.diferencia_porcentual)),
        row('Tolerancia aplicada', fmt.pct(result.tolerancia)),
        row('Resultado', c.resultado.replace(/_/g, ' ')))) : h('div', { class: 'muted small' }, 'Sin soporte con valor para comparar.'),
      c ? h('div', { class: 'small muted', style: { marginTop: '6px' } }, c.detalle) : null,
      h('div', { class: 'small muted', style: { marginTop: '6px' } }, 'Fórmula: |extraído − esperado| / |esperado| ≤ tolerancia'));
  };

  const criteriaCard = () => {
    const expl = result.criterios_por_documento && result.criterios_por_documento[state.docId];
    const fallback = (result.explicacion || {}).mejor_candidato_descartado;
    const data = (expl && expl.criterios) ? expl : fallback;
    if (!data || !data.criterios) return h('div', { class: 'card' }, h('h3', null, 'Explicación de la relación'), h('div', { class: 'small muted' }, 'Este documento no tiene puntaje por criterio (relación manual o sin candidato).'));
    return h('div', { class: 'card' },
      h('h3', null, expl ? 'Por qué se relacionó este documento' : `Mejor candidato descartado: ${fallback.archivo}`),
      h('div', { class: 'small muted', style: { marginBottom: '6px' } }, `Puntaje total ${Math.round((data.puntaje_total || 0) * 100)}/100 · promedio ponderado de los criterios evaluables`,
        data.contradiccion_nit ? h('span', { class: 'no' }, ' · NIT contradictorio') : null),
      h('table', { class: 'table cmp-table' },
        h('thead', null, h('tr', null, ['Criterio', 'Esperado', 'Extraído', 'Peso', 'Puntaje'].map((x) => h('th', null, x)))),
        h('tbody', null, data.criterios.filter((c) => c.evaluado || c.esperado).map((c) => h('tr', {
          class: c.campo_id ? 'pick' : '', title: c.campo_id ? 'Ver evidencia en el documento' : c.detalle,
          onclick: c.campo_id ? () => focusField(c.campo_id) : null,
        }, h('td', null, h('b', null, c.criterio), h('div', { class: 'small muted' }, `Prioridad ${(c.prioridad || '').toLowerCase()}`)),
        h('td', { class: 'small' }, fmt.text(c.esperado)),
        h('td', { class: 'small' }, fmt.text(c.extraido), c.pagina ? h('div', { class: 'muted' }, `pág. ${c.pagina} ⌖`) : null),
        h('td', { class: 'num small' }, c.peso),
        h('td', null, scoreTag(c.puntaje)))))),
      h('div', { class: 'small', style: { marginTop: '8px' } }, h('b', null, 'Detalle por criterio'),
        h('ul', { class: 'reasons' }, data.criterios.filter((c) => c.evaluado).map((c) => h('li', null, h('b', null, `${c.criterio}: `), c.detalle)))));
  };

  const docInfoCard = (d) => h('div', { class: 'card' },
    h('h3', null, d.archivo),
    h('div', { class: 'small' }, badge(d.estado), ' ', d.tipo_nombre || '—', d.confianza_tipo ? ` (${Math.round(d.confianza_tipo * 100)} %)` : ''),
    h('div', { class: 'small muted', style: { marginTop: '6px' } }, d.motivo_clasificacion || ''),
    h('div', { class: 'small muted' }, `Método ${d.metodo || '—'} · ${d.paginas} pág. · OCR ${d.confianza_ocr ? `${Math.round(d.confianza_ocr)} %` : 'no aplica'} · Antivirus ${d.antivirus}`),
    h('div', { class: 'mono', style: { marginTop: '6px' } }, `SHA-256 ${d.sha256}`),
    d.observacion ? h('div', { class: 'alert warn small', style: { marginTop: '8px' } }, d.observacion) : null,
    d.relaciones && d.relaciones.length ? h('div', { class: 'small', style: { marginTop: '6px' } }, 'Relacionado con: ', d.relaciones.map((r) => `${r.id_muestra} (${r.rol}${r.manual ? ', manual' : ''})`).join(', ')) : null);

  const fieldsCard = (d) => {
    const canReview = session.can('revisar');
    return h('div', { class: 'card' }, h('h3', null, 'Campos extraídos y evidencia'),
      d.campos.length ? d.campos.map((f) => h('div', { class: 'field-row', style: f.id === state.activeField ? { background: '#fff8e6' } : null },
        h('div', { class: 'top' },
          h('span', { class: 'name', onclick: () => focusField(f.id) }, label(f.campo), f.pagina ? h('span', { class: 'tag' }, `pág. ${f.pagina}`) : null,
            h('span', { class: 'tag' }, f.metodo), f.estado_revision !== 'PENDIENTE' ? h('span', { class: 'tag', style: { background: '#e1ecfb' } }, f.estado_revision) : null),
          h('span', { class: 'small muted' }, f.confianza !== null ? `conf. ${Math.round(f.confianza)}` : '')),
        h('div', { class: 'val' }, f.estado_revision === 'CORREGIDO' ? [h('s', { class: 'muted' }, f.valor), ' → ', h('b', null, f.valor_corregido)] : f.valor),
        f.evidencia ? h('div', { class: 'meta' }, `“${f.evidencia.slice(0, 160)}”`) : null,
        f.revisor ? h('div', { class: 'meta' }, `${f.estado_revision} por ${f.revisor} · ${fmt.datetime(f.fecha_revision)}${f.comentario ? ` · “${f.comentario}”` : ''}`) : null,
        canReview ? h('div', { class: 'btn-row', style: { marginTop: '4px' } },
          h('button', { class: 'btn sm', onclick: () => reviewField(f, 'ACEPTAR') }, 'Aceptar'),
          h('button', { class: 'btn sm', onclick: () => reviewField(f, 'CORREGIR') }, 'Corregir'),
          h('button', { class: 'btn sm ghost', onclick: () => reviewField(f, 'RECHAZAR') }, 'Rechazar')) : null))
        : h('div', { class: 'small muted' }, 'Sin campos extraídos.'));
  };

  const reviewField = async (f, action) => {
    const value = h('input', { class: 'input', value: f.valor_efectivo || f.valor || '' });
    const comment = h('textarea', { class: 'input', rows: 2, placeholder: action === 'ACEPTAR' ? 'Opcional' : 'Obligatorio' });
    const titles = { ACEPTAR: 'Aceptar valor extraído', CORREGIR: 'Corregir valor', RECHAZAR: 'Rechazar valor extraído' };
    const ok = await modal(`${titles[action]} — ${label(f.campo)}`, h('div', null,
      h('p', { class: 'small muted' }, `Valor extraído original (se conserva siempre): ${f.valor}`),
      action === 'CORREGIR' ? field('Valor corregido', value) : null, field('Comentario del revisor', comment)),
    [{ label: 'Guardar', value: async () => {
      try {
        await api(`/fields/${f.id}/review`, { method: 'POST', body: { action, corrected_value: action === 'CORREGIR' ? value.value : null, comment: comment.value || null } });
        return true;
      } catch (err) { toast(err.message, 'bad'); return false; }
    } }]);
    if (ok) {
      toast(action === 'ACEPTAR' ? 'Valor aceptado' : 'Registrado. El motor de vouching se recalculará en segundo plano.', 'ok');
      await loadDoc(state.docId, { fieldId: f.id });
    }
  };

  const decisionCard = () => {
    const canReview = session.can('revisar');
    if (!canReview) return null;
    const candidate = h('select', { class: 'input' }, h('option', { value: '' }, '— Seleccione un documento —'),
      (result.candidatos || []).map((c) => h('option', { value: c.id }, `${c.archivo} · ${c.tipo || ''} · ${c.estado}`)));
    const role = h('select', { class: 'input' }, h('option', { value: 'PRINCIPAL' }, 'Soporte principal'), h('option', { value: 'COMPLEMENTARIO' }, 'Soporte complementario'));
    const comment = h('textarea', { class: 'input', rows: 2, placeholder: 'Comentario del revisor' });
    const reload = () => ctx.navigate(`/p/${pid}/visor/${result.id}?t=${Date.now()}`);
    return h('div', { class: 'card' }, h('h3', null, 'Decisión del revisor'),
      h('div', { class: 'btn-row' },
        h('button', { class: 'btn ok', onclick: () => decide(result, 'APROBADO', reload) }, '✔ Aprobar'),
        h('button', { class: 'btn danger', onclick: () => decide(result, 'RECHAZADO', reload) }, '✖ Rechazar / cambiar estado')),
      h('div', { style: { marginTop: '12px' } }, field('Agregar comentario', comment),
        h('button', { class: 'btn sm', onclick: async () => {
          if (!comment.value.trim()) return;
          await api(`/results/${result.id}/comment`, { method: 'POST', body: { comment: comment.value } });
          toast('Comentario registrado', 'ok'); reload();
        } }, 'Guardar comentario')),
      h('details', { style: { marginTop: '12px' } }, h('summary', null, 'Relacionar documento manualmente'),
        h('div', { style: { marginTop: '8px' } }, field('Documento', candidate), field('Rol', role),
          h('button', { class: 'btn sm primary', onclick: async () => {
            if (!candidate.value) return;
            await api(`/results/${result.id}/links`, { method: 'POST', body: { document_id: Number(candidate.value), role: role.value, comment: comment.value || null } });
            toast('Relación registrada; el motor se recalculará', 'ok'); reload();
          } }, 'Relacionar'),
          result.documentos.filter((d) => d.manual).map((d) => h('div', { class: 'small', style: { marginTop: '6px' } }, `Manual: ${d.archivo} `,
            h('button', { class: 'btn sm ghost', onclick: async () => { await api(`/links/${d.link_id}`, { method: 'DELETE' }); toast('Relación retirada'); reload(); } }, 'Retirar'))))));
  };

  const historyCard = () => h('div', { class: 'card' }, h('h3', null, 'Historial de la partida'),
    result.historial.length ? h('ul', { class: 'timeline' }, result.historial.map((e) => h('li', null,
      h('b', null, e.accion.replace(/_/g, ' ')), ` · ${e.usuario} · ${fmt.datetime(e.fecha)}`,
      e.campo ? h('div', { class: 'small' }, `${label(e.campo)}: ${e.anterior ?? '—'} → ${e.nuevo ?? '—'}`) : (e.anterior || e.nuevo) ? h('div', { class: 'small' }, `${e.anterior ?? ''} → ${e.nuevo ?? ''}`) : null,
      e.comentario ? h('div', { class: 'small muted' }, `“${e.comentario}”`) : null)))
      : h('div', { class: 'small muted' }, 'Sin revisiones registradas.'));

  const tablesCard = (d) => h('div', { class: 'card' }, h('h3', null, 'Tablas detectadas (valores y cantidades)'),
    d.tablas.map((t) => h('div', { class: 'table-wrap', style: { marginBottom: '8px', maxHeight: '220px' } }, h('table', { class: 'table' },
      h('thead', null, h('tr', null, (t.header || []).map((c) => h('th', null, c)))),
      h('tbody', null, (t.rows || []).map((r) => h('tr', null, r.map((c) => h('td', { class: 'small' }, c)))))))));

  if (state.docId) await loadDoc(state.docId);
  else { drawDoc(); drawSide(); }

  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Visor documental'),
      h('div', { class: 'muted' }, 'Haga clic en un criterio o campo para ir a la página y resaltar el texto que lo soporta.')),
    h('a', { class: 'btn', href: `#/p/${pid}/${isResult ? 'resultados' : 'procesamiento'}` }, '← Volver')),
    h('div', { class: 'viewer' }, docPane, side));
}
