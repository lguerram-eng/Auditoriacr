// Pantalla 10: revisión y aprobación de resultados.
import { api, session } from '../api.js';
import { badge, clear, field, fmt, h, modal, sortableTable, STATUSES, toast } from '../ui.js';

export async function render(ctx) {
  const pid = ctx.project.id;
  const box = h('div');
  let filter = 'pendientes';
  const load = async () => {
    const rows = await api(`/projects/${pid}/results`);
    const needs = (r) => !['COINCIDE', 'COINCIDE CON TOLERANCIA', 'PENDIENTE'].includes(r.estado);
    const sets = {
      pendientes: rows.filter((r) => !r.decision_revisor && needs(r)),
      coincidencias: rows.filter((r) => !r.decision_revisor && !needs(r) && r.estado !== 'PENDIENTE'),
      revisadas: rows.filter((r) => r.decision_revisor),
    };
    const data = sets[filter];
    const canReview = session.can('revisar');
    clear(box,
      h('div', { class: 'chips', style: { marginBottom: '12px' } },
        [['pendientes', 'Excepciones sin decisión'], ['coincidencias', 'Coincidencias por aprobar'], ['revisadas', 'Revisadas']].map(([k, label]) =>
          h('button', { class: `chip ${filter === k ? 'active' : ''}`, onclick: () => { filter = k; load(); } }, label, h('span', { class: 'count' }, sets[k].length)))),
      filter === 'coincidencias' && canReview && data.length ? h('div', { class: 'btn-row', style: { marginBottom: '10px' } },
        h('button', { class: 'btn ok', onclick: async () => {
          const ok = await modal('Aprobar coincidencias', h('p', null, `Se aprobarán ${data.length} partidas en estado COINCIDE / COINCIDE CON TOLERANCIA. Cada aprobación queda registrada a su nombre.`), [{ label: 'Aprobar todas', class: 'ok' }]);
          if (!ok) return;
          for (const r of data) await api(`/results/${r.id}/review`, { method: 'POST', body: { decision: 'APROBADO', comment: 'Aprobación en lote de coincidencias' } });
          toast(`${data.length} partidas aprobadas`, 'ok'); load();
        } }, `✔ Aprobar ${data.length} coincidencias`)) : null,
      sortableTable([
        { key: 'id_muestra', label: 'ID' },
        { key: 'estado', label: 'Estado', render: (r) => badge(r.estado) },
        { key: 'tercero_esperado', label: 'Tercero' },
        { key: 'valor_esperado', label: 'Valor esperado', num: true, render: (r) => fmt.money(r.valor_esperado) },
        { key: 'valor_extraido', label: 'Valor extraído', num: true, render: (r) => fmt.money(r.valor_extraido) },
        { key: 'motivos', label: 'Motivos', render: (r) => h('div', { class: 'small', style: { maxWidth: '360px' } }, (r.motivos || []).join(' · ')) },
        { key: 'decision_revisor', label: 'Decisión', render: (r) => (r.decision_revisor ? h('div', { class: 'small' }, h('b', null, r.decision_revisor), h('div', { class: 'muted' }, `${r.revisor || ''} · ${fmt.datetime(r.fecha_revision)}`), r.comentario_revisor ? h('div', null, `“${r.comentario_revisor}”`) : null) : '—') },
        { key: 'acc', label: '', render: (r) => h('div', { class: 'btn-row' },
          h('a', { class: 'btn sm', href: `#/p/${pid}/visor/${r.id}`, onclick: (e) => e.stopPropagation() }, 'Visor'),
          canReview ? h('button', { class: 'btn sm ok', onclick: (e) => { e.stopPropagation(); decide(r, 'APROBADO', load); } }, 'Aprobar') : null,
          canReview ? h('button', { class: 'btn sm danger', onclick: (e) => { e.stopPropagation(); decide(r, 'RECHAZADO', load); } }, 'Rechazar') : null) },
      ], data, { empty: 'Nada pendiente en esta categoría' }));
  };
  await load();
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Revisión y aprobación'),
      h('div', { class: 'muted' }, 'Toda decisión registra usuario, fecha y hora; el estado automático se conserva.'))),
    h('div', { class: 'card' }, box));
}

export async function decide(r, decision, after) {
  const comment = h('textarea', { class: 'input', rows: 3, placeholder: decision === 'RECHAZADO' ? 'Obligatorio: explique el rechazo' : 'Opcional' });
  const status = h('select', { class: 'input' }, h('option', { value: '' }, `Mantener estado automático (${r.estado_automatico || r.estado})`), STATUSES.map((s) => h('option', { value: s }, s)));
  const res = await modal(`${decision === 'APROBADO' ? 'Aprobar' : 'Rechazar'} partida ${r.id_muestra}`,
    h('div', null, field('Estado final (opcional)', status, 'Cambiar el estado exige comentario'), field('Comentario del revisor', comment)),
    [{ label: decision === 'APROBADO' ? 'Aprobar' : 'Rechazar', class: decision === 'APROBADO' ? 'ok' : 'danger', value: async () => {
      try {
        await api(`/results/${r.id}/review`, { method: 'POST', body: { decision, status: status.value || null, comment: comment.value || null } });
        return true;
      } catch (err) { toast(err.message, 'bad'); return false; }
    } }]);
  if (res) { toast('Decisión registrada', 'ok'); if (after) after(); }
}
