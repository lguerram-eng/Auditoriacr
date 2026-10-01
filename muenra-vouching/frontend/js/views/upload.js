// Pantalla 5: carga masiva de documentos por arrastrar y soltar.
import { api } from '../api.js';
import { clear, fmt, h, toast } from '../ui.js';

const ACCEPT = '.pdf,.xml,.png,.jpg,.jpeg,.tif,.tiff,.docx';
const BATCH = 20;

export async function render(ctx) {
  const pid = ctx.project.id;
  let queue = [];
  const list = h('div');
  const result = h('div');
  const progress = h('div', { style: { width: '0%' } });
  const progressWrap = h('div', { class: 'progress hidden', style: { marginTop: '12px' } }, progress);
  const files = h('input', { type: 'file', multiple: true, accept: ACCEPT, class: 'hidden', onchange: (e) => add(e.target.files) });
  const folder = h('input', { type: 'file', multiple: true, webkitdirectory: true, class: 'hidden', onchange: (e) => add(e.target.files) });
  const add = (fl) => {
    const ok = [...fl].filter((f) => ACCEPT.split(',').some((ext) => f.name.toLowerCase().endsWith(ext)));
    const skipped = fl.length - ok.length;
    queue = queue.concat(ok);
    if (skipped) toast(`${skipped} archivos omitidos por tipo no permitido`, 'bad');
    draw();
  };
  const draw = () => {
    const total = queue.reduce((a, f) => a + f.size, 0);
    clear(list, queue.length ? h('div', null,
      h('div', { class: 'small muted', style: { margin: '10px 0' } }, `${queue.length} archivos · ${fmt.int(Math.round(total / 1024))} KB`),
      h('div', { class: 'table-wrap', style: { maxHeight: '260px' } }, h('table', { class: 'table' },
        h('tbody', null, queue.map((f, i) => h('tr', null, h('td', null, f.name), h('td', { class: 'num' }, `${fmt.int(Math.round(f.size / 1024))} KB`),
          h('td', null, h('button', { class: 'btn sm ghost', onclick: () => { queue.splice(i, 1); draw(); } }, 'Quitar'))))))))
      : null);
  };
  const drop = h('div', { class: 'dropzone', onclick: () => files.click(),
    ondragover: (e) => { e.preventDefault(); drop.classList.add('over'); }, ondragleave: () => drop.classList.remove('over'),
    ondrop: (e) => { e.preventDefault(); drop.classList.remove('over'); add(e.dataTransfer.files); } },
  h('div', { class: 'big' }, '📂'), h('div', null, h('b', null, 'Arrastre aquí los documentos soporte'), ' o haga clic para seleccionarlos'),
  h('div', { class: 'small muted' }, 'PDF (texto o escaneado), XML de facturación electrónica, PNG, JPG/JPEG, TIFF y DOCX'), files, folder);

  const upload = async () => {
    if (!queue.length) { toast('No hay archivos para cargar', 'bad'); return; }
    const accepted = []; const rejected = [];
    progressWrap.classList.remove('hidden');
    for (let i = 0; i < queue.length; i += BATCH) {
      const form = new FormData();
      queue.slice(i, i + BATCH).forEach((f) => form.append('files', f, f.name));
      try {
        const r = await api(`/projects/${pid}/documents`, { method: 'POST', form });
        accepted.push(...r.aceptados); rejected.push(...r.rechazados);
      } catch (err) {
        queue.slice(i, i + BATCH).forEach((f) => rejected.push({ archivo: f.name, motivo: err.message }));
      }
      progress.style.width = `${Math.min(100, ((i + BATCH) / queue.length) * 100)}%`;
    }
    queue = [];
    draw();
    clear(result,
      h('div', { class: `alert ${rejected.length ? 'warn' : 'ok'}` }, `${accepted.length} documentos cargados y encolados para procesamiento; ${rejected.length} rechazados.`),
      accepted.some((a) => a.duplicado_de) ? h('div', { class: 'alert warn' }, 'Algunos archivos son idénticos (mismo SHA-256) a documentos ya cargados: se marcarán como POSIBLE DUPLICADO.') : null,
      rejected.length ? h('table', { class: 'table' }, h('thead', null, h('tr', null, h('th', null, 'Archivo rechazado'), h('th', null, 'Motivo'))),
        h('tbody', null, rejected.map((r) => h('tr', null, h('td', null, r.archivo), h('td', null, r.motivo))))) : null,
      accepted.length ? h('table', { class: 'table', style: { marginTop: '10px' } },
        h('thead', null, h('tr', null, ['Archivo', 'SHA-256', 'Tipo', 'Antivirus', 'Duplicado de'].map((c) => h('th', null, c)))),
        h('tbody', null, accepted.map((a) => h('tr', null, h('td', null, a.archivo), h('td', { class: 'mono' }, a.sha256), h('td', null, a.tipo),
          h('td', null, a.antivirus), h('td', null, a.duplicado_de ? `#${a.duplicado_de}` : '—'))))) : null,
      h('div', { class: 'btn-row', style: { marginTop: '12px' } }, h('a', { class: 'btn primary', href: `#/p/${pid}/procesamiento` }, 'Ver procesamiento →')));
  };

  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Cargar documentos'),
      h('div', { class: 'muted' }, 'Los archivos se validan (tipo real, tamaño, contenido activo y antivirus), se cifran en reposo y se calcula su SHA-256.'))),
    h('div', { class: 'card' }, drop,
      h('div', { class: 'btn-row', style: { marginTop: '12px' } },
        h('button', { class: 'btn', onclick: () => folder.click() }, 'Seleccionar carpeta'),
        h('button', { class: 'btn primary', onclick: upload }, 'Cargar y procesar')),
      progressWrap, list),
    h('div', { class: 'card' }, result));
}
