// Pantalla 1: inicio de sesión, recuperación y restablecimiento de contraseña.
import { api, session } from '../api.js';
import { clear, field, h, toast } from '../ui.js';

export function render(ctx) {
  const box = h('div', { class: 'card' });
  const hero = h('div', { class: 'login-hero' },
    h('div', { class: 'brand' }, h('div', { class: 'brand-mark' }, 'M'), h('div', { class: 'brand-name' }, 'Muenra', h('small', null, 'Vouching'))),
    h('div', null,
      h('h1', null, 'Evidencia trazable para cada partida auditada'),
      h('p', null, 'Muenra Vouching relaciona la población contable con sus documentos soporte, explica cada coincidencia criterio por criterio y conserva la evidencia por campo y página.'),
      h('ul', null,
        h('li', null, 'XML DIAN, PDF con texto, PDF escaneado e imágenes con OCR'),
        h('li', null, 'Tolerancia de valor configurable (10 % por defecto)'),
        h('li', null, 'Revisión humana con historial auditable'),
        h('li', null, 'Exportación completa a Excel'))),
    h('div', { class: 'small', style: { color: '#9fb2cc' } }, 'Información confidencial de auditoría. Acceso restringido a usuarios autorizados.'));
  const wrap = h('div', { class: 'login-wrap' }, hero, h('div', { class: 'login-form' }, box));
  if (ctx.mode === 'reset') resetForm(box, ctx); else loginForm(box, ctx);
  return wrap;
}

function loginForm(box, ctx) {
  const email = h('input', { class: 'input', type: 'email', autocomplete: 'username', required: true, placeholder: 'usuario@empresa.com' });
  const pass = h('input', { class: 'input', type: 'password', autocomplete: 'current-password', required: true });
  const msg = h('div');
  const submit = async (e) => {
    e.preventDefault();
    clear(msg);
    try {
      const data = await api('/auth/login', { method: 'POST', body: { email: email.value, password: pass.value } });
      session.save(data.access_token, { ...data.user, permisos: data.permisos });
      if (data.user.must_change_password) toast('Por seguridad, cambie su contraseña temporal en "Configuración → Mi cuenta".', 'bad');
      ctx.navigate('/');
    } catch (err) {
      clear(msg, h('div', { class: 'alert bad' }, err.message));
    }
  };
  clear(box,
    h('h2', null, 'Iniciar sesión en Muenra Vouching'),
    h('p', { class: 'muted small' }, 'Ingrese con su correo corporativo.'),
    msg,
    h('form', { onsubmit: submit }, field('Correo electrónico', email), field('Contraseña', pass),
      h('button', { class: 'btn primary', type: 'submit', style: { width: '100%', justifyContent: 'center' } }, 'Ingresar')),
    h('p', { class: 'small', style: { marginTop: '14px' } }, h('a', { href: '#', onclick: (e) => { e.preventDefault(); forgotForm(box, ctx); } }, '¿Olvidó su contraseña?')));
}

function forgotForm(box, ctx) {
  const email = h('input', { class: 'input', type: 'email', required: true });
  const msg = h('div');
  clear(box, h('h2', null, 'Recuperar contraseña'), msg,
    h('form', { onsubmit: async (e) => {
      e.preventDefault();
      const r = await api('/auth/password/forgot', { method: 'POST', body: { email: email.value } });
      clear(msg, h('div', { class: 'alert ok' }, r.mensaje),
        r.token_desarrollo ? h('div', { class: 'alert warn' }, 'Modo desarrollo (sin SMTP): ',
          h('a', { href: `#/restablecer?token=${encodeURIComponent(r.token_desarrollo)}` }, 'abrir enlace de restablecimiento')) : null);
    } }, field('Correo electrónico', email), h('button', { class: 'btn primary', type: 'submit' }, 'Enviar instrucciones')),
    h('p', { class: 'small' }, h('a', { href: '#', onclick: (e) => { e.preventDefault(); loginForm(box, ctx); } }, '← Volver')));
}

function resetForm(box, ctx) {
  const token = ctx.query.get('token') || '';
  const p1 = h('input', { class: 'input', type: 'password', autocomplete: 'new-password', required: true });
  const p2 = h('input', { class: 'input', type: 'password', autocomplete: 'new-password', required: true });
  const msg = h('div');
  clear(box, h('h2', null, 'Definir nueva contraseña'), msg,
    h('form', { onsubmit: async (e) => {
      e.preventDefault();
      if (p1.value !== p2.value) { clear(msg, h('div', { class: 'alert bad' }, 'Las contraseñas no coinciden')); return; }
      try {
        await api('/auth/password/reset', { method: 'POST', body: { token, new_password: p1.value } });
        toast('Contraseña actualizada. Inicie sesión.', 'ok');
        ctx.navigate('/login');
      } catch (err) { clear(msg, h('div', { class: 'alert bad' }, err.message)); }
    } },
    field('Nueva contraseña', p1, 'Mínimo 10 caracteres con mayúscula, minúscula, número y símbolo.'),
    field('Confirmar contraseña', p2),
    h('button', { class: 'btn primary', type: 'submit' }, 'Guardar')));
}
