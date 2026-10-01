// Pantalla 12: usuarios y permisos (administrador).
import { api, session } from '../api.js';
import { clear, field, fmt, h, modal, toast } from '../ui.js';

const ROLE_LABEL = { administrador: 'Administrador', auditor: 'Auditor', revisor: 'Revisor', consulta: 'Consulta' };

export async function render() {
  const box = h('div');
  const [users, roles] = await Promise.all([api('/users'), api('/users/roles')]);
  const draw = (list) => clear(box, h('table', { class: 'table' },
    h('thead', null, h('tr', null, ['Nombre', 'Correo', 'Rol', 'Estado', 'Último acceso', ''].map((c) => h('th', null, c)))),
    h('tbody', null, list.map((u) => h('tr', null, h('td', null, u.full_name), h('td', null, u.email),
      h('td', null, h('select', { class: 'input', style: { width: '150px' }, disabled: u.id === session.user.id,
        onchange: async (e) => { try { await api(`/users/${u.id}`, { method: 'PATCH', body: { role: e.target.value } }); toast('Rol actualizado', 'ok'); } catch (err) { toast(err.message, 'bad'); } } },
      roles.roles.map((r) => h('option', { value: r, selected: r === u.role }, ROLE_LABEL[r] || r)))),
      h('td', null, u.is_active ? 'Activo' : h('span', { class: 'no' }, 'Inactivo')),
      h('td', null, fmt.datetime(u.last_login_at)),
      h('td', null, h('div', { class: 'btn-row' },
        u.id !== session.user.id ? h('button', { class: 'btn sm', onclick: async () => { await api(`/users/${u.id}`, { method: 'PATCH', body: { is_active: !u.is_active } }); reload(); } }, u.is_active ? 'Desactivar' : 'Activar') : null,
        h('button', { class: 'btn sm ghost', onclick: () => resetPwd(u) }, 'Restablecer contraseña'))))))));
  const reload = async () => draw(await api('/users'));
  draw(users);
  const create = async () => {
    const name = h('input', { class: 'input' }); const email = h('input', { class: 'input', type: 'email' });
    const role = h('select', { class: 'input' }, roles.roles.map((r) => h('option', { value: r, selected: r === 'consulta' }, ROLE_LABEL[r])));
    const pwd = h('input', { class: 'input', type: 'password', autocomplete: 'new-password' });
    const ok = await modal('Nuevo usuario', h('div', null, field('Nombre completo', name), field('Correo', email), field('Rol', role),
      field('Contraseña temporal', pwd, 'El usuario deberá cambiarla. Mínimo 10 caracteres con mayúscula, minúscula, número y símbolo.')),
    [{ label: 'Crear', value: async () => {
      try { await api('/users', { method: 'POST', body: { full_name: name.value, email: email.value, role: role.value, password: pwd.value } }); return true; } catch (err) { toast(err.message, 'bad'); return false; }
    } }]);
    if (ok) { toast('Usuario creado', 'ok'); reload(); }
  };
  const resetPwd = async (u) => {
    const pwd = h('input', { class: 'input', type: 'password', autocomplete: 'new-password' });
    const ok = await modal(`Restablecer contraseña de ${u.email}`, field('Contraseña temporal', pwd), [{ label: 'Guardar', value: async () => {
      try { await api(`/users/${u.id}`, { method: 'PATCH', body: { password: pwd.value } }); return true; } catch (err) { toast(err.message, 'bad'); return false; }
    } }]);
    if (ok) toast('Contraseña restablecida', 'ok');
  };
  const perms = ['ver', 'exportar', 'revisar', 'proyecto.crear', 'proyecto.editar', 'proyecto.eliminar', 'importar', 'cargar', 'procesar', 'configurar', 'usuarios', 'auditoria.proyecto', 'auditoria.global'];
  return h('div', null,
    h('div', { class: 'page-head' }, h('div', { class: 'grow' }, h('h1', null, 'Usuarios y permisos')), h('button', { class: 'btn primary', onclick: create }, '＋ Nuevo usuario')),
    h('div', { class: 'card' }, box),
    h('div', { class: 'card' }, h('h2', null, 'Matriz de permisos por rol'),
      h('table', { class: 'table' }, h('thead', null, h('tr', null, h('th', null, 'Permiso'), roles.roles.map((r) => h('th', null, ROLE_LABEL[r])))),
        h('tbody', null, perms.map((p) => h('tr', null, h('td', null, p), roles.roles.map((r) => h('td', null, roles.permisos[r].includes(p) ? h('span', { class: 'yes' }, '✓') : h('span', { class: 'muted' }, '—')))))))));
}
