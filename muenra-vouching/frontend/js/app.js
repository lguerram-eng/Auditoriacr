// Muenra Vouching — aplicación de una sola página (enrutador por hash, sin dependencias externas).
import { api, session, setUnauthorizedHandler } from './api.js';
import { t } from './i18n.js';
import { clear, errorBox, h, loading } from './ui.js';

import * as auditView from './views/audit.js';
import * as dashboardView from './views/dashboard.js';
import * as exceptionsView from './views/exceptions.js';
import * as exportsView from './views/exports.js';
import * as importView from './views/import.js';
import * as loginView from './views/login.js';
import * as processingView from './views/processing.js';
import * as projectsView from './views/projects.js';
import * as resultsView from './views/results.js';
import * as reviewView from './views/review.js';
import * as settingsView from './views/settings.js';
import * as uploadView from './views/upload.js';
import * as usersView from './views/users.js';
import * as viewerView from './views/viewer.js';

const root = document.getElementById('app');
let projects = [];
let leaveHandlers = [];

const ROUTES = [
  { re: /^\/login$/, view: loginView, public: true },
  { re: /^\/restablecer$/, view: loginView, public: true, mode: 'reset' },
  { re: /^\/$/, view: dashboardView, title: 'dashboard', nav: 'dashboard' },
  { re: /^\/proyectos$/, view: projectsView, title: 'projects', nav: 'projects' },
  { re: /^\/proyectos\/nuevo$/, view: projectsView, title: 'new_project', nav: 'new_project', mode: 'new' },
  { re: /^\/usuarios$/, view: usersView, title: 'users', nav: 'users', perm: 'usuarios' },
  { re: /^\/auditoria$/, view: auditView, title: 'audit', nav: 'audit_global', perm: 'auditoria.global', mode: 'global' },
  { re: /^\/p\/(\d+)\/tablero$/, view: dashboardView, title: 'dashboard', nav: 'p_dashboard', project: true },
  { re: /^\/p\/(\d+)\/importar$/, view: importView, title: 'import', nav: 'import', project: true },
  { re: /^\/p\/(\d+)\/cargar$/, view: uploadView, title: 'upload', nav: 'upload', project: true },
  { re: /^\/p\/(\d+)\/procesamiento$/, view: processingView, title: 'processing', nav: 'processing', project: true },
  { re: /^\/p\/(\d+)\/resultados$/, view: resultsView, title: 'results', nav: 'results', project: true },
  { re: /^\/p\/(\d+)\/visor\/(\d+)$/, view: viewerView, title: 'viewer', nav: 'results', project: true, mode: 'result' },
  { re: /^\/p\/(\d+)\/documento\/(\d+)$/, view: viewerView, title: 'viewer', nav: 'processing', project: true, mode: 'document' },
  { re: /^\/p\/(\d+)\/excepciones$/, view: exceptionsView, title: 'exceptions', nav: 'exceptions', project: true },
  { re: /^\/p\/(\d+)\/revision$/, view: reviewView, title: 'review', nav: 'review', project: true },
  { re: /^\/p\/(\d+)\/configuracion$/, view: settingsView, title: 'settings', nav: 'settings', project: true },
  { re: /^\/p\/(\d+)\/historial$/, view: auditView, title: 'audit', nav: 'audit', project: true, mode: 'project' },
  { re: /^\/p\/(\d+)\/exportaciones$/, view: exportsView, title: 'exports', nav: 'exports', project: true },
];

export function navigate(path) { window.location.hash = `#${path}`; }

setUnauthorizedHandler(() => navigate('/login'));

function parse() {
  const raw = window.location.hash.replace(/^#/, '') || '/';
  const [path, qs] = raw.split('?');
  return { path, query: new URLSearchParams(qs || '') };
}

async function loadProjects() {
  try { projects = await api('/projects'); } catch { projects = []; }
  return projects;
}

function currentProjectId() {
  const m = parse().path.match(/^\/p\/(\d+)/);
  if (m) return Number(m[1]);
  try { return Number(sessionStorage.getItem('mv_project')) || null; } catch { return null; }
}

function sidebar(route, pid) {
  const user = session.user;
  const link = (path, key, ico, nav, perm) => (perm && !session.can(perm) ? null
    : h('a', { href: `#${path}`, class: route.nav === nav ? 'active' : '' }, h('span', { class: 'ico' }, ico), t(key)));
  const proj = pid ? `/p/${pid}` : null;
  const picker = h('select', {
    'aria-label': t('select_project'),
    onchange: (e) => { if (e.target.value) navigate(`/p/${e.target.value}/tablero`); },
  }, h('option', { value: '' }, `— ${t('select_project')} —`),
  projects.map((p) => h('option', { value: p.id, selected: p.id === pid }, `${p.code} · ${p.client_name}`)));
  return h('aside', { class: 'sidebar' },
    h('div', { class: 'brand' }, h('div', { class: 'brand-mark' }, 'M'), h('div', { class: 'brand-name' }, 'Muenra', h('small', null, 'Vouching'))),
    h('nav', { class: 'nav nav-section' },
      h('div', { class: 'nav-title' }, t('nav_general')),
      link('/', 'dashboard', '◧', 'dashboard'),
      link('/proyectos', 'projects', '▤', 'projects'),
      link('/proyectos/nuevo', 'new_project', '＋', 'new_project', 'proyecto.crear')),
    h('div', { class: 'project-picker' }, picker),
    proj ? h('nav', { class: 'nav nav-section' },
      h('div', { class: 'nav-title' }, t('nav_project')),
      link(`${proj}/tablero`, 'dashboard', '◔', 'p_dashboard'),
      link(`${proj}/importar`, 'import', '⇪', 'import', 'importar'),
      link(`${proj}/cargar`, 'upload', '⬆', 'upload', 'cargar'),
      link(`${proj}/procesamiento`, 'processing', '⚙', 'processing'),
      link(`${proj}/resultados`, 'results', '☰', 'results'),
      link(`${proj}/excepciones`, 'exceptions', '⚠', 'exceptions'),
      link(`${proj}/revision`, 'review', '✔', 'review'),
      link(`${proj}/exportaciones`, 'exports', '⤓', 'exports'),
      link(`${proj}/configuracion`, 'settings', '⚑', 'settings'),
      link(`${proj}/historial`, 'audit', '⏱', 'audit', 'auditoria.proyecto')) : null,
    (session.can('usuarios') || session.can('auditoria.global')) ? h('nav', { class: 'nav nav-section' },
      h('div', { class: 'nav-title' }, t('nav_admin')),
      link('/usuarios', 'users', '👥', 'users', 'usuarios'),
      link('/auditoria', 'audit', '⏱', 'audit_global', 'auditoria.global')) : null,
    h('div', { class: 'sidebar-foot' }, 'Muenra Vouching v1.0', h('br'), user ? `${user.role}` : ''));
}

function topbar(route, project) {
  const user = session.user || {};
  const initials = (user.full_name || user.email || '?').split(' ').map((s) => s[0]).slice(0, 2).join('').toUpperCase();
  return h('header', { class: 'topbar' },
    h('button', { class: 'btn sm menu-toggle', 'aria-label': 'Menú', onclick: () => document.querySelector('.sidebar').classList.toggle('open') }, '☰ Menú'),
    h('div', { class: 'crumbs' }, 'Muenra Vouching', project ? [' / ', h('b', null, `${project.code} — ${project.name}`)] : null,
      route.title ? [' / ', t(route.title)] : null),
    h('div', { class: 'user-chip' }, h('div', { class: 'avatar' }, initials), h('div', null, user.full_name, h('div', { class: 'small muted' }, user.email))),
    h('button', { class: 'btn sm', onclick: logout }, t('logout')));
}

async function logout() {
  try { await api('/auth/logout', { method: 'POST' }); } catch { /* */ }
  session.clear();
  navigate('/login');
}

async function render() {
  leaveHandlers.forEach((fn) => { try { fn(); } catch { /* */ } });
  leaveHandlers = [];
  const { path, query } = parse();
  const route = ROUTES.find((r) => r.re.test(path));
  if (!route) { navigate(session.token ? '/' : '/login'); return; }
  if (!route.public && !session.token) { navigate('/login'); return; }
  const params = path.match(route.re).slice(1).map(Number);
  const ctx = {
    params, query, mode: route.mode, navigate,
    onLeave: (fn) => leaveHandlers.push(fn),
    refreshProjects: loadProjects,
    get projects() { return projects; },
  };
  if (route.public) { clear(root, await route.view.render(ctx)); return; }
  if (route.perm && !session.can(route.perm)) { navigate('/'); return; }
  if (!projects.length) await loadProjects();
  const pid = route.project ? params[0] : currentProjectId();
  let project = pid ? projects.find((p) => p.id === pid) : null;
  if (route.project && !project) {
    await loadProjects();
    project = projects.find((p) => p.id === pid);
    if (!project) { navigate('/proyectos'); return; }
  }
  if (project) { try { sessionStorage.setItem('mv_project', String(project.id)); } catch { /* */ } }
  ctx.project = project;
  const content = h('section', { class: 'content' }, loading());
  clear(root, h('div', { class: 'layout' }, sidebar(route, project ? project.id : null), h('main', { class: 'main' }, topbar(route, project), content)));
  try {
    const node = await route.view.render(ctx);
    clear(content, node);
  } catch (err) {
    if (err.status !== 401) clear(content, errorBox(err));
  }
  window.scrollTo(0, 0);
}

window.addEventListener('hashchange', render);
render();
