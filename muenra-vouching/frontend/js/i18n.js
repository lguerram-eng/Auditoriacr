// Textos de la interfaz. Español por defecto; inglés preparado como opción futura.
const DICT = {
  es: {
    app: 'Muenra Vouching', tagline: 'Vouching de auditoría',
    nav_general: 'General', nav_project: 'Proyecto', nav_admin: 'Administración',
    dashboard: 'Tablero principal', projects: 'Proyectos', new_project: 'Crear proyecto',
    import: 'Importar Excel', upload: 'Cargar documentos', processing: 'Procesamiento',
    results: 'Resultados', viewer: 'Visor documental', exceptions: 'Excepciones', review: 'Revisión y aprobación',
    settings: 'Configuración', users: 'Usuarios y permisos', audit: 'Historial de auditoría', exports: 'Exportaciones',
    logout: 'Cerrar sesión', select_project: 'Seleccione un proyecto',
  },
  en: {
    app: 'Muenra Vouching', tagline: 'Audit vouching',
    nav_general: 'General', nav_project: 'Project', nav_admin: 'Administration',
    dashboard: 'Dashboard', projects: 'Projects', new_project: 'New project',
    import: 'Import Excel', upload: 'Upload documents', processing: 'Processing',
    results: 'Results', viewer: 'Document viewer', exceptions: 'Exceptions', review: 'Review & approval',
    settings: 'Settings', users: 'Users & permissions', audit: 'Audit trail', exports: 'Exports',
    logout: 'Sign out', select_project: 'Select a project',
  },
};

let lang = 'es';
try { lang = localStorage.getItem('mv_lang') || 'es'; } catch { /* sin almacenamiento */ }

export function t(key) { return (DICT[lang] && DICT[lang][key]) || DICT.es[key] || key; }
export function setLang(l) { lang = DICT[l] ? l : 'es'; try { localStorage.setItem('mv_lang', lang); } catch { /* */ } }
export function getLang() { return lang; }
