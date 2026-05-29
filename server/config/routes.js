/**
 * Peta URL publik (profesional) → file HTML di folder public/.
 * Dipakai Express lokal; Vercel memakai vercel.json dengan mapping yang sama.
 */
const FILES = {
  app: 'pages/participant/taskpane.html',
  login: 'pages/participant/login.html',
  admin: 'pages/admin/dashboard.html',
  adminLogin: 'pages/admin/admin-login.html',
  adminCommands: 'pages/admin/commands.html',
};

const URLS = {
  app: '/app',
  login: '/login',
  admin: '/admin',
  adminLogin: '/admin/login',
  adminCommands: '/admin/commands',
};

/** Path lama → path baru (redirect 301) */
const LEGACY_REDIRECTS = {
  '/pages/participant/taskpane.html': URLS.app,
  '/pages/participant/login.html': URLS.login,
  '/pages/admin/dashboard.html': URLS.admin,
  '/pages/admin/admin-login.html': URLS.adminLogin,
  '/pages/admin/commands.html': URLS.adminCommands,
  '/taskpane.html': URLS.app,
  '/login.html': URLS.login,
  '/dashboard.html': URLS.admin,
  '/admin-login.html': URLS.adminLogin,
  '/commands.html': URLS.adminCommands,
};

/** Normalisasi path di manifest (path lama → path baru) */
const MANIFEST_PATH_MAP = {
  '/pages/participant/taskpane.html': URLS.app,
  '/pages/participant/login.html': URLS.login,
  '/pages/admin/dashboard.html': URLS.admin,
  '/pages/admin/admin-login.html': URLS.adminLogin,
  '/pages/admin/commands.html': URLS.adminCommands,
  '/taskpane.html': URLS.app,
  '/login.html': URLS.login,
  '/dashboard.html': URLS.admin,
  '/admin-login.html': URLS.adminLogin,
  '/commands.html': URLS.adminCommands,
};

module.exports = { FILES, URLS, LEGACY_REDIRECTS, MANIFEST_PATH_MAP };
