// Applies the saved theme before the app bundle loads (prevents a white flash in dark mode).
try {
  var s = JSON.parse(localStorage.getItem('affice.boot') || '{}');
  var dark = s.theme === 'dark' || (s.theme !== 'light' && window.matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
} catch (e) {
  /* ignore */
}
