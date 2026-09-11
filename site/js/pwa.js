// Registers the service worker (offline app shell). The version query makes each release a fresh worker.
import { CONFIG } from './config.js';

if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register(`sw.js?v=${encodeURIComponent(CONFIG.appVersion)}`).catch(err => console.warn('Service worker not registered:', err));
  });
}
