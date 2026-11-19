/* Shared source-origin manifest. A failed load is never an empty policy. */
(() => {
  const approvedZrokOrigin = 'https://bentang10ops20261119.shares.zrok.io';
  const publicRoutes = ['/api/recordings/publication', '/api/schedule/calendar'];
  let acceptedApiOrigin = null;
  const checkedOrigin = (value, allowLoopback = false) => {
    if (typeof value !== 'string' || /\s/.test(value)) throw Error('Invalid Bentan origin');
    const parsed = new URL(value);
    const protocolAllowed = parsed.protocol === 'https:' || (allowLoopback && parsed.protocol === 'http:' && ['127.0.0.1', 'localhost'].includes(parsed.hostname));
    if (!protocolAllowed || parsed.username || parsed.password || parsed.pathname !== '/' || parsed.search || parsed.hash)
      throw Error('Invalid Bentan origin');
    return parsed.origin;
  };
  const source = new URL('bentan-endpoints.json', document.currentScript.src);
  const ready = fetch(source, {cache: 'no-store'}).then(response => {
    if (!response.ok) throw Error('Bentan endpoint manifest unavailable');
    return response.json();
  }).then(manifest => {
    if (manifest.version !== 1) throw Error('Unsupported Bentan endpoint manifest');
    const config = Object.freeze({api: checkedOrigin(manifest.apiBaseUrl), staff: checkedOrigin(manifest.staffBaseUrl, true)});
    acceptedApiOrigin = config.api;
    return config;
  });
  // Existing page request helpers handle rejection as an unavailable policy.
  ready.catch(() => {});
  window.bentanEndpointOrigins = () => ready;
  window.bentanApiUrl = route => {
    if (!publicRoutes.includes(route)) return Promise.reject(Error('Unknown public integration route'));
    return ready.then(config => config.api + route);
  };
  window.bentanPublication = async loader => {
    try { const payload = await loader([await window.bentanApiUrl('/api/recordings/publication')]); return Array.isArray(payload?.rules) ? payload : null; }
    catch { return null; }
  };
  window.bentanApiFetchOptions = (value, options = {}) => {
    const url = new URL(value, source);
    if (url.origin !== approvedZrokOrigin || url.origin !== acceptedApiOrigin || !publicRoutes.includes(url.pathname)) return options;
    return {...options, credentials: 'omit', headers: {...options.headers, skip_zrok_interstitial: 'true'}};
  };
  window.bentanStaffUrl = route => {
    if (!route.startsWith('/staff/recordings?') || /[\r\n]/.test(route)) return Promise.reject(Error('Unknown staff destination'));
    return ready.then(config => config.staff + route);
  };
})();
