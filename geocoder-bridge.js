// Ponte de geocoding — contexto da página, acessa google.maps
window.__gfxBridgeReady = true;

window.addEventListener('__gfxGeocodeRequest', function(e) {
  const { callbackId, address } = e.detail;
  try {
    const geocoder = new google.maps.Geocoder();
    geocoder.geocode({ address: address, region: 'br' }, function(results, status) {
      if (status !== 'OK' || !results || !results.length) {
        window.dispatchEvent(new CustomEvent(callbackId, { detail: { error: 'Geocoding: ' + status } }));
        return;
      }
      const loc = results[0].geometry.location;
      window.dispatchEvent(new CustomEvent(callbackId, { detail: {
        lat: typeof loc.lat === 'function' ? loc.lat() : loc.lat,
        lng: typeof loc.lng === 'function' ? loc.lng() : loc.lng
      }}));
    });
  } catch(err) {
    window.dispatchEvent(new CustomEvent(callbackId, { detail: { error: err.message } }));
  }
});
console.log('[GeoFixer] geocoder-bridge pronto ✓');
