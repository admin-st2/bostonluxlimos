/* Sends interaction events only when an analytics collector is connected.
   Never reports a click as a submitted lead or completed booking. */
(() => {
  'use strict';
  const corporatePage = location.pathname.endsWith('/boston-corporate-car-service.html');
  const corporate = corporatePage || new URLSearchParams(location.search).get('service') === 'corporate';
  function emit(name, method) {
    if (window.bllAnalyticsAllowed !== true || typeof window.gtag !== 'function') return;
    window.gtag('event', name, {service_type: corporate ? 'corporate' : 'general', contact_method: method, transport_type: 'beacon'});
  }
  document.addEventListener('click', (event) => {
    const link = event.target.closest('a[href]');
    if (!link) return;
    const href = link.getAttribute('href');
    if (href.startsWith('tel:')) emit('call_click', 'phone');
    else if (href.startsWith('mailto:')) emit('email_click', 'email');
    else if (corporatePage && href.startsWith('contact.html') && href.includes('#quote')) emit('corporate_quote_click', 'contact_page');
    else if (location.pathname.endsWith('/contact.html') && link.hostname === 'docs.google.com' && link.pathname.includes('/forms/')) emit('quote_form_open', 'google_forms');
  });
  const embedded = location.pathname.endsWith('/contact.html') ? document.querySelector('.embedded-form') : null;
  if (embedded) embedded.addEventListener('toggle', () => {
    if (embedded.open) emit('quote_form_expand', 'embedded_google_forms');
  });
})();
