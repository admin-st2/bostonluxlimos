/* Native details remain usable without JavaScript; only phone presentation collapses. */
(() => {
  const groups = [...document.querySelectorAll('footer .footer-group')];
  if (!groups.length) return;
  const phone = window.matchMedia('(max-width: 600px)');
  const applyLayout = () => {
    groups.forEach(group => {
      group.open = !phone.matches;
      const summary = group.querySelector('summary');
      if (phone.matches) summary.removeAttribute('tabindex');
      else summary.setAttribute('tabindex', '-1');
    });
  };
  groups.forEach(group => {
    group.querySelector('summary').addEventListener('click', event => {
      if (!phone.matches) event.preventDefault();
    });
    group.addEventListener('toggle', () => {
      if (phone.matches && group.open) {
        groups.forEach(other => { if (other !== group) other.open = false; });
      }
    });
  });
  applyLayout();
  phone.addEventListener('change', applyLayout);
})();
