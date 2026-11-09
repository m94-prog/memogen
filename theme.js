/* Sets the saved light/dark theme before the page draws. A file rather than an inline script, so the page's CSP can forbid inline scripts. */
(function () { try { var t = localStorage.getItem('memogen.theme'); document.documentElement.setAttribute('data-theme', t === 'light' ? 'light' : 'dark'); } catch (e) { /* storage blocked */ } })();
