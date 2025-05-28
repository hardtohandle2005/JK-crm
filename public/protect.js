// protect.js (silent mode)
document.addEventListener('contextmenu', e => e.preventDefault());

document.addEventListener('keydown', function(e) {
  const blocked =
    e.key === 'F12' ||
    (e.ctrlKey && e.key.toLowerCase() === 'u') || // Ctrl+U
    (e.ctrlKey && e.shiftKey && ['i', 'j', 'c'].includes(e.key.toLowerCase())) || // DevTools
    (e.ctrlKey && e.key.toLowerCase() === 's'); // Ctrl+S (Save)

  if (blocked) {
    e.preventDefault();
    return false;
  }
});

(function devtoolsDetector() {
  let threshold = 160;
  setInterval(function () {
    const widthThreshold = window.outerWidth - window.innerWidth > threshold;
    const heightThreshold = window.outerHeight - window.innerHeight > threshold;

    if (widthThreshold || heightThreshold) {
      document.body.innerHTML = ""; // blank the page
    }
  }, 1000);
})();
