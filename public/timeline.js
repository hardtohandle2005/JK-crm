document.addEventListener('DOMContentLoaded', () => {
    // Get client name from query string
const urlParams = new URLSearchParams(window.location.search);
const clientName = urlParams.get('client');

if (clientName) {
  document.getElementById('clientName').innerText = `Client: ${clientName}`;
} else {
  alert("No client specified in URL. Redirecting back.");
  window.location.href = "/index.html";
}


    const steps = document.querySelectorAll('.step');
    const timelineContainer = document.getElementById('timeline-container');
  
    steps.forEach(step => {
      step.addEventListener('click', () => {
        const section = step.getAttribute('data-section');
  
        // Reset active class
        steps.forEach(s => s.classList.remove('active'));
        step.classList.add('active');
  
        // Load section content
        loadSection(section);
      });
    });
  
    function loadSection(section) {
      timelineContainer.innerHTML = ''; // clear previous
  
      const sectionDiv = document.createElement('div');
      sectionDiv.classList.add('accordion-section', 'active');
      sectionDiv.innerHTML = `<h2>${capitalize(section)}</h2><div id="${section}-timeline">Loading...</div>`;
  
      timelineContainer.appendChild(sectionDiv);
  
      // In next step, we'll load dynamic content from backend for each section
    }
  
    function capitalize(word) {
      return word.charAt(0).toUpperCase() + word.slice(1);
    }
  });
  