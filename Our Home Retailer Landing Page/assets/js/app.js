(function(){
  const cardsEl   = document.getElementById('cards');
  const searchEl  = document.getElementById('search');
  const sortEl    = document.getElementById('sortSelect');
  const typeEl    = document.getElementById('typeSelect');

  // No dates anymore—just render from the static REPORTS list
  render();

  function render(){
    const q    = (searchEl.value || '').trim().toLowerCase();
    const type = typeEl.value;
    const sort = sortEl.value;

    let items = (window.REPORTS || []).slice();

    // Filter by search
    if(q){
      items = items.filter(x => x.name.toLowerCase().includes(q));
    }
    // Filter by type (Power BI / Excel) but do not show it on cards
    if(type !== 'all'){
      items = items.filter(x => x.type === type);
    }

    // Sort by name only
    if(sort === 'name'){
      items.sort((a,b) => a.name.localeCompare(b.name));
    }

    cardsEl.innerHTML = items.map(cardHtml).join('');
  }

  function cardHtml(x){
    return `
      <div class="col-12 col-sm-6 col-lg-4">
        <div class="card report h-100">
          <div class="logo-wrap">
            <img src="${x.logo}" alt="${x.name} logo" onerror="this.src='assets/logos/logo.png'">
          </div>
          <div class="card-body d-flex flex-column gap-2">
            <h2 class="report-title h6 mb-0">${x.name}</h2>
            <a class="btn btn-primary w-100 mt-auto" href="${x.url}" target="_blank" rel="noopener">
              Open Dashboard <i class="bi bi-box-arrow-up-right ms-1"></i>
            </a>
          </div>
        </div>
      </div>`;
  }

  // Events
  searchEl.addEventListener('input', render);
  sortEl.addEventListener('change', render);
  typeEl.addEventListener('change', render);
})();
