/* ═══════════════════════════════════════════════════════
   PROMART STORE PLANNING — script.js
   Sections: Excel upload · Filters · Chart.js · Legend · Template
═══════════════════════════════════════════════════════ */

/* ── STATE ────────────────────────────────────────────── */
const PALETTE = [
  '#F15A22','#38bdf8','#fbbf24','#10b981',
  '#f43f5e','#a855f7','#06b6d4','#84cc16',
  '#fb923c','#e879f9','#f97316','#14b8a6'
];

let rawData   = [];
let chartInst = null;
let curPais   = '';

/* ── DOM REFS ─────────────────────────────────────────── */
const upzone      = document.getElementById('upzone');
const excelInput  = document.getElementById('excelInput');
const dynCtrl     = document.getElementById('dyn-ctrl');
const paisFilter  = document.getElementById('paisFilter');
const areaFilter  = document.getElementById('areaFilter');
const metFilter   = document.getElementById('metricaFilter');
const typeFilter  = document.getElementById('chartTypeFilter');
const checklist   = document.getElementById('tiendaChecklist');
const canvas      = document.getElementById('mainChart');
const legendEl    = document.getElementById('legend');
const emptyEl     = document.getElementById('emptyMsg');
const ctitleEl    = document.getElementById('ctitle');
const csubEl      = document.getElementById('csub');

/* ── DRAG & DROP ──────────────────────────────────────── */
upzone.addEventListener('dragover',  e => { e.preventDefault(); upzone.classList.add('drag'); });
upzone.addEventListener('dragleave', ()  => upzone.classList.remove('drag'));
upzone.addEventListener('drop',      e => {
  e.preventDefault(); upzone.classList.remove('drag');
  const f = e.dataTransfer.files[0];
  if (f) readFile(f);
});

/* ── FILE INPUT ───────────────────────────────────────── */
excelInput.addEventListener('change', e => {
  if (e.target.files[0]) readFile(e.target.files[0]);
});

function readFile(file) {
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const wb   = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: 0 });
      if (!rows.length) { alert('El archivo está vacío.'); return; }
      rawData = rows;

      /* update upzone feedback */
      upzone.innerHTML = `
        <div class="upzone-ico">✅</div>
        <p style="color:rgba(255,255,255,.7);">${rows.length} registros cargados</p>
        <p style="font-size:11px;color:rgba(255,255,255,.35);">${file.name}</p>`;

      dynCtrl.style.display = 'block';
      initPaises();
    } catch (err) {
      alert('Error al leer el archivo. Asegúrate de que sea .xlsx o .xls válido.');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ── HELPER: get column value safely ─────────────────── */
function col(row, ...keys) {
  for (const k of keys) {
    const v = row[k] ?? row[k.toUpperCase()] ?? row[k.toLowerCase()];
    if (v !== undefined && v !== null) return v;
  }
  return null;
}

/* ── INIT PAÍSES ──────────────────────────────────────── */
function initPaises() {
  const paises = [...new Set(rawData.map(r => col(r,'Pais','pais','PAIS') || '—'))];
  paisFilter.innerHTML = paises.map(p => `<option value="${p}">${p}</option>`).join('');
  curPais = paises[0];
  updateStores();
}

/* ── UPDATE TIENDAS & AREAS ───────────────────────────── */
function updateStores() {
  curPais = paisFilter.value;

  const filtered = rawData.filter(r => (col(r,'Pais','pais','PAIS') || '—') === curPais);
  const tiendas  = [...new Set(filtered.map(r => col(r,'Tienda','tienda','TIENDA') || '—'))];
  const areas    = [...new Set(filtered.map(r => col(r,'Area_Comercial','Area','area') || '—'))];

  /* checklist */
  checklist.innerHTML = tiendas.map((t, i) => `
    <label>
      <input type="checkbox" class="store-cb" value="${t}" onchange="renderChart()" checked>
      <span style="display:inline-flex;align-items:center;gap:6px;">
        <span style="width:9px;height:9px;border-radius:50%;background:${PALETTE[i % PALETTE.length]};flex-shrink:0;"></span>
        ${t}
      </span>
    </label>`).join('');

  /* area filter */
  areaFilter.innerHTML = '<option value="__all__">— Todas las áreas —</option>' +
    areas.map(a => `<option value="${a}">${a}</option>`).join('');

  renderChart();
}

function selectAll(v) {
  document.querySelectorAll('.store-cb').forEach(cb => cb.checked = v);
  renderChart();
}

/* ── RENDER CHART ─────────────────────────────────────── */
function renderChart() {
  const selected  = Array.from(document.querySelectorAll('.store-cb:checked')).map(cb => cb.value);
  const metrica   = metFilter.value;          /* m2 | m3 | N_RACKS */
  const areaVal   = areaFilter.value;
  const chartType = typeFilter.value;         /* line | bar */

  /* no stores selected */
  if (!selected.length) {
    if (chartInst) chartInst.destroy();
    canvas.style.display = 'none';
    emptyEl.style.display = 'flex';
    legendEl.innerHTML = '<span class="legend-hint">Selecciona al menos una tienda</span>';
    return;
  }

  /* filter rows */
  let subset = rawData.filter(r =>
    (col(r,'Pais','pais','PAIS') || '—') === curPais &&
    selected.includes(col(r,'Tienda','tienda','TIENDA') || '—')
  );
  if (areaVal !== '__all__') {
    subset = subset.filter(r =>
      (col(r,'Area_Comercial','Area','area') || '—') === areaVal
    );
  }

  const areasX = [...new Set(subset.map(r => col(r,'Area_Comercial','Area','area') || '—'))];

  /* metric label */
  const metricLabels = { m2: 'm²', m3: 'm³', N_RACKS: 'racks' };
  const metricTitles = { m2: 'Metros Cuadrados (m²)', m3: 'Metros Cúbicos (m³)', N_RACKS: 'Número de Racks' };
  const unit   = metricLabels[metrica] || metrica;
  const mTitle = metricTitles[metrica] || metrica;

  ctitleEl.textContent = `Comparativo ${mTitle} — ${curPais}`;
  csubEl.textContent   = `${selected.length} tienda${selected.length > 1 ? 's' : ''} · ${areasX.length} área${areasX.length !== 1 ? 's' : ''}`;

  /* datasets */
  const isLine = chartType === 'line';
  const datasets = selected.map((tienda, i) => {
    const color = PALETTE[i % PALETTE.length];
    const rows  = subset.filter(r => (col(r,'Tienda','tienda','TIENDA') || '—') === tienda);
    const vals  = areasX.map(area => {
      const row = rows.find(r => (col(r,'Area_Comercial','Area','area') || '—') === area);
      if (!row) return 0;
      const v = col(row, metrica, metrica.toUpperCase(), metrica.toLowerCase());
      return parseFloat(v) || 0;
    });
    return {
      label: tienda,
      data: vals,
      borderColor: color,
      backgroundColor: isLine ? color + '22' : color + 'BB',
      borderWidth: 2.5,
      tension: 0.35,
      pointRadius: isLine ? 4 : 0,
      pointHoverRadius: 7,
      fill: isLine,
      borderRadius: isLine ? 0 : 5,
      borderSkipped: false,
    };
  });

  /* legend */
  legendEl.innerHTML = datasets.map(ds => `
    <div class="legend-item">
      <div class="ldot" style="background:${ds.borderColor}"></div>
      ${ds.label}
    </div>`).join('');

  /* chart */
  if (chartInst) chartInst.destroy();
  emptyEl.style.display = 'none';
  canvas.style.display  = 'block';

  chartInst = new Chart(canvas.getContext('2d'), {
    type: chartType,
    data: { labels: areasX, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#111',
          titleColor: '#fff',
          bodyColor: 'rgba(255,255,255,.7)',
          borderColor: 'rgba(241,90,34,.4)',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('es-PE')} ${unit}`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: 'rgba(255,255,255,.42)',
            font: { size: 11, family: 'Barlow' },
            maxRotation: 38
          }
        },
        y: {
          grid: { color: 'rgba(255,255,255,.05)' },
          beginAtZero: true,
          ticks: {
            color: 'rgba(255,255,255,.42)',
            font: { size: 11, family: 'JetBrains Mono' },
            callback: v => v.toLocaleString('es-PE') + ' ' + unit
          }
        }
      },
      animation: { duration: 500, easing: 'easeOutQuart' }
    }
  });
}

/* ── CLEAR DATA ────────────────────────────────────────── */
function clearData() {
  rawData = [];
  if (chartInst) { chartInst.destroy(); chartInst = null; }
  dynCtrl.style.display    = 'none';
  canvas.style.display     = 'none';
  emptyEl.style.display    = 'flex';
  excelInput.value         = '';
  legendEl.innerHTML       = '<span class="legend-hint">La leyenda aparecerá aquí con el color de cada tienda</span>';
  ctitleEl.textContent     = 'Comparativo por área comercial';
  csubEl.textContent       = 'Carga un Excel para comenzar';
  upzone.innerHTML = `
    <div class="upzone-ico">📂</div>
    <p>Arrastra o <b>haz clic</b> para cargar</p>
    <p style="font-size:11px;color:rgba(255,255,255,.3);">.xlsx · .xls · .csv</p>`;
}

/* ── DOWNLOAD TEMPLATE ─────────────────────────────────── */
function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Pais', 'Tienda', 'Area_Comercial', 'm2', 'm3', 'N_RACKS'],
    ['Peru', 'Promart Chorrillos',  'Herramientas',       1850, 7400, 42],
    ['Peru', 'Promart Chorrillos',  'Ferretería',          920, 3680, 21],
    ['Peru', 'Promart Chorrillos',  'Jardín y Mascotas',   640, 2560, 14],
    ['Peru', 'Promart Miraflores',  'Herramientas',       1200, 4800, 28],
    ['Peru', 'Promart Miraflores',  'Ferretería',          680, 2720, 15],
    ['Peru', 'Promart Miraflores',  'Jardín y Mascotas',   480, 1920, 10],
    ['Peru', 'Promart Lima Norte',  'Herramientas',       2200, 8800, 52],
    ['Peru', 'Promart Lima Norte',  'Ferretería',         1100, 4400, 25],
    ['Peru', 'Promart Lima Norte',  'Jardín y Mascotas',   750, 3000, 17],
    ['Peru', 'Promart San Borja',   'Herramientas',        980, 3920, 22],
    ['Peru', 'Promart San Borja',   'Ferretería',          520, 2080, 12],
    ['Chile','Promart Santiago',    'Herramientas',       2100, 8400, 49],
    ['Chile','Promart Santiago',    'Ferretería',          980, 3920, 22],
    ['Chile','Promart Concepción',  'Herramientas',       1650, 6600, 38],
    ['Chile','Promart Concepción',  'Ferretería',          820, 3280, 18],
  ]);
  ws['!cols'] = [12, 22, 22, 10, 10, 10].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');
  XLSX.writeFile(wb, 'plantilla_analisis_promart.xlsx');
}

/* ── NAV ACTIVE ON SCROLL ─────────────────────────────── */
const navLinks = document.querySelectorAll('.menu a');
const observer = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const id = entry.target.id;
      navLinks.forEach(a => a.classList.toggle('active', a.getAttribute('href') === '#' + id));
    }
  });
}, { threshold: 0.3 });
document.querySelectorAll('section[id]').forEach(s => observer.observe(s));
