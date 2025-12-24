// Protótipo funcional: import .xls/.xlsx/.xlsm via SheetJS, exibição, edição e export.
// Persistência: localStorage (chave: kite_data)
// Dependências: Chart.js, XLSX (inclusos via CDN no HTML)

const fileInput = document.getElementById('fileInput');
const loadedSheetEl = document.getElementById('loadedSheet');
const kpisEl = document.getElementById('kpis');
const tableWrap = document.getElementById('tableWrap');
const scoreCtx = document.getElementById('scoreChart').getContext('2d');
const searchInput = document.getElementById('searchInput');
const filterStatus = document.getElementById('filterStatus');
const addBtn = document.getElementById('addBtn');
const exportBtn = document.getElementById('exportBtn');
const clearBtn = document.getElementById('clearBtn');
const prevPageBtn = document.getElementById('prevPage');
const nextPageBtn = document.getElementById('nextPage');
const pageInfo = document.getElementById('pageInfo');
const formPanel = document.getElementById('formPanel');
const evalForm = document.getElementById('evalForm');
const cancelEditBtn = document.getElementById('cancelEdit');

let state = {
  rawSheets: {}, // {sheetName: array of objects}
  currentSheet: null,
  data: [], // array of objects currently in use
  headers: [],
  chart: null,
  page: 1,
  pageSize: 10,
  editIndex: null
};

// util: salvar/ler localStorage
const STORAGE_KEY = 'kite_data_v1';
function saveToStorage() {
  const payload = { rawSheets: state.rawSheets, currentSheet: state.currentSheet };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}
function loadFromStorage() {
  const s = localStorage.getItem(STORAGE_KEY);
  if (!s) return false;
  try {
    const parsed = JSON.parse(s);
    state.rawSheets = parsed.rawSheets || {};
    state.currentSheet = parsed.currentSheet || null;
    return true;
  } catch (e) { return false; }
}

// Ler arquivo Excel via SheetJS
fileInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = evt.target.result;
    const workbook = XLSX.read(data, { type: 'binary' });
    // converter cada aba para array de objetos (XLSX.utils.sheet_to_json)
    workbook.SheetNames.forEach(name => {
      const sheet = workbook.Sheets[name];
      const arr = XLSX.utils.sheet_to_json(sheet, { defval: '' }); // array de objetos
      state.rawSheets[name] = arr;
    });
    // escolher primeira aba automaticamente
    state.currentSheet = workbook.SheetNames[0];
    saveToStorage();
    renderAfterLoad();
  };
  reader.readAsBinaryString(f);
});

// Exibir abas carregadas / dados
function renderAfterLoad() {
  if (!state.currentSheet) {
    // tentar carregar do storage
    if (!loadFromStorage()) {
      loadedSheetEl.textContent = 'Nenhuma aba carregada';
      return;
    }
  }
  loadedSheetEl.textContent = `Aba: ${state.currentSheet}`;
  // merge headers and data
  const arr = state.rawSheets[state.currentSheet] || [];
  state.data = arr.slice(); // copia
  state.headers = inferHeaders(state.data);
  populateStatusFilter();
  state.page = 1;
  renderKPIs();
  renderChart();
  renderTable();
}

// Inferir headers (chaves do objeto)
function inferHeaders(data) {
  const set = new Set();
  data.forEach(row => Object.keys(row).forEach(k => set.add(k)));
  // heurística: ordem preferida
  const preferred = ['Nome','Matrícula','Matricula','Cargo','Data','Data Avaliação','Nota','Pontuacao','Pontuação','Status','Comentários','Comentarios','Observacao','Observações'];
  const keys = Array.from(set);
  const sorted = Array.from(new Set(preferred.concat(keys))).filter(k => set.has(k));
  // se sobraram chaves não cobertas, adiciona
  keys.forEach(k => { if (!sorted.includes(k)) sorted.push(k); });
  return sorted;
}

// KPIs simples
function renderKPIs() {
  kpisEl.innerHTML = '';
  const total = state.data.length;
  const filled = state.data.filter(r => Object.values(r).some(v => String(v).trim() !== '')).length;
  const avgScore = computeAverageScore();

  const cards = [
    { label: 'Registros', value: total },
    { label: 'Registros preenchidos', value: filled },
    { label: 'Média de nota', value: isNaN(avgScore) ? '—' : avgScore.toFixed(1) },
    { label: 'Colunas detectadas', value: state.headers.length }
  ];

  cards.forEach((c,i) => {
    const d = document.createElement('div');
    d.className = 'glass kpi';
    d.innerHTML = `<div class="label">${c.label}</div><div class="value">${c.value}</div>`;
    kpisEl.appendChild(d);
  });
}

// calcular média a partir de possíveis nomes de coluna
function computeAverageScore() {
  const scoreKeys = ['Nota','Pontuacao','Pontuação','Score','Score'];
  let sum = 0, count = 0;
  state.data.forEach(r => {
    for (let k of scoreKeys) {
      if (k in r) {
        const v = String(r[k]).replace(',', '.');
        const n = Number(v);
        if (!isNaN(n)) { sum += n; count++; }
        break;
      }
    }
  });
  return count ? sum / count : NaN;
}

// render chart (distribuição das notas)
function renderChart() {
  const labels = [];
  const data = [];
  // pegar até 100 registros
  for (let i=0;i<Math.min(100,state.data.length);i++) {
    const r = state.data[i];
    const name = r['Nome'] || r['name'] || r['Nome completo'] || `R${i+1}`;
    labels.push(name);
    // encontrar nota em chaves prováveis
    const val = ['Nota','Pontuacao','Pontuação','Score'].reduce((acc,k) => acc ?? r[k], undefined);
    const n = Number(String(val ?? '').replace(',', '.'));
    data.push(isNaN(n) ? 0 : n);
  }

  if (state.chart) { state.chart.destroy(); state.chart = null; }
  state.chart = new Chart(scoreCtx, {
    type: 'bar',
    data: { labels, datasets: [{ label:'Notas (amostra)', data, backgroundColor: 'rgba(124,231,196,0.7)' }] },
    options: { responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}} }
  });
}

// preencher filtro de status com valores detectados
function populateStatusFilter() {
  const statuses = new Set();
  state.data.forEach(r => {
    const s = r['Status'] || r['status'] || r['Situação'] || '';
    if (s && String(s).trim() !== '') statuses.add(String(s).trim());
  });
  filterStatus.innerHTML = '<option value="">Todos status</option>';
  Array.from(statuses).forEach(s => {
    const opt = document.createElement('option'); opt.value = s; opt.textContent = s;
    filterStatus.appendChild(opt);
  });
}

// render tabela com paginação e filtros
function renderTable() {
  tableWrap.innerHTML = '';
  if (!state.data || state.data.length === 0) {
    tableWrap.innerHTML = '<div style="color:var(--muted)">Nenhum dado</div>';
    pageInfo.textContent = 'Página 0/0';
    return;
  }

  const query = (searchInput.value || '').toLowerCase();
  const statusFilter = filterStatus.value;
  let filtered = state.data.filter(r => {
    let matchQ = true;
    if (query) {
      const hay = Object.values(r).join(' ').toLowerCase();
      matchQ = hay.includes(query);
    }
    let matchS = true;
    if (statusFilter) {
      const s = r['Status'] || r['status'] || '';
      matchS = String(s).trim() === statusFilter;
    }
    return matchQ && matchS;
  });

  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  if (state.page > totalPages) state.page = totalPages;
  const start = (state.page - 1) * state.pageSize;
  const pageRows = filtered.slice(start, start + state.pageSize);

  // montar tabela
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerKeys = state.headers;
  thead.innerHTML = `<tr>${headerKeys.map(h => `<th>${h}</th>`).join('')}<th>Ações</th></tr>`;
  table.appendChild(thead);
  const tbody = document.createElement('tbody');

  pageRows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    headerKeys.forEach(k => {
      const v = row[k] ?? '';
      tr.innerHTML += `<td>${escapeHtml(String(v))}</td>`;
    });
    const globalIndex = state.data.indexOf(pageRows[idx]);
    const actions = `<td class="table-actions">
      <button data-index="${globalIndex}" class="editBtn">Editar</button>
      <button data-index="${globalIndex}" class="delBtn">Excluir</button>
    </td>`;
    tr.innerHTML += actions;
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tableWrap.appendChild(table);

  pageInfo.textContent = `Página ${state.page}/${totalPages}`;

  // eventos ações
  Array.from(document.getElementsByClassName('editBtn')).forEach(b => {
    b.onclick = (ev) => {
      const i = Number(ev.currentTarget.dataset.index);
      openEditForm(i);
    };
  });
  Array.from(document.getElementsByClassName('delBtn')).forEach(b => {
    b.onclick = (ev) => {
      const i = Number(ev.currentTarget.dataset.index);
      if (confirm('Excluir este registro?')) {
        state.data.splice(i,1);
        // atualizar rawSheets
        state.rawSheets[state.currentSheet] = state.data;
        saveToStorage();
        renderAfterLoad();
      }
    };
  });
}

// helper: escapar HTML
function escapeHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// abrir formulário para adicionar/editar
addBtn.addEventListener('click', () => openEditForm(null));

function openEditForm(index) {
  formPanel.style.display = 'block';
  document.getElementById('tablePanel').style.display = 'none';
  document.getElementById('dashboardPanel').style.display = 'none';
  evalForm.reset();
  state.editIndex = index;
  if (index !== null) {
    const row = state.data[index];
    for (let k of Object.keys(row)) {
      const el = evalForm.elements.namedItem(k);
      if (el) el.value = row[k];
    }
  }
}

// cancelar
cancelEditBtn.addEventListener('click', () => {
  formPanel.style.display = 'none';
  document.getElementById('tablePanel').style.display = 'block';
  document.getElementById('dashboardPanel').style.display = 'block';
  state.editIndex = null;
});

// submeter form
evalForm.addEventListener('submit', (e) => {
  e.preventDefault();
  // coletar valores
  const formData = new FormData(evalForm);
  const obj = {};
  // garantir que headers contenham os campos do form
  for (let [k,v] of formData.entries()) {
    obj[k] = v;
    if (!state.headers.includes(k)) state.headers.push(k);
  }
  if (state.editIndex === null) {
    // adicionar
    state.data.push(obj);
  } else {
    // editar
    state.data[state.editIndex] = Object.assign({}, state.data[state.editIndex], obj);
  }
  state.rawSheets[state.currentSheet] = state.data;
  saveToStorage();
  formPanel.style.display = 'none';
  document.getElementById('tablePanel').style.display = 'block';
  document.getElementById('dashboardPanel').style.display = 'block';
  renderAfterLoad();
});

// export CSV do sheet atual
exportBtn.addEventListener('click', () => {
  if (!state.data || state.data.length === 0) return alert('Nenhum dado a exportar.');
  const keys = state.headers;
  const lines = [keys.join(',')];
  state.data.forEach(r => {
    const row = keys.map(k => `"${String(r[k] ?? '').replace(/"/g,'""')}"`).join(',');
    lines.push(row);
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'export_kite.csv'; a.click();
  URL.revokeObjectURL(url);
});

// busca e filtro
searchInput.addEventListener('input', () => { state.page = 1; renderTable(); });
filterStatus.addEventListener('change', () => { state.page = 1; renderTable(); });

// paginação
prevPageBtn.addEventListener('click', () => { if (state.page>1) { state.page--; renderTable(); } });
nextPageBtn.addEventListener('click', () => { state.page++; renderTable(); });

// limpar dados locais
clearBtn.addEventListener('click', () => {
  if (!confirm('Remover dados carregados e o armazenamento local?')) return;
  localStorage.removeItem(STORAGE_KEY);
  state = { rawSheets:{}, currentSheet:null, data:[], headers:[], chart:null, page:1, pageSize:10, editIndex:null };
  loadedSheetEl.textContent = 'Nenhuma aba carregada';
  renderAfterLoad();
});

// inicial: tentar carregar do storage
if (loadFromStorage()) renderAfterLoad();
