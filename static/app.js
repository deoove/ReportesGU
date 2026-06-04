// ── Config ────────────────────────────────────────────────────────────────────
// Usa la API si se abre desde el servidor (http://). Con doble clic (file://) avisa.
const API = window.location.protocol === 'file:'
  ? null
  : `${window.location.protocol}//${window.location.host}/api`;

// ── Estado global ─────────────────────────────────────────────────────────────
let DATA         = [];
let newOdsSet    = new Set();
let sortCol      = 'ods';
let sortAsc      = true;
let filtered     = [];
let pendingNew   = [];
let pendingDups  = [];
let odsBorrarActual  = null;
let odsEditarActual  = null;

// ── Definición de campos ──────────────────────────────────────────────────────
const CAMPOS = [
  'ods', 'ubicacion_tecnica', 'fecha_generacion', 'fecha_entrega', 'fecha_inspeccion',
  'hora_llegada', 'hora_salida', 'inspector', 'distrito', 'exp_ref', 'nro_serie',
  'material', 'lectura', 'funcionamiento', 'resid', 'actividad', 'estado_conexion',
  'tipo_caja', 'movil', 'precinto', 'pozo', 'valvula_retencion', 'cx_libres',
  'corte', 'observaciones', 'resuelve', 'motivo', 'alto_consumo',
  'ods_inconsistencia_rcc', 'zrec', 'ods_recambio', 'fecha_comunicacion', 'nro_contacto',
];

const LABELS = {
  ods:                  'ODS',
  ubicacion_tecnica:    'Ubic. Técnica',
  fecha_generacion:     'F. Generación',
  fecha_entrega:        'F. Entrega',
  fecha_inspeccion:     'F. Inspección',
  hora_llegada:         'Hora Llegada',
  hora_salida:          'Hora Salida',
  inspector:            'Inspector',
  distrito:             'Distrito',
  exp_ref:              'Exp. Ref.',
  nro_serie:            'N° Serie',
  material:             'Material',
  lectura:              'Lectura',
  funcionamiento:       'Funcionamiento',
  resid:                'Resid/No Resid',
  actividad:            'Actividad',
  estado_conexion:      'Estado Cx.',
  tipo_caja:            'Tipo Caja',
  movil:                'Móvil',
  precinto:             'Precinto',
  pozo:                 'Pozo',
  valvula_retencion:    'Válvula Ret.',
  cx_libres:            'Cx. Libres',
  corte:                'Corte',
  observaciones:        'Observaciones',
  resuelve:             'Resuelve',
  motivo:               'Motivo',
  alto_consumo:         'Alto Consumo',
  ods_inconsistencia_rcc: 'ODS Inc. RCC',
  zrec:                 'ZREC',
  ods_recambio:         'ODS Recambio',
  fecha_comunicacion:   'F. Comunicación',
  nro_contacto:         'N° Contacto',
};

// Columnas del Excel en orden (coincide con el .xls original)
const COLS_XLS = [
  'ods', 'ubicacion_tecnica', 'sem', 'bim', 'anio', 'fecha_generacion', 'rcc',
  'primera_lect', 'aud', 'consumo_cero', 'integral', 'distrito', 'analisis', 'repite',
  'fecha_entrega', 'inspector', 'exp_ref', 'nro_serie', 'material', 'lectura',
  'fecha_inspeccion', 'hora_llegada', 'hora_salida', 'funcionamiento', 'resid',
  'actividad', 'corte', 'precinto', 'pozo', 'valvula_retencion', 'estado_conexion',
  'tipo_caja', 'movil', 'cx_libres', 'observaciones', 'resuelve', 'motivo',
  'ods_inconsistencia_rcc', 'zrec', 'ods_recambio', 'alto_consumo',
  'fecha_comunicacion', 'nro_contacto',
];


// ── API helpers ───────────────────────────────────────────────────────────────

async function apiFetch(path, opts = {}) {
  if (!API) {
    showToast('Abrí la página desde el servidor: http://localhost:8000', 'error');
    throw new Error('No API');
  }
  const res = await fetch(API + path, {
    headers: { 'Content-Type': 'application/json' },
    ...opts,
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}


// ── Carga de datos desde la API ───────────────────────────────────────────────

async function loadData() {
  if (!API) {
    showToast('Usá el servidor local: doble clic en arrancar.bat y abrí http://localhost:8000', 'warn');
    return;
  }
  try {
    showLoading('Cargando registros...');
    DATA = await apiFetch('/registros');
  } catch (e) {
    showToast('No se pudo conectar al servidor.', 'error');
    DATA = [];
  } finally {
    hideLoading();
  }
}


// ── Parseo de Excel ───────────────────────────────────────────────────────────

function xlsxDate(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number' && val > 1000) {
    const d = XLSX.SSF.parse_date_code(val);
    if (!d) return String(val);
    return `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}`;
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})$/);
  if (m) return `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`;
  return s;
}

function xlsxHora(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number' && val > 0 && val < 1) {
    const t = Math.round(val * 24 * 60);
    return `${String(Math.floor(t / 60)).padStart(2, '0')}:${String(t % 60).padStart(2, '0')}`;
  }
  return String(val).trim();
}

function cleanNum(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number') return val === Math.floor(val) ? String(Math.floor(val)) : String(val);
  return String(val).trim();
}

function parseXlsRow(raw) {
  const r = {};
  COLS_XLS.forEach((col, i) => {
    const v = raw[i];
    if ([5, 14, 20, 41].includes(i))  r[col] = xlsxDate(v);
    else if ([21, 22].includes(i))    r[col] = xlsxHora(v);
    else                              r[col] = cleanNum(v);
  });
  return r;
}

function handleFile(event) {
  const file = event.target.files[0];
  event.target.value = '';
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const wb    = XLSX.read(e.target.result, { type: 'array' });
      const sheet = wb.Sheets['REGISTRO_GENERAL'];
      if (!sheet) { showToast('No se encontró la hoja REGISTRO_GENERAL.', 'error'); return; }

      const rows     = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
      const incoming = rows.slice(1).filter(r => r.length > 0 && r[0]).map(parseXlsRow);
      if (!incoming.length) { showToast('El archivo no tiene datos.', 'warn'); return; }

      const existingMap = {};
      DATA.forEach(d => { existingMap[d.ods] = d; });

      pendingNew  = incoming.filter(r => !existingMap[r.ods]);
      pendingDups = incoming.filter(r =>  existingMap[r.ods]).map(r => ({ newRow: r, oldRow: existingMap[r.ods] }));

      if (!pendingDups.length) applyImport(pendingNew, []);
      else openDupModal(pendingDups, pendingNew);
    } catch (err) {
      showToast('Error al leer el archivo: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}


// ── Importación con detección de duplicados ───────────────────────────────────

function openDupModal(dups, news) {
  document.getElementById('dupDesc').textContent =
    `Se encontraron ${dups.length} ODS ya existentes y ${news.length} nuevas. Marcá las que querés pisar:`;
  document.getElementById('checkAll').checked = false;
  document.getElementById('dupTableBody').innerHTML = dups.map((d, i) => `
    <tr>
      <td style="text-align:center">
        <input type="checkbox" class="dup-cb" data-idx="${i}">
      </td>
      <td><strong>${d.newRow.ods}</strong></td>
      <td>${d.oldRow.inspector || '—'}</td>
      <td>${d.newRow.inspector || '—'}</td>
      <td>${d.oldRow.fecha_generacion || '—'}</td>
      <td>${d.newRow.fecha_generacion || '—'}</td>
    </tr>`).join('');
  openModal('modalDups');
}

function toggleAll(cb) {
  document.querySelectorAll('.dup-cb').forEach(c => c.checked = cb.checked);
}

function cancelImport() {
  closeModal('modalDups');
  pendingNew = [];
  pendingDups = [];
  showToast('Importación cancelada.', 'warn');
}

function confirmImport(useCheckboxes) {
  let overwriteOds = [];
  if (useCheckboxes) {
    document.querySelectorAll('.dup-cb:checked').forEach(cb => {
      overwriteOds.push(pendingDups[parseInt(cb.dataset.idx)].newRow.ods);
    });
  }
  closeModal('modalDups');
  applyImport([...pendingNew, ...pendingDups.map(d => d.newRow)], overwriteOds);
  pendingNew  = [];
  pendingDups = [];
}

async function applyImport(rows, overwriteOds) {
  try {
    showLoading('Guardando en la base de datos...');
    const result = await apiFetch('/registros', {
      method: 'POST',
      body: JSON.stringify({ rows, overwrite: overwriteOds }),
    });
    await loadData();
    rebuildFilters();
    applyFilters();

    const msg = [];
    if (result.agregados)    msg.push(`${result.agregados} ODS agregadas`);
    if (result.actualizados) msg.push(`${result.actualizados} ODS actualizadas`);
    showToast(msg.join(' · ') || 'Sin cambios.', 'success');

    // Marcar nuevas en verde por 8 segundos
    const nuevasOds = rows.filter(r => !overwriteOds.includes(r.ods)).map(r => r.ods);
    DATA.forEach(d => { if (nuevasOds.includes(d.ods)) newOdsSet.add(d.ods); });
    renderTable();
    setTimeout(() => { newOdsSet.clear(); renderTable(); }, 8000);
  } catch (e) {
    showToast('Error al guardar: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}


// ── Detalle ───────────────────────────────────────────────────────────────────

function verDetalle(ods) {
  const d = DATA.find(x => x.ods === ods);
  if (!d) return;

  document.getElementById('detalleTitulo').textContent = `ODS ${d.ods}`;

  const sections = [
    { title: 'Identificación', fields: ['ods', 'ubicacion_tecnica', 'exp_ref', 'nro_serie', 'material', 'distrito'] },
    { title: 'Fechas',         fields: ['fecha_generacion', 'fecha_entrega', 'fecha_inspeccion', 'hora_llegada', 'hora_salida', 'fecha_comunicacion'] },
    { title: 'Inspección',     fields: ['inspector', 'lectura', 'funcionamiento', 'estado_conexion', 'tipo_caja', 'movil', 'resid', 'actividad'] },
    { title: 'Condiciones',    fields: ['precinto', 'pozo', 'valvula_retencion', 'cx_libres', 'corte'] },
    { title: 'Gestión',        fields: ['resuelve', 'motivo', 'zrec', 'ods_recambio', 'alto_consumo', 'nro_contacto', 'observaciones'] },
  ];

  document.getElementById('detalleBody').innerHTML = sections.map(s => `
    <div class="detail-section">
      <h3>${s.title}</h3>
      ${s.fields.filter(f => d[f]).map(f => `
        <div class="detail-row">
          <span class="lbl">${LABELS[f] || f}</span>
          <span class="val">${d[f]}</span>
        </div>`).join('') || '<div style="color:#a0aec0;font-size:0.82rem">Sin datos</div>'}
    </div>`).join('');

  openModal('modalDetalle');
}


// ── Editar ────────────────────────────────────────────────────────────────────

function verEditar(ods) {
  const d = DATA.find(x => x.ods === ods);
  if (!d) return;
  odsEditarActual = ods;
  document.getElementById('editarTitulo').textContent = `Editar ODS ${ods}`;
  document.getElementById('formGrid').innerHTML = CAMPOS.map(c => `
    <div class="form-group${c === 'observaciones' || c === 'motivo' ? ' full' : ''}">
      <label>${LABELS[c] || c}</label>
      <input type="text" name="${c}" value="${(d[c] || '').replace(/"/g, '&quot;')}"
        ${c === 'ods' ? 'readonly style="background:#f7fafc;color:#718096"' : ''}>
    </div>`).join('');
  openModal('modalEditar');
}

async function guardarEdicion(e) {
  e.preventDefault();
  const form = document.getElementById('formEditar');
  const data = {};
  CAMPOS.forEach(c => { data[c] = form.elements[c]?.value || ''; });
  try {
    showLoading('Guardando...');
    await apiFetch(`/registros/${encodeURIComponent(odsEditarActual)}`, {
      method: 'PUT',
      body: JSON.stringify(data),
    });
    closeModal('modalEditar');
    await loadData();
    rebuildFilters();
    applyFilters();
    showToast('Registro actualizado correctamente.', 'success');
  } catch (err) {
    showToast('Error al guardar: ' + err.message, 'error');
  } finally {
    hideLoading();
  }
}


// ── Eliminar ──────────────────────────────────────────────────────────────────

function verBorrar(ods) {
  odsBorrarActual = ods;
  document.getElementById('borrarOds').textContent = ods;
  openModal('modalBorrar');
}

async function confirmarBorrar() {
  try {
    showLoading('Eliminando...');
    await apiFetch(`/registros/${encodeURIComponent(odsBorrarActual)}`, { method: 'DELETE' });
    closeModal('modalBorrar');
    await loadData();
    rebuildFilters();
    applyFilters();
    showToast('Registro eliminado.', 'success');
  } catch (err) {
    showToast('Error al eliminar: ' + err.message, 'error');
  } finally {
    hideLoading();
  }
}


// ── Render de la tabla ────────────────────────────────────────────────────────

function badgeEstado(val) {
  if (!val) return '<span style="color:#a0aec0">—</span>';
  const cls = val === 'Bueno'  ? 'badge-bueno'  :
              val === 'Malo'   ? 'badge-malo'   :
              val === 'Normal' ? 'badge-normal' : 'badge-pendiente';
  return `<span class="badge ${cls}">${val}</span>`;
}

function buildStats() {
  const total       = DATA.length;
  const conInspector  = DATA.filter(d => d.inspector).length;
  const conInspeccion = DATA.filter(d => d.fecha_inspeccion).length;
  const buenos      = DATA.filter(d => d.estado_conexion === 'Bueno').length;
  const malos       = DATA.filter(d => d.estado_conexion === 'Malo').length;
  const distritos   = new Set(DATA.map(d => d.distrito).filter(Boolean)).size;

  document.getElementById('statsBar').innerHTML = [
    { val: total,         lbl: 'Total ODS' },
    { val: conInspector,  lbl: 'Con Inspector' },
    { val: conInspeccion, lbl: 'Inspeccionados' },
    { val: buenos,        lbl: 'Estado Bueno' },
    { val: malos,         lbl: 'Estado Malo' },
    { val: distritos,     lbl: 'Distritos' },
  ].map(s => `
    <div class="stat-card">
      <div class="val">${s.val}</div>
      <div class="lbl">${s.lbl}</div>
    </div>`).join('');
}

function rebuildFilters() {
  buildStats();

  const dSel = document.getElementById('filterDistrito');
  const iSel = document.getElementById('filterInspector');
  const dVal = dSel.value;
  const iVal = iSel.value;

  dSel.innerHTML = '<option value="">Todos</option>';
  [...new Set(DATA.map(d => d.distrito).filter(Boolean))].sort((a, b) => a - b)
    .forEach(d => dSel.innerHTML += `<option value="${d}">${d}</option>`);

  iSel.innerHTML = '<option value="">Todos</option>';
  [...new Set(DATA.map(d => d.inspector).filter(Boolean))].sort()
    .forEach(i => iSel.innerHTML += `<option value="${i}">${i}</option>`);

  dSel.value = dVal;
  iSel.value = iVal;
}

function applyFilters() {
  const q    = document.getElementById('searchInput').value.toLowerCase();
  const dist = document.getElementById('filterDistrito').value;
  const insp = document.getElementById('filterInspector').value;
  const est  = document.getElementById('filterEstado').value;

  filtered = DATA.filter(d => {
    const s = [d.ods, d.inspector, d.nro_serie, d.distrito, d.material, d.movil, d.exp_ref]
      .join(' ').toLowerCase();
    return (!q    || s.includes(q))          &&
           (!dist || d.distrito  === dist)   &&
           (!insp || d.inspector === insp)   &&
           (!est  || d.estado_conexion === est);
  });

  filtered.sort((a, b) => {
    let va = a[sortCol] || '', vb = b[sortCol] || '';
    if (!isNaN(va) && !isNaN(vb)) { va = parseFloat(va) || 0; vb = parseFloat(vb) || 0; }
    return va < vb ? (sortAsc ? -1 : 1) : va > vb ? (sortAsc ? 1 : -1) : 0;
  });

  renderTable();
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  document.getElementById('countLabel').textContent = `${filtered.length} de ${DATA.length} registros`;

  if (!filtered.length) {
    tbody.innerHTML = '';
    document.getElementById('noResults').style.display = 'block';
    return;
  }
  document.getElementById('noResults').style.display = 'none';

  tbody.innerHTML = filtered.map(d => `
    <tr class="${newOdsSet.has(d.ods) ? 'new-row' : ''}">
      <td>
        <strong>${d.ods}</strong>
        ${newOdsSet.has(d.ods) ? ' <span style="color:#276749;font-size:0.72rem;font-weight:700">NUEVO</span>' : ''}
      </td>
      <td>${d.fecha_generacion || '—'}</td>
      <td>${d.fecha_entrega    || '—'}</td>
      <td>${d.fecha_inspeccion || '—'}</td>
      <td>${d.inspector || '<span style="color:#a0aec0">Sin asignar</span>'}</td>
      <td><span class="badge badge-pendiente">${d.distrito || '—'}</span></td>
      <td>${d.nro_serie  || '—'}</td>
      <td>${d.material   || '—'}</td>
      <td>${d.lectura    || '—'}</td>
      <td>${badgeEstado(d.funcionamiento)}</td>
      <td>${badgeEstado(d.estado_conexion)}</td>
      <td>${d.resid ? d.resid.trim() : '—'}</td>
      <td>
        <button class="btn-accion btn-ver"    onclick="verDetalle('${d.ods}')">Ver</button>
        <button class="btn-accion btn-editar" onclick="verEditar('${d.ods}')">Editar</button>
        <button class="btn-accion btn-borrar" onclick="verBorrar('${d.ods}')">Borrar</button>
      </td>
    </tr>`).join('');
}

function sortTable(col) {
  sortAsc  = sortCol === col ? !sortAsc : true;
  sortCol  = col;
  applyFilters();
}


// ── Modales ───────────────────────────────────────────────────────────────────

function openModal(id)  { document.getElementById(id).classList.add('open');    }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    ['modalDetalle', 'modalEditar', 'modalBorrar', 'modalDups'].forEach(closeModal);
  }
});


// ── Loading ───────────────────────────────────────────────────────────────────

function showLoading(msg = 'Cargando...') {
  document.getElementById('loadingMsg').textContent = msg;
  document.getElementById('loadingOverlay').classList.add('show');
}

function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('show');
}


// ── Toast ─────────────────────────────────────────────────────────────────────

let toastTimer;

function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent  = msg;
  t.className    = `toast show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = 'toast'; }, 4000);
}


// ── Inicialización ────────────────────────────────────────────────────────────

document.getElementById('searchInput').addEventListener('input',   applyFilters);
document.getElementById('filterDistrito').addEventListener('change', applyFilters);
document.getElementById('filterInspector').addEventListener('change', applyFilters);
document.getElementById('filterEstado').addEventListener('change',  applyFilters);

(async function init() {
  await loadData();
  rebuildFilters();
  applyFilters();
})();
