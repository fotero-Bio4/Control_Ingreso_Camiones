'use strict';

const ALLOWED_TYPES  = ['image/jpeg', 'image/png', 'image/heic', 'image/heif', 'image/webp'];
const ALLOWED_EXT    = /\.(jpe?g|png|heic|heif|webp)$/i;
const MAX_SIZE_BYTES = 8 * 1024 * 1024; // 8 MB

// Campos que se vuelven obligatorios en modo Alta, agrupados por sección
const ALTA_FIELDS = {
  chofer: ['nombre', 'empresa', 'petrolera', 'patente_ch_asig',
           'venc_carnet', 'img_carnet', 'venc_psico', 'img_psico',
           'venc_cargas', 'img_cargas'],
  unidad: ['venc_itv_chas', 'img_itv_chas', 'venc_itv_cis', 'img_itv_cis'],
};

let currentModo = null;
let currentTipo = null;

// ── Paso 1: selección de modo ───────────────────────────────────────────────
document.querySelectorAll('input[name="modo"]').forEach(radio => {
  radio.addEventListener('change', function () {
    currentModo = this.value;
    currentTipo = null; // resetear tipo al cambiar modo

    // Desmarcar tipo y ocultar secciones
    document.querySelectorAll('input[name="tipo"]').forEach(r => { r.checked = false; });
    document.getElementById('sectionChofer').classList.add('hidden');
    document.getElementById('sectionUnidad').classList.add('hidden');
    document.getElementById('submitArea').classList.add('hidden');

    // Mostrar card de tipo
    document.getElementById('cardTipo').classList.remove('hidden');

    // Indicador de modo dentro del card de tipo
    const ind = document.getElementById('modoIndicator');
    if (currentModo === 'alta') {
      ind.className = 'modo-indicator modo-indicator-alta';
      ind.textContent = '📝 Alta Documentación — todos los campos marcados con * son obligatorios';
    } else {
      ind.className = 'modo-indicator modo-indicator-act';
      ind.textContent = '🔄 Actualizar Documentación — solo completá lo que cambió';
    }

    // Actualizar visibilidad de asteriscos de Alta
    actualizarAsteriscos();

    document.getElementById('cardTipo').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  });
});

// ── Paso 2: selección de tipo ───────────────────────────────────────────────
document.querySelectorAll('input[name="tipo"]').forEach(radio => {
  radio.addEventListener('change', function () {
    currentTipo = this.value;
    document.getElementById('sectionChofer').classList.toggle('hidden', currentTipo === 'unidad');
    document.getElementById('sectionUnidad').classList.toggle('hidden', currentTipo === 'chofer');
    document.getElementById('submitArea').classList.remove('hidden');

    // Aplicar/quitar required según modo y tipo activo
    aplicarRequired();

    const primera = currentTipo === 'unidad'
      ? document.getElementById('sectionUnidad')
      : document.getElementById('sectionChofer');
    primera.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
});

// ── Lógica de campos obligatorios ───────────────────────────────────────────
function aplicarRequired() {
  const isAlta      = currentModo === 'alta';
  const choferActivo = currentTipo === 'chofer' || currentTipo === 'ambos';
  const unidadActivo = currentTipo === 'unidad' || currentTipo === 'ambos';

  document.querySelectorAll('[data-alta]').forEach(el => {
    const sect    = el.dataset.alta;
    const activo  = (sect === 'chofer' && choferActivo) || (sect === 'unidad' && unidadActivo);
    el.required   = isAlta && activo;
  });
}

function actualizarAsteriscos() {
  const isAlta = currentModo === 'alta';
  document.querySelectorAll('.req-alta').forEach(el => {
    el.style.display = isAlta ? 'inline' : 'none';
  });
}

// ── Auto-mayúsculas en campos de patente ────────────────────────────────────
document.querySelectorAll('.uppercase').forEach(el => {
  el.addEventListener('input', function () {
    const pos = this.selectionStart;
    this.value = this.value.toUpperCase();
    this.setSelectionRange(pos, pos);
  });
});

// ── Manejo de carga de fotos ────────────────────────────────────────────────
document.querySelectorAll('.foto-input').forEach(input => {
  input.addEventListener('change', function () {
    const fieldId   = this.id.replace('f_img_', '');
    const previewEl = document.getElementById('prev_' + fieldId);
    const fnEl      = document.getElementById('fn_' + fieldId);
    if (!previewEl) return;

    previewEl.innerHTML = '';
    if (fnEl) fnEl.textContent = '';

    const file = this.files[0];
    if (!file) return;

    if (file.size > MAX_SIZE_BYTES) {
      alert(`"${file.name}" supera el límite de 8 MB. Elegí una imagen más pequeña.`);
      this.value = '';
      return;
    }

    if (!ALLOWED_TYPES.includes(file.type) && !ALLOWED_EXT.test(file.name)) {
      alert(`Solo se permiten imágenes (JPG, PNG, HEIC, WebP).\n"${file.name}" no es válido.`);
      this.value = '';
      return;
    }

    if (fnEl) fnEl.textContent = file.name;

    const canPreview = file.type !== 'image/heic' && file.type !== 'image/heif';
    if (canPreview) {
      const reader = new FileReader();
      reader.onload = e => {
        const img = document.createElement('img');
        img.src = e.target.result;
        img.alt = file.name;
        previewEl.appendChild(img);
      };
      reader.readAsDataURL(file);
    }

    const info = document.createElement('div');
    info.className = 'foto-info';
    info.textContent = `✓ ${file.name} — ${(file.size / 1024).toFixed(0)} KB`;
    previewEl.appendChild(info);
  });
});

// ── Envío del formulario ────────────────────────────────────────────────────
document.getElementById('mainForm').addEventListener('submit', async function (e) {
  e.preventDefault();

  if (!currentModo) {
    showResult('error', 'Seleccioná si es Alta o Actualización antes de continuar.');
    return;
  }
  if (!currentTipo) {
    showResult('error', 'Seleccioná qué datos vas a cargar (Chofer, Unidad o Ambos).');
    return;
  }

  // Validar clave siempre obligatoria
  if (currentTipo === 'chofer' || currentTipo === 'ambos') {
    const dni = document.querySelector('[name="dni"]').value.trim();
    if (!dni) {
      showResult('error', 'El DNI / CUIT es obligatorio.');
      document.querySelector('[name="dni"]').focus();
      return;
    }
  }
  if (currentTipo === 'unidad' || currentTipo === 'ambos') {
    const pat = document.querySelector('[name="patente_ch"]').value.trim();
    if (!pat) {
      showResult('error', 'La Patente Chasis es obligatoria.');
      document.querySelector('[name="patente_ch"]').focus();
      return;
    }
  }

  // Validar campos Alta adicionales manualmente (por si el browser no frenó)
  if (currentModo === 'alta') {
    const err = validarCamposAlta();
    if (err) { showResult('error', err); return; }
  }

  const btn      = document.getElementById('btnSubmit');
  const resultEl = document.getElementById('resultMsg');
  resultEl.className = 'hidden';

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Enviando...';

  try {
    const formData = new FormData(this);
    formData.set('tipo', currentTipo);
    formData.set('modo', currentModo);

    const resp = await fetch('/.netlify/functions/submit', {
      method: 'POST',
      body: formData,
    });

    let data;
    try { data = await resp.json(); }
    catch { data = { error: `Error HTTP ${resp.status}` }; }

    if (resp.ok && data.ok) {
      showResult('ok', `✓ ${data.message}`);
      resetForm();
    } else {
      showResult('error', `✗ ${data.error || 'Error desconocido al procesar el formulario.'}`);
    }
  } catch (err) {
    showResult('error', `✗ Error de conexión: ${err.message}`);
  } finally {
    btn.disabled = false;
    btn.innerHTML = 'Enviar Formulario';
  }
});

function validarCamposAlta() {
  const labels = {
    nombre: 'Apellido y Nombre', empresa: 'Empresa Transporte',
    petrolera: 'Petrolera', patente_ch_asig: 'Patente Chasis asignada',
    venc_carnet: 'Fecha venc. Carnet', venc_psico: 'Fecha venc. Psicofísico',
    venc_cargas: 'Fecha venc. Cargas Peligrosas',
    venc_itv_chas: 'Fecha venc. ITV Chasis', venc_itv_cis: 'Fecha venc. ITV Cisterna',
  };
  const fotoLabels = {
    img_carnet: 'Foto del Carnet', img_psico: 'Foto del Psicofísico',
    img_cargas: 'Foto de Cargas Peligrosas',
    img_itv_chas: 'Foto ITV Chasis', img_itv_cis: 'Foto ITV Cisterna',
  };

  const choferActivo = currentTipo === 'chofer' || currentTipo === 'ambos';
  const unidadActivo = currentTipo === 'unidad' || currentTipo === 'ambos';

  const camposChofer = ['nombre','empresa','petrolera','patente_ch_asig',
                        'venc_carnet','venc_psico','venc_cargas'];
  const fotosChofer  = ['img_carnet','img_psico','img_cargas'];
  const camposUnidad = ['venc_itv_chas','venc_itv_cis'];
  const fotosUnidad  = ['img_itv_chas','img_itv_cis'];

  if (choferActivo) {
    for (const name of camposChofer) {
      const el = document.querySelector(`[name="${name}"]`);
      if (!el || !el.value.trim()) return `Campo obligatorio en Alta: ${labels[name]}`;
    }
    for (const name of fotosChofer) {
      const el = document.querySelector(`[name="${name}"]`);
      if (!el || !el.files || !el.files[0]) return `Obligatorio en Alta: ${fotoLabels[name]}`;
    }
  }
  if (unidadActivo) {
    for (const name of camposUnidad) {
      const el = document.querySelector(`[name="${name}"]`);
      if (!el || !el.value.trim()) return `Campo obligatorio en Alta: ${labels[name]}`;
    }
    for (const name of fotosUnidad) {
      const el = document.querySelector(`[name="${name}"]`);
      if (!el || !el.files || !el.files[0]) return `Obligatorio en Alta: ${fotoLabels[name]}`;
    }
  }
  return null;
}

function resetForm() {
  document.getElementById('mainForm').reset();
  currentModo = null;
  currentTipo = null;
  document.getElementById('cardTipo').classList.add('hidden');
  document.getElementById('sectionChofer').classList.add('hidden');
  document.getElementById('sectionUnidad').classList.add('hidden');
  document.getElementById('submitArea').classList.add('hidden');
  document.getElementById('modoIndicator').className = 'modo-indicator';
  document.getElementById('modoIndicator').textContent = '';
  document.querySelectorAll('.foto-preview').forEach(el => { el.innerHTML = ''; });
  document.querySelectorAll('.foto-filename').forEach(el => { el.textContent = ''; });
  document.querySelectorAll('[data-alta]').forEach(el => { el.required = false; });
  document.querySelectorAll('.req-alta').forEach(el => { el.style.display = 'none'; });
}

function showResult(type, msg) {
  const el = document.getElementById('resultMsg');
  el.className = type;
  el.textContent = msg;
  el.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
