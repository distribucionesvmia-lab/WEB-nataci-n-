// Cien por Cien Solidarios: guarda las inscripciones en una hoja de Google.
// Hoja con esta primera fila (A1:H1):
// fecha | nombre | apellidos | email | club | equipo | equipoNombre | estado | telefono

const MAX_PARTICIPANTES = 100;
const MAX_EQUIPO = 10;

// Pega aquí el ID de tu hoja: está en su dirección, entre /d/ y /edit
const SHEET_ID = 'PEGA_AQUI_EL_ID_DE_TU_HOJA';

function hoja_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Lista pública: solo nombre de pila y equipo. Nunca apellidos, email ni teléfono.
function doGet() {
  const filas = hoja_().getDataRange().getValues().slice(1);
  const lista = filas
    .filter(f => f[1])
    .map(f => ({
      nombre: String(f[1]).split(' ')[0],
      equipo: f[5] ? Number(f[5]) : null,
      equipoNombre: f[6] ? String(f[6]) : ''
    }));
  return json_(lista);
}

// Comprueba un equipo. Devuelve el texto del error o '' si todo está bien.
function validarEquipo_(filas, equipo, equipoNombre, telefono) {
  const delEquipo = filas.filter(f => Number(f[5]) === equipo);
  if (!delEquipo.length) {
    // Equipo nuevo: el nombre y el teléfono son obligatorios y el nombre no puede repetirse.
    if (!equipoNombre) return 'Ponle un nombre a tu equipo para crearlo.';
    if (filas.some(f => String(f[6]).trim().toLowerCase() === equipoNombre.toLowerCase())) {
      return 'Ya existe un equipo con ese nombre. Elige otro.';
    }
    if (!/^\+?\d{9,15}$/.test(telefono)) return 'Pon un teléfono de contacto válido para el equipo.';
  }
  if (delEquipo.length >= MAX_EQUIPO) return 'Ese equipo ya está completo. Elige otro.';
  return '';
}

// Alguien que se apuntó sin equipo se une a uno (o crea el suyo) desde la lista.
// Se identifica con el email con el que se inscribió.
function cambiarEquipo_(d) {
  const email = String(d.email || '').trim().toLowerCase();
  const equipo = Number(d.equipo) || 0;
  const equipoNombre = String(d.equipoNombre || '').trim().slice(0, 30);
  const telefono = String(d.telefono || '').replace(/[\s.\-()]/g, '').slice(0, 16);
  if (!email || !equipo) return { ok: false, error: 'Faltan datos.' };

  const hoja = hoja_();
  const datos = hoja.getDataRange().getValues();
  const i = datos.findIndex((f, n) => n > 0 && String(f[3]).trim().toLowerCase() === email);
  if (i < 1) return { ok: false, error: 'No encontramos ninguna inscripción con ese email.' };
  if (datos[i][5]) return { ok: false, error: 'Ya estás en un equipo.' };

  const otras = datos.slice(1).filter((f, n) => f[1] && n + 1 !== i);
  const error = validarEquipo_(otras, equipo, equipoNombre, telefono);
  if (error) return { ok: false, error: error };

  const nuevo = !otras.some(f => Number(f[5]) === equipo);
  hoja.getRange(i + 1, 6, 1, 2).setValues([[equipo, nuevo ? equipoNombre : '']]);
  if (nuevo) hoja.getRange(i + 1, 9).setValue("'" + telefono);
  return { ok: true, nombre: String(datos[i][1]).split(' ')[0] };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.accion === 'cambiarEquipo') return json_(cambiarEquipo_(d));
    const nombre = String(d.nombre || '').trim().slice(0, 60);
    const apellidos = String(d.apellidos || '').trim().slice(0, 100);
    const email = String(d.email || '').trim().toLowerCase().slice(0, 120);
    if (!nombre || !apellidos || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ ok: false, error: 'Revisa nombre, apellidos y email.' });
    }

    const hoja = hoja_();
    const filas = hoja.getDataRange().getValues().slice(1).filter(f => f[1]);

    if (filas.length >= MAX_PARTICIPANTES) {
      return json_({ ok: false, error: 'Las plazas están completas.' });
    }
    if (filas.some(f => String(f[3]).toLowerCase() === email)) {
      return json_({ ok: false, error: 'Ya existe una inscripción con ese email.' });
    }

    const equipo = d.equipo ? Number(d.equipo) : '';
    const equipoNombre = String(d.equipoNombre || '').trim().slice(0, 30);
    const telefono = String(d.telefono || '').replace(/[\s.\-()]/g, '').slice(0, 16);
    if (equipo) {
      const error = validarEquipo_(filas, equipo, equipoNombre, telefono);
      if (error) return json_({ ok: false, error: error });
    }

    hoja.appendRow([
      new Date(), nombre, apellidos, email,
      String(d.club || '').slice(0, 80),
      equipo,
      equipoNombre,
      'Pendiente de pago',
      telefono ? "'" + telefono : ''
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'No se pudo guardar. Inténtalo de nuevo.' });
  } finally {
    lock.releaseLock();
  }
}
