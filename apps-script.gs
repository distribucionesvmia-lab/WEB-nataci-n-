// Cien por Cien Solidarios: guarda las inscripciones en una hoja de Google.
// Hoja con esta primera fila (A1:H1):
// fecha | nombre | apellidos | email | club | equipo | equipoNombre | estado

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

// Lista pública: solo nombre de pila y equipo. Nunca apellidos ni email.
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

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
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
    if (equipo) {
      const enEquipo = filas.filter(f => Number(f[5]) === equipo).length;
      if (enEquipo >= MAX_EQUIPO) {
        return json_({ ok: false, error: 'Ese equipo ya está completo. Elige otro.' });
      }
    }

    hoja.appendRow([
      new Date(), nombre, apellidos, email,
      String(d.club || '').slice(0, 80),
      equipo,
      String(d.equipoNombre || '').slice(0, 40),
      'Pendiente de pago'
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'No se pudo guardar. Inténtalo de nuevo.' });
  } finally {
    lock.releaseLock();
  }
}
