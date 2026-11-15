/**
 * VACACIONES 2027 — elección anónima de días
 * Cada persona elige su nombre, crea su PIN la primera vez, elige sus días y ve los días ya cogidos
 * como "Ocupado", sin saber de quién son. Solo tú (en la hoja o con el código
 * de administración) ves los nombres.
 *
 * PASO 1: ejecuta la función "preparar" una vez (crea las pestañas).
 * PASO 2: Implementar > Nueva implementación > Aplicación web
 *         Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 */

const ANIO = 2027;
const DIAS_POR_PERSONA = 23;
// Meses en los que la empresa no deja coger vacaciones (1 = enero … 12 = diciembre)
const MESES_BLOQUEADOS = [1, 2, 11, 12];
// Reparto por semestre: lo normal es 13 + 10 o 10 + 13, con 1 de margen → máximo 14 en un semestre
const MAX_SEMESTRE = 14;
const OBJETIVO_SEMESTRE = '10 o 13';
// Máximo de semanas con festivo que puede coger una misma persona
const MAX_SEMANAS_FESTIVO = 2;

const HOJA_PERSONAS = 'Personas';
const HOJA_RESERVAS = 'Reservas';
const HOJA_FESTIVOS = 'Festivos';
const HOJA_AJUSTES = 'Ajustes';

/* ---------- Preparación (ejecutar una vez) ---------- */
function preparar() {
  const ss = SpreadsheetApp.getActive();

  let p = ss.getSheetByName(HOJA_PERSONAS) || ss.insertSheet(HOJA_PERSONAS);
  if (p.getLastRow() < 2) {
    p.clear();
    const filas = [];
    for (let i = 1; i <= 7; i++) filas.push(['Persona ' + i, '', DIAS_POR_PERSONA]);
    p.getRange(2, 1, filas.length, 3).setValues(filas);
  }
  // Columna B = PIN que crea cada persona al entrar por primera vez (vacía al principio)
  if (String(p.getRange(1, 2, 1, 1).getValues()[0][0]) !== 'PIN (lo crea cada persona)') {
    const n = Math.max(1, p.getLastRow() - 1);
    p.getRange(2, 2, n, 1).setValues(Array.from({ length: n }, () => ['']));
  }
  p.getRange(1, 1, 1, 3).setValues([['Nombre', 'PIN (lo crea cada persona)', 'Días']]).setFontWeight('bold');
  p.getRange('B:B').setNumberFormat('@');
  p.setColumnWidths(1, 3, 170);

  let r = ss.getSheetByName(HOJA_RESERVAS) || ss.insertSheet(HOJA_RESERVAS);
  if (r.getLastRow() < 1) {
    r.getRange(1, 1, 1, 3).setValues([['Fecha', 'Nombre', 'Cuándo lo eligió']]).setFontWeight('bold');
    r.getRange('A:A').setNumberFormat('@');
    r.setColumnWidths(1, 3, 150);
  }

  let f = ss.getSheetByName(HOJA_FESTIVOS) || ss.insertSheet(HOJA_FESTIVOS);
  f.getRange('A:A').setNumberFormat('@');
  if (f.getLastRow() < 1) f.getRange(1, 1, 1, 2).setValues([['Fecha (AAAA-MM-DD)', 'Festivo']]);
  f.getRange(1, 1, 1, 2).setFontWeight('bold');
  f.setColumnWidths(1, 2, 160);
  const fest = [
    ['2027-01-01', 'Año Nuevo'], ['2027-01-06', 'Reyes'], ['2027-03-19', 'San José'],
    ['2027-03-26', 'Viernes Santo'], ['2027-03-29', 'Lunes de Pascua'], ['2027-08-05', 'Festivo Calpe'],
    ['2027-10-12', 'Fiesta Nacional'], ['2027-10-22', 'Festivo Calpe'], ['2027-11-01', 'Todos los Santos'],
    ['2027-12-06', 'Constitución'], ['2027-12-08', 'Inmaculada']
  ];
  // Añade solo los que falten (no borra los que ya tengas)
  const yaHay = f.getDataRange().getValues().map(r => fechaClave_(r[0]));
  const faltan = fest.filter(r => yaHay.indexOf(r[0]) < 0);
  if (faltan.length) f.getRange(f.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);

  let a = ss.getSheetByName(HOJA_AJUSTES) || ss.insertSheet(HOJA_AJUSTES);
  if (a.getLastRow() < 2) {
    a.clear();
    a.getRange('B:B').setNumberFormat('@');
    a.getRange(1, 1, 3, 2).setValues([
      ['Código de administración', 'ADMIN-' + codigoAleatorio()],
      ['Permitir elegir / cambiar días', 'SI'],
      ['Nota', 'Pon NO en la celda B2 para cerrar la elección.']
    ]);
    a.getRange('A1:A3').setFontWeight('bold');
    a.setColumnWidths(1, 2, 240);
  }
  // Si el código de administración está vacío (por ejemplo, en una copia), se crea uno nuevo
  if (!String(a.getRange(1, 2, 1, 1).getValues()[0][0]).trim()) a.getRange(1, 2, 1, 1).setValues([['ADMIN-' + codigoAleatorio()]]);
  // Correo donde llegan los avisos cada vez que alguien guarda (por defecto, el tuyo)
  if (!String(a.getRange(4, 1, 1, 1).getValues()[0][0]).trim()) {
    let correo = '';
    try { correo = Session.getEffectiveUser().getEmail(); } catch (e) {}
    a.getRange(4, 1, 2, 2).setValues([
      ['Correo para avisos', correo],
      ['Nota avisos', 'Cada vez que alguien guarde te llega un correo aquí. Déjalo vacío para no recibir avisos.']
    ]);
    a.getRange(4, 1, 2, 1).setFontWeight('bold');
  }

  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
  ss.setActiveSheet(p);
}

function codigoAleatorio() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += c[Math.floor(Math.random() * c.length)];
  return s;
}

/* ---------- Web ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Vacaciones ' + ANIO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Lectura de datos ---------- */
function leerTodo_() {
  const ss = SpreadsheetApp.getActive();
  const personas = ss.getSheetByName(HOJA_PERSONAS).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ nombre: String(r[0]).trim(), pin: String(r[1]).trim(), max: Number(r[2]) || DIAS_POR_PERSONA }));

  const reservas = {};
  const rs = ss.getSheetByName(HOJA_RESERVAS);
  rs.getDataRange().getValues().slice(1).forEach(r => {
    const k = fechaClave_(r[0]);
    if (k) reservas[k] = String(r[1]).trim();
  });

  const festivos = {};
  ss.getSheetByName(HOJA_FESTIVOS).getDataRange().getValues().slice(1).forEach(r => {
    const k = fechaClave_(r[0]);
    if (k) festivos[k] = String(r[1] || 'Festivo');
  });

  const aj = ss.getSheetByName(HOJA_AJUSTES).getRange(1, 2, 4, 1).getValues();
  const correoAvisos = String(aj[3][0] || '').trim();
  const admin = String(aj[0][0]).trim().toUpperCase();
  const abierto = String(aj[1][0]).trim().toUpperCase() !== 'NO';

  return { personas, reservas, festivos, admin, abierto, correoAvisos };
}

function fechaClave_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
}

function mesBloqueado_(k) {
  return MESES_BLOQUEADOS.indexOf(Number(k.slice(5, 7))) >= 0;
}
function semestre_(k) { return Number(k.slice(5, 7)) <= 6 ? 1 : 2; }
function porSemestre_(lista) {
  const s = [0, 0];
  lista.forEach(k => s[semestre_(k) - 1]++);
  return s;
}

function esLaborable_(k, festivos) {
  const [y, m, d] = k.split('-').map(Number);
  const dow = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  return y === ANIO && dow >= 1 && dow <= 5 && !festivos[k];
}

/** Lista de nombres para la pantalla de entrada (sin PIN ni días). */
function nombres() {
  const t = leerTodo_();
  return t.personas.map(p => ({ nombre: p.nombre, tienePin: !!p.pin }));
}

/** Comprueba quién entra. cred = {admin:'...'} o {nombre:'...', pin:'1234'}.
 *  Si la persona aún no tiene PIN, el que escriba ahora queda guardado como suyo. */
function identificar_(t, cred) {
  cred = cred || {};
  if (cred.admin) {
    if (String(cred.admin).trim().toUpperCase() === t.admin) return { admin: true };
    throw new Error('Código de administración no válido.');
  }
  const pin = String(cred.pin || '').trim();
  const yo = t.personas.find(p => p.nombre === String(cred.nombre || '').trim());
  if (!yo) throw new Error('Elige tu nombre de la lista.');
  if (!/^\d{4}$/.test(pin)) throw new Error('El PIN tiene que ser de 4 números.');
  if (!yo.pin) {
    const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_PERSONAS);
    const col = sh.getDataRange().getValues().map(r => String(r[0]).trim());
    const fila = col.indexOf(yo.nombre);
    sh.getRange(fila + 1, 2, 1, 1).setValues([[pin]]);
    yo.pin = pin;
    return yo;
  }
  if (yo.pin !== pin) throw new Error('PIN incorrecto. Si no lo recuerdas, pídele a tu encargada que lo borre.');
  return yo;
}

/** Lo que ve cada persona: sus días, y los de las demás solo como "ocupado". */
function vista_(t, quien) {
  if (quien.admin) {
    const usados = {}, diasDe = {};
    Object.keys(t.reservas).forEach(k => {
      if (!esLaborable_(k, t.festivos)) return;
      const n = t.reservas[k];
      usados[n] = usados[n] || [0, 0];
      usados[n][semestre_(k) - 1]++;
      (diasDe[n] = diasDe[n] || []).push(k);
    });
    return {
      admin: true, abierto: t.abierto, anio: ANIO, festivos: t.festivos, bloqueados: MESES_BLOQUEADOS,
      reservas: t.reservas, maxSemestre: MAX_SEMESTRE, objetivo: OBJETIVO_SEMESTRE,
      personas: t.personas.map(p => {
        const u = usados[p.nombre] || [0, 0];
        const rv = revisarReglas(diasDe[p.nombre] || [], t.festivos, p.max, MAX_SEMANAS_FESTIVO);
        return { nombre: p.nombre, usados: u[0] + u[1], s1: u[0], s2: u[1], max: p.max, reglas: rv.ok, reglasTxt: rv.items.filter(i => !i.ok).map(i => i.txt).join(' · ') };
      })
    };
  }
  const yo = quien;
  const mios = [], ocupados = [];
  Object.keys(t.reservas).forEach(k => {
    if (!esLaborable_(k, t.festivos)) return;
    (t.reservas[k] === yo.nombre ? mios : ocupados).push(k);
  });
  return {
    admin: false, abierto: t.abierto, anio: ANIO, nombre: yo.nombre, max: yo.max, mios, ocupados, festivos: t.festivos,
    bloqueados: MESES_BLOQUEADOS, maxSemestre: MAX_SEMESTRE, objetivo: OBJETIVO_SEMESTRE, maxFest: MAX_SEMANAS_FESTIVO
  };
}

function entrar(cred) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = leerTodo_();
    return vista_(t, identificar_(t, cred));
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Guardar (con bloqueo para que dos personas no cojan el mismo día) ---------- */
function guardar(cred, anadir, quitar) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = leerTodo_();
    const yo = identificar_(t, cred);
    if (yo.admin) throw new Error('Desde la vista de administración no se eligen días.');
    if (!t.abierto) throw new Error('La elección de vacaciones está cerrada. Habla con tu encargada.');

    anadir = (anadir || []).filter(k => /^\d{4}-\d{2}-\d{2}$/.test(k));
    quitar = (quitar || []).filter(k => /^\d{4}-\d{2}-\d{2}$/.test(k));

    const mios = Object.keys(t.reservas).filter(k => t.reservas[k] === yo.nombre && esLaborable_(k, t.festivos));
    const quedan = mios.filter(k => quitar.indexOf(k) < 0);
    const nuevos = [], ocupados = [];
    anadir.forEach(k => {
      if (!esLaborable_(k, t.festivos) || mesBloqueado_(k)) return;
      if (t.reservas[k] && t.reservas[k] !== yo.nombre) ocupados.push(k);
      else if (!t.reservas[k]) nuevos.push(k);
    });
    if (ocupados.length) {
      return { ok: false, aviso: 'Alguien acaba de coger ' + ocupados.length + (ocupados.length === 1 ? ' día' : ' días') + ' de los que marcaste. No se ha guardado nada; revisa el calendario.', vista: vista_(t, yo) };
    }
    if (quedan.length + nuevos.length > yo.max) {
      return { ok: false, aviso: 'Te pasarías de ' + yo.max + ' días (tendrías ' + (quedan.length + nuevos.length) + '). No se ha guardado nada.', vista: vista_(t, yo) };
    }

    const sem = porSemestre_(quedan.concat(nuevos));
    for (let i = 0; i < 2; i++) {
      if (sem[i] > MAX_SEMESTRE) {
        return { ok: false, aviso: 'En el ' + (i === 0 ? '1er' : '2º') + ' semestre tendrías ' + sem[i] + ' días y el máximo es ' + MAX_SEMESTRE + ' (reparto ' + OBJETIVO_SEMESTRE + ' por semestre). No se ha guardado nada.', vista: vista_(t, yo) };
      }
    }

    const final = quedan.concat(nuevos);
    const rf = revisarReglas(final, t.festivos, yo.max, MAX_SEMANAS_FESTIVO);
    if (!rf.okFest) {
      return { ok: false, aviso: 'Tendrías ' + rf.semFest + ' semanas con festivo y el máximo por persona es ' + MAX_SEMANAS_FESTIVO + '. No se ha guardado nada.', vista: vista_(t, yo) };
    }
    if (final.length >= yo.max) {
      const rv = rf;
      if (!rv.ok) {
        return { ok: false, aviso: 'Con estos días no se cumple la regla: ' + rv.items.filter(i => !i.ok).map(i => i.txt).join(' · ') + '. Recuerda: 4 semanas de lunes a viernes + 3 días juntos (pegados a una semana o en otra semana). No se ha guardado nada.', vista: vista_(t, yo) };
      }
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_RESERVAS);
    // Quitar
    if (quitar.length) {
      const datos = sh.getDataRange().getValues();
      for (let i = datos.length - 1; i >= 1; i--) {
        const k = fechaClave_(datos[i][0]);
        if (String(datos[i][1]).trim() === yo.nombre && quitar.indexOf(k) >= 0) sh.deleteRow(i + 1);
      }
    }
    // Añadir
    const ahora = new Date();
    if (nuevos.length) {
      const filas = nuevos.sort().map(k => [k, yo.nombre, ahora]);
      sh.getRange(sh.getLastRow() + 1, 1, filas.length, 3).setValues(filas);
    }
    SpreadsheetApp.flush();
    const t2 = leerTodo_();
    const quitados = mios.filter(k => quitar.indexOf(k) >= 0);
    if (nuevos.length || quitados.length) {
      try { avisar_(t2, yo.nombre, nuevos, quitados); } catch (e) { console.log('No se pudo enviar el aviso: ' + e); }
    }
    const registro = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy 'a las' HH:mm");
    return { ok: true, registro: registro, vista: vista_(t2, t2.personas.find(p => p.nombre === yo.nombre)) };
  } finally {
    lock.releaseLock();
  }
}

/* Regla: 4 semanas completas (lunes a viernes) + los días restantes juntos,
   pegados a una semana o dentro de otra semana. Si una semana tiene festivo,
   ese día se recupera el lunes siguiente (pegado a la semana) o junto a los 3. */
function revisarReglas(dias, festivos, max, maxFest) {
  if (maxFest == null) maxFest = 2;
  const pad = n => String(n).padStart(2, '0');
  const toK = t => t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1) + '-' + pad(t.getUTCDate());
  const toD = k => { const p = k.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])); };
  const lab = k => { const w = toD(k).getUTCDay(); return w >= 1 && w <= 5 && !festivos[k]; };
  const lunes = k => { const t = toD(k); const w = (t.getUTCDay() + 6) % 7; return toK(new Date(t.getTime() - w * 864e5)); };
  const mover = (k, paso) => { let t = toD(k); for (let i = 0; i < 20; i++) { t = new Date(t.getTime() + paso * 864e5); const x = toK(t); if (lab(x)) return x; } return null; };
  const sel = {}; dias.forEach(k => sel[k] = true);
  const lista = Object.keys(sel).sort();
  // Semanas completas
  const semanas = {};
  lista.forEach(k => semanas[lunes(k)] = true);
  const completas = [], enCompleta = {};
  Object.keys(semanas).forEach(l => {
    const lab5 = [];
    let fest = 0;
    for (let i = 0; i < 5; i++) { const x = toK(new Date(toD(l).getTime() + i * 864e5)); if (lab(x)) lab5.push(x); else if (festivos[x]) fest++; }
    if (lab5.length && lab5.every(x => sel[x])) { completas.push({ lunes: l, fest: fest }); lab5.forEach(x => enCompleta[x] = true); }
  });
  // Días sueltos agrupados en tramos seguidos
  const sueltos = lista.filter(k => !enCompleta[k]);
  const tramos = [];
  sueltos.forEach(k => {
    const ult = tramos[tramos.length - 1];
    if (ult && mover(ult[ult.length - 1], 1) === k) ult.push(k); else tramos.push([k]);
  });
  const pegado = t => enCompleta[mover(t[0], -1)] || enCompleta[mover(t[t.length - 1], 1)];
  const libres = tramos.filter(t => !pegado(t));
  const semFest = completas.filter(c => c.fest > 0).length;
  const okSemanas = completas.length >= 4;
  const okSueltos = libres.length <= 1 && libres.every(t => lunes(t[0]) === lunes(t[t.length - 1])) && tramos.length <= 1 + semFest;
  const okFest = semFest <= maxFest;
  const total = lista.length;
  return {
    ok: okSemanas && okSueltos && okFest,
    okFest: okFest,
    semFest: semFest,
    completo: total >= max,
    semanas: completas.length,
    items: [
      { ok: okSemanas, txt: 'Semanas completas de lunes a viernes: ' + completas.length + ' de 4' },
      { ok: okSueltos, txt: sueltos.length ? 'Días sueltos (' + sueltos.length + '): ' + (okSueltos ? 'van juntos, bien' : 'tienen que ir juntos, pegados a una semana o dentro de otra semana') : 'Días sueltos: ninguno todavía' },
      { ok: okFest, txt: 'Semanas con festivo: ' + semFest + ' (máximo ' + maxFest + ')' }
    ]
  };
}

/* ---------- Aviso por correo a la encargada ---------- */
function avisar_(t, nombre, anadidos, quitados) {
  if (!t.correoAvisos || t.correoAvisos.indexOf('@') < 0) return;
  const MS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const f = k => Number(k.slice(8, 10)) + ' ' + MS[Number(k.slice(5, 7)) - 1];
  const tramos = lista => {
    const out = [];
    lista.slice().sort().forEach(k => {
      const u = out[out.length - 1];
      if (u) {
        let d = new Date(u.hasta + 'T12:00:00Z');
        let seguido = true;
        while (true) {
          d = new Date(d.getTime() + 864e5);
          const x = d.toISOString().slice(0, 10);
          if (x === k) break;
          if (esLaborable_(x, t.festivos) || x > k) { seguido = false; break; }
        }
        if (seguido) { u.hasta = k; u.n++; return; }
      }
      out.push({ desde: k, hasta: k, n: 1 });
    });
    return out.map(p => p.desde === p.hasta ? f(p.desde) : f(p.desde) + ' – ' + f(p.hasta) + ' (' + p.n + ' días)').join(', ');
  };
  const suyos = Object.keys(t.reservas).filter(k => t.reservas[k] === nombre && esLaborable_(k, t.festivos));
  const sem = porSemestre_(suyos);
  const max = (t.personas.find(p => p.nombre === nombre) || {}).max || DIAS_POR_PERSONA;
  const rv = revisarReglas(suyos, t.festivos, max, MAX_SEMANAS_FESTIVO);
  const ss = SpreadsheetApp.getActive();
  const lineas = [
    nombre + ' ha guardado cambios en su solicitud de vacaciones ' + ANIO + ' (pendiente de aprobación por RRHH).',
    '',
    anadidos.length ? 'Añadido: ' + tramos(anadidos) : '',
    quitados.length ? 'Quitado: ' + tramos(quitados) : '',
    '',
    'Ahora tiene ' + suyos.length + ' de ' + max + ' días (1er semestre: ' + sem[0] + ' · 2º semestre: ' + sem[1] + ').',
    'Sus días: ' + (suyos.length ? tramos(suyos) : 'ninguno'),
    suyos.length >= max ? (rv.ok ? 'Regla: cumple ✔' : 'Regla: NO cumple ✘ — ' + rv.items.filter(i => !i.ok).map(i => i.txt).join(' · ')) : '',
    '',
    'Hoja: ' + ss.getUrl()
  ].filter((l, i, a) => l !== '' || (a[i - 1] !== '' && i > 0));
  MailApp.sendEmail({
    to: t.correoAvisos,
    subject: 'Vacaciones ' + ANIO + ': ' + nombre + ' ha guardado sus días',
    body: lineas.join('\n'),
    name: 'Vacaciones ' + ANIO
  });
}

/* ---------- Imprimir: pestaña "Imprimir" y PDF ---------- */
const COLORES_FONDO = ['#dbeafe', '#fee2e2', '#dcfce7', '#fef3c7', '#ede9fe', '#cffafe', '#fce7f3', '#e5e7eb', '#ecfccb', '#ccfbf1'];
const COLORES_TEXTO = ['#1d4ed8', '#b91c1c', '#15803d', '#b45309', '#6d28d9', '#0e7490', '#be185d', '#374151', '#4d7c0f', '#0f766e'];
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
const DIAS_CORTOS = ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sá'];

/** Menú en la hoja: Vacaciones > Preparar para imprimir */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Vacaciones')
    .addItem('Preparar hoja para imprimir', 'imprimir')
    .addToUi();
}

function imprimir() {
  const sh = generarHojaImprimir_();
  SpreadsheetApp.getActive().setActiveSheet(sh);
  SpreadsheetApp.getActive().toast('Listo. Archivo > Imprimir (horizontal, ajustar a la página) o Archivo > Descargar > PDF.', 'Vacaciones', 8);
}

function generarHojaImprimir_() {
  const t = leerTodo_();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Imprimir') || ss.insertSheet('Imprimir');
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  const idx = {};
  t.personas.forEach((p, i) => idx[p.nombre] = i);
  const usados = {};

  // Fila 1 título, fila 2 meses, filas 3-33 días
  const filas = 33, cols = 12;
  const val = [], bg = [], fc = [], fw = [];
  for (let r = 0; r < filas; r++) {
    val.push(new Array(cols).fill('')); bg.push(new Array(cols).fill('#ffffff'));
    fc.push(new Array(cols).fill('#1f2328')); fw.push(new Array(cols).fill('normal'));
  }
  val[0][0] = /vacaciones/i.test(ss.getName()) ? ss.getName() : 'Vacaciones ' + ANIO + ' · ' + ss.getName();
  fw[0][0] = 'bold'; fc[0][0] = '#0b4a8b';
  for (let m = 0; m < 12; m++) {
    val[1][m] = MESES[m];
    bg[1][m] = '#374151'; fc[1][m] = '#ffffff'; fw[1][m] = 'bold';
    const dim = new Date(Date.UTC(ANIO, m + 1, 0)).getUTCDate();
    for (let d = 1; d <= 31; d++) {
      const r = d + 1;
      if (d > dim) { bg[r][m] = '#f3f4f6'; continue; }
      const k = ANIO + '-' + String(m + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      const dow = new Date(Date.UTC(ANIO, m, d)).getUTCDay();
      let txt = d + ' ' + DIAS_CORTOS[dow];
      if (dow === 0 || dow === 6) { bg[r][m] = '#eef0f3'; fc[r][m] = '#6b7280'; }
      if (t.festivos[k]) { txt += '  ' + t.festivos[k]; bg[r][m] = '#fde2e2'; fc[r][m] = '#991b1b'; }
      else if (t.reservas[k] && esLaborable_(k, t.festivos)) {
        const n = t.reservas[k], i = (idx[n] != null ? idx[n] : 7) % COLORES_FONDO.length;
        txt += '  ' + n; bg[r][m] = COLORES_FONDO[i]; fc[r][m] = COLORES_TEXTO[i]; fw[r][m] = 'bold';
        usados[n] = usados[n] || [0, 0]; usados[n][semestre_(k) - 1]++;
      } else if (MESES_BLOQUEADOS.indexOf(m + 1) >= 0 && dow !== 0 && dow !== 6) { fc[r][m] = '#9ca3af'; }
      val[r][m] = txt;
    }
  }
  const rg = sh.getRange(1, 1, filas, cols);
  rg.setValues(val).setBackgrounds(bg).setFontColors(fc).setFontWeights(fw)
    .setFontSize(8).setFontFamily('Arial').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(1, 1, 1, cols).merge().setFontSize(16);
  sh.getRange(2, 1, 1, cols).setHorizontalAlignment('center').setFontSize(9);
  sh.getRange(2, 1, filas - 1, cols).setBorder(true, true, true, true, true, true, '#c7ccd4', SpreadsheetApp.BorderStyle.SOLID);

  // Leyenda: nombre · total (1er sem / 2º sem)
  const porFila = 4, ini = filas + 2;
  t.personas.forEach((p, i) => {
    const u = usados[p.nombre] || [0, 0];
    const c = sh.getRange(ini + Math.floor(i / porFila), 1 + (i % porFila) * 3, 1, 3).merge();
    c.setValue(p.nombre + ' · ' + (u[0] + u[1]) + ' días (1er sem: ' + u[0] + ' · 2º sem: ' + u[1] + ')')
      .setBackground(COLORES_FONDO[i % COLORES_FONDO.length]).setFontColor(COLORES_TEXTO[i % COLORES_TEXTO.length])
      .setFontWeight('bold').setFontSize(8);
  });

  sh.setColumnWidths(1, cols, 96);
  sh.setRowHeight(1, 30);
  sh.setRowHeights(2, filas + 2 + Math.ceil(t.personas.length / porFila), 16);
  sh.setHiddenGridlines(true);
  SpreadsheetApp.flush();
  return sh;
}

/** Desde la vista de encargada: crea el PDF en tu Drive y devuelve el enlace. */
function pdfAdmin(cred) {
  const t = leerTodo_();
  const quien = identificar_(t, cred);
  if (!quien.admin) throw new Error('Solo la encargada puede descargar el PDF.');
  const ss = SpreadsheetApp.getActive();
  const sh = generarHojaImprimir_();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + sh.getSheetId() +
    '&size=A4&portrait=false&scale=4&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
    '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3&horizontal_alignment=CENTER';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  const nombre = (/vacaciones/i.test(ss.getName()) ? ss.getName() : 'Vacaciones ' + ANIO + ' - ' + ss.getName()) + '.pdf';
  const viejos = DriveApp.getFilesByName(nombre);
  while (viejos.hasNext()) viejos.next().setTrashed(true);
  const file = DriveApp.createFile(blob.setName(nombre));
  return file.getUrl();
}
