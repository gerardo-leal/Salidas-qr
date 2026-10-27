/**
 * Salidas QR – servidor (API)
 * Escuela Primaria Profesor Gerardo Guadalupe Leal Leal
 *
 * Base de datos: esta hoja de cálculo en Google Drive.
 * Fotos: carpeta "Salidas QR - Fotos" en Drive.
 * La página (index.html) vive en el dominio de la escuela y llama a esta API.
 */

const TZ = 'America/Monterrey';
const HOJAS = {
  Alumnos:  ['ID', 'Nombre', 'CURP', 'Grupo', 'QR', 'FamiliaID', 'Activo', 'FotoID'],
  Familias: ['ID', 'Tutor', 'PIN', 'Telefono', 'Correo'],
  Personas: ['ID', 'FamiliaID', 'Nombre', 'Parentesco', 'FotoID'],
  Maestros: ['ID', 'Nombre', 'Grupo', 'Usuario', 'Contrasena', 'Correo'],
  Salidas:  ['Fecha', 'Hora', 'AlumnoID', 'Alumno', 'Grupo', 'Recogio', 'Parentesco', 'FotoID', 'Entrego'],
  Permisos: ['Fecha', 'Grupo', 'MaestroID'],
  Config:   ['Clave', 'Valor']
};
const TEXTO = {
  Alumnos: ['CURP', 'Grupo'], Familias: ['PIN', 'Telefono'], Maestros: ['Contrasena'],
  Salidas: ['Fecha', 'Hora'], Permisos: ['Fecha'], Config: ['Valor']
};
const GRUPOS = ['1°A','1°B','2°A','2°B','3°A','3°B','4°A','4°B','5°A','5°B','6°A','6°B'];

/* ================= API ================= */

const API = {
  entrar, salir,
  datosPadre, cambiarQR, cambiarPin, cambiarContacto, agregarPersona, quitarPersona,
  buscarQR, registrarSalida, datosMaestro, darPermiso, quitarPermiso,
  resumenDir, listaGrupo, buscarAlumnos, fichaAlumno, altaAlumno, editarAlumno, fotoAlumno,
  editarFamilia, restablecerPin, listaMaestros, guardarMaestro,
  importarFilas, reporte, exportarExcel, cerrarCicloDir, cambiarClaveDir
};

function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    const f = API[b.fn];
    if (!f) throw new Error('Acción no válida.');
    out = { ok: true, data: f.apply(null, b.args || []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'Salidas QR', escuela: cfg('nombre_escuela') }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= INSTALACIÓN ================= */

function instalar() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HOJAS).forEach(nombre => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) sh = ss.insertSheet(nombre);
    const enc = HOJAS[nombre];
    sh.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold').setBackground('#dbefee');
    sh.setFrozenRows(1);
    (TEXTO[nombre] || []).forEach(col => sh.getRange(2, enc.indexOf(col) + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@'));
  });
  ['Hoja 1', 'Sheet1'].forEach(n => { const h = ss.getSheetByName(n); if (h && ss.getSheets().length > 1) ss.deleteSheet(h); });

  if (!cfg('usuario_direccion')) {
    const carpeta = DriveApp.createFolder('Salidas QR - Fotos');
    const clave = String(Math.floor(100000 + Math.random() * 900000));
    [['nombre_escuela', 'Escuela Primaria Profesor Gerardo Guadalupe Leal Leal'],
     ['usuario_direccion', 'direccion'],
     ['contrasena_direccion', clave],
     ['carpeta_fotos', carpeta.getId()]].forEach(r => hoja('Config').appendRow(r));
    avisar('Instalación lista.\n\nUsuario de dirección: direccion\nContraseña: ' + clave +
      '\n\nAnótala. Podrás cambiarla desde la página (Más → Contraseña).');
  }
}

// Muestra el mensaje en el registro y como aviso flotante en la hoja (sin bloquear la ejecución)
function avisar(msg) {
  console.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'Salidas QR', 15); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Salidas QR')
    .addItem('Cargar datos de prueba', 'cargarDatosDePrueba')
    .addItem('Borrar todos los datos', 'borrarTodo')
    .addToUi();
}

/* ================= DATOS DE PRUEBA ================= */

const CLAVE_DIR_PRUEBA = 'direccion2026';
const CLAVE_CARLOS_PRUEBA = 'carlos2026';
const CLAVE_MAESTROS_PRUEBA = 'maestro2026';

// Deja vacías las hojas de datos (no toca Config ni la carpeta de fotos)
function borrarTodo() {
  conBloqueo(() => ['Alumnos', 'Familias', 'Personas', 'Maestros', 'Salidas', 'Permisos'].forEach(n => {
    const h = hoja(n);
    if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, HOJAS[n].length).clearContent();
  }));
  avisar('Listo: se borraron alumnos, familias, personas, maestros y salidas.');
}

// Llena la base con datos ficticios (siempre los mismos) para probar
function cargarDatosDePrueba() {
  if (!cfg('usuario_direccion')) instalar();
  if (filas('Alumnos').length) throw new Error('Ya hay alumnos. Usa primero "Borrar todos los datos".');
  const d = generarDatosDePrueba();
  const poner = (n, objs) => {
    if (!objs.length) return;
    const sh = hoja(n);
    sh.getRange(2, 1, objs.length, HOJAS[n].length).setValues(objs.map(o => HOJAS[n].map(k => o[k] !== undefined ? o[k] : '')));
  };
  conBloqueo(() => {
    poner('Maestros', d.maestros); poner('Familias', d.familias); poner('Alumnos', d.alumnos);
    poner('Personas', d.personas); poner('Salidas', d.salidas);
    setCfg('usuario_direccion', 'direccion');
    setCfg('contrasena_direccion', CLAVE_DIR_PRUEBA);
  });
  avisar('Datos de prueba cargados: ' + d.alumnos.length + ' alumnos, ' + d.familias.length + ' familias, ' +
    d.maestros.length + ' maestros y ' + d.salidas.length + ' salidas.\n\n' +
    'Dirección: direccion / ' + CLAVE_DIR_PRUEBA + '\nMaestro: carlos / ' + CLAVE_CARLOS_PRUEBA +
    '\n\nLas familias de prueba y sus PIN están en la pestaña Familias; el CURP de acceso es el del hijo mayor.');
}

function generarDatosDePrueba() {
  let semilla = 20261027;
  const azar = () => { // mulberry32: siempre genera los mismos datos
    semilla = (semilla + 0x6D2B79F5) | 0;
    let t = Math.imul(semilla ^ (semilla >>> 15), 1 | semilla);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const uno = arr => arr[Math.floor(azar() * arr.length)];
  const NOM_H = ['Mateo','Santiago','Diego','Emiliano','Leonardo','Iker','Gael','Sebastián','Ángel','Julián','Rodrigo','Tadeo','Daniel','Adrián','Joaquín','Matías','Alejandro','Bruno','Thiago','Nicolás'];
  const NOM_M = ['Sofía','Valentina','Regina','Ximena','Camila','Renata','Natalia','Andrea','Fernanda','Mariana','Paola','Daniela','Victoria','Isabella','Romina','Valeria','Lucía','Julieta','Emma','Ana Paula'];
  const APE = ['García','Martínez','Hernández','López','Treviño','Garza','Rodríguez','Villarreal','Cantú','Salinas','Flores','Elizondo','González','Leal','Cavazos','Guerra','Tamez','Montemayor','Benavides','Saldaña','Zambrano','Ramos','Sepúlveda','Chapa'];
  const MAE = [['Jorge Salinas Ríos','1°A'],['Patricia Guerra Leal','1°B'],['Laura Treviño Cantú','2°A'],['Miguel Ángel Ramos','2°B'],
    ['Claudia Villarreal Garza','3°A'],['Roberto Tamez Flores','3°B'],['Carlos Estala','4°A'],['Diana Chapa Rodríguez','4°B'],
    ['Héctor Montemayor Saldaña','5°A'],['Gabriela Elizondo Ruiz','5°B'],['Ricardo Benavides López','6°A'],['Mónica Sepúlveda Zambrano','6°B']];
  const sinAcento = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z ]/g, '');
  const consonante = s => (sinAcento(s).slice(1).match(/[BCDFGHJKLMNPQRSTVWXYZ]/) || ['X'])[0];
  const usados = {};
  function curp(nombre, ap1, ap2, nac, sexo) {
    const n = sinAcento(nombre.split(' ')[0]), a1 = sinAcento(ap1), a2 = sinAcento(ap2);
    const vocal = (a1.slice(1).match(/[AEIOU]/) || ['X'])[0];
    let base = a1[0] + vocal + a2[0] + n[0] + nac + sexo + 'NL' + consonante(a1) + consonante(a2) + consonante(n);
    let c, i = 0;
    do { c = base + 'A' + (i++ % 10); } while (usados[c] && i < 20);
    usados[c] = 1;
    return c;
  }
  const maestros = MAE.map(([n, g], i) => ({ ID: 'M' + String(i + 1).padStart(3, '0'), Nombre: n, Grupo: g,
    Usuario: n === 'Carlos Estala' ? 'carlos' : sinAcento(n.split(' ')[0]).toLowerCase(),
    Contrasena: n === 'Carlos Estala' ? CLAVE_CARLOS_PRUEBA : CLAVE_MAESTROS_PRUEBA, Correo: '' }));
  const familias = [], alumnos = [], personas = [];
  let nf = 0, na = 0, np = 0;
  function nuevaFamilia(ap1, ap2) {
    const id = 'F' + String(++nf).padStart(3, '0');
    const mama = uno(NOM_M) + ' ' + ap2 + ' ' + uno(APE), papa = uno(NOM_H) + ' ' + ap1 + ' ' + uno(APE);
    familias.push({ ID: id, Tutor: mama, PIN: String(1000 + Math.floor(azar() * 9000)), Telefono: '81' + String(10000000 + Math.floor(azar() * 89999999)), Correo: '' });
    personas.push({ ID: 'P' + String(++np).padStart(4, '0'), FamiliaID: id, Nombre: mama, Parentesco: 'Mamá', FotoID: '' });
    personas.push({ ID: 'P' + String(++np).padStart(4, '0'), FamiliaID: id, Nombre: papa, Parentesco: 'Papá', FotoID: '' });
    if (azar() < 0.35) personas.push({ ID: 'P' + String(++np).padStart(4, '0'), FamiliaID: id, Nombre: uno(NOM_M) + ' ' + ap1 + ' ' + uno(APE), Parentesco: 'Abuela', FotoID: '' });
    return id;
  }
  function nuevoAlumno(grupo, famId, ap1, ap2) {
    const sexo = azar() < 0.5 ? 'H' : 'M';
    const nombre = uno(sexo === 'H' ? NOM_H : NOM_M);
    const anio = 2026 - 6 - grado(grupo) + (azar() < 0.7 ? 0 : 1);
    const nac = String(anio).slice(2) + String(1 + Math.floor(azar() * 12)).padStart(2, '0') + String(1 + Math.floor(azar() * 28)).padStart(2, '0');
    const id = 'A' + String(++na).padStart(4, '0');
    alumnos.push({ ID: id, Nombre: nombre + ' ' + ap1 + ' ' + ap2, CURP: curp(nombre, ap1, ap2, nac, sexo), Grupo: grupo,
      QR: 'QR-PRUEBA' + String(na).padStart(4, '0'), FamiliaID: famId, Activo: 'SI', FotoID: '' });
  }
  GRUPOS.forEach(g => {
    for (let i = 0; i < 8; i++) {
      const ap1 = uno(APE), ap2 = uno(APE);
      const fam = nuevaFamilia(ap1, ap2);
      nuevoAlumno(g, fam, ap1, ap2);
      // Uno de cada cuatro tiene un hermano menor en un grado más bajo
      if (i % 4 === 0 && grado(g) > 1) nuevoAlumno(GRUPOS[Math.floor(azar() * (grado(g) - 1)) * 2 + (azar() < 0.5 ? 0 : 1)], fam, ap1, ap2);
    }
  });
  // Historial: últimos 5 días hábiles antes de hoy
  const salidas = [], ahoraMs = Date.now();
  let dias = 0;
  for (let k = 1; dias < 5 && k < 15; k++) {
    const f = new Date(ahoraMs - k * 86400000);
    const dow = Number(Utilities.formatDate(f, TZ, 'u'));
    if (dow > 5) continue;
    dias++;
    const fecha = Utilities.formatDate(f, TZ, 'yyyy-MM-dd');
    alumnos.forEach(a => {
      if (azar() < 0.1) return; // faltó
      const ps = personas.filter(p => p.FamiliaID === a.FamiliaID);
      const p = uno(ps);
      const m = maestros.find(x => x.Grupo === a.Grupo);
      const min = 30 + Math.floor(azar() * 45);
      salidas.push({ Fecha: fecha, Hora: (15 + Math.floor(min / 60)) + ':' + String(min % 60).padStart(2, '0'), AlumnoID: a.ID, Alumno: a.Nombre,
        Grupo: a.Grupo, Recogio: p.Nombre, Parentesco: p.Parentesco, FotoID: '', Entrego: m ? m.Nombre : 'Dirección' });
    });
  }
  salidas.sort((x, y) => (x.Fecha + x.Hora).localeCompare(y.Fecha + y.Hora));
  return { maestros, familias, alumnos, personas, salidas };
}

/* ================= AYUDANTES ================= */

function hoja(n) { return SpreadsheetApp.getActive().getSheetByName(n); }

function filas(n) {
  const v = hoja(n).getDataRange().getValues();
  const enc = v.shift();
  return v.map((r, i) => {
    const o = { _fila: i + 2 };
    enc.forEach((k, j) => {
      let x = r[j];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, k === 'Hora' ? 'HH:mm' : 'yyyy-MM-dd');
      o[k] = (x === null || x === undefined) ? '' : String(x).trim();
    });
    return o;
  }).filter(o => HOJAS[n].some(k => o[k] !== ''));
}

function agregar(n, obj) {
  const fila = HOJAS[n].map(k => obj[k] !== undefined ? obj[k] : '');
  const sh = hoja(n);
  sh.getRange(sh.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
}

function actualizar(n, numFila, obj) {
  const sh = hoja(n);
  Object.keys(obj).forEach(k => {
    const c = HOJAS[n].indexOf(k);
    if (c >= 0) sh.getRange(numFila, c + 1).setValue(obj[k]);
  });
}

function cfg(k) {
  const sh = hoja('Config');
  if (!sh) return '';
  const r = sh.getDataRange().getValues().find(x => x[0] === k);
  return r ? String(r[1]) : '';
}

function setCfg(k, v) {
  const f = filas('Config').find(x => x.Clave === k);
  if (f) actualizar('Config', f._fila, { Valor: v }); else agregar('Config', { Clave: k, Valor: v });
}

const nuevoId = () => Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
const nuevoQR = () => 'QR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
const nuevoPin = () => String(Math.floor(1000 + Math.random() * 9000));
const hoy = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const ahora = () => Utilities.formatDate(new Date(), TZ, 'HH:mm');
const grado = g => parseInt(g, 10) || 0;
const txt = x => String(x == null ? '' : x).trim();

function normalizarGrupo(g) {
  const m = txt(g).toUpperCase().replace(/\s/g, '').match(/^([1-6])\D*([AB])$/);
  return m ? m[1] + '°' + m[2] : '';
}

function conBloqueo(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function fotoDataUrl(fotoId) {
  if (!fotoId) return '';
  try {
    const b = DriveApp.getFileById(fotoId).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

function guardarFoto(dataUrl, nombre) {
  const m = txt(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre + '.jpg');
  return DriveApp.getFolderById(cfg('carpeta_fotos')).createFile(blob).getId();
}

function borrarFoto(id) { if (id) try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }

function mayorActivo(familiaId, alumnos) {
  return alumnos.filter(a => a.FamiliaID === familiaId && a.Activo !== 'NO')
    .sort((a, b) => grado(b.Grupo) - grado(a.Grupo))[0];
}

const personasDe = famId => filas('Personas').filter(p => p.FamiliaID === famId)
  .map(p => ({ id: p.ID, nombre: p.Nombre, parentesco: p.Parentesco, foto: fotoDataUrl(p.FotoID) }));

const salidaCorta = x => ({ fecha: x.Fecha, hora: x.Hora, alumno: x.Alumno, grupo: x.Grupo, recogio: x.Recogio, parentesco: x.Parentesco, entrego: x.Entrego });

/* ================= SESIONES ================= */

function crearSesion(datos) {
  const t = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, JSON.stringify(datos), 21600); // 6 horas
  return t;
}

function sesion(token, roles) {
  const c = token && CacheService.getScriptCache().get('s_' + token);
  if (!c) throw new Error('Tu sesión terminó. Vuelve a entrar.');
  const s = JSON.parse(c);
  if (roles && roles.indexOf(s.rol) < 0) throw new Error('Tu usuario no tiene permiso para esto.');
  return s;
}

function salir(token) { CacheService.getScriptCache().remove('s_' + token); return true; }

/* ================= LOGIN ================= */

function entrar(rol, usuario, clave) {
  usuario = txt(usuario); clave = txt(clave);
  if (rol === 'padre') {
    const curp = usuario.toUpperCase();
    const alumnos = filas('Alumnos');
    const a = alumnos.find(x => x.CURP.toUpperCase() === curp && x.Activo !== 'NO');
    if (!a) {
      const baja = alumnos.find(x => x.CURP.toUpperCase() === curp);
      const sigue = baja && mayorActivo(baja.FamiliaID, alumnos);
      if (sigue) throw new Error('Ese alumno ya egresó. Ahora entra con el CURP de ' + sigue.Nombre.split(' ')[0] + '.');
      throw new Error('No encontramos ese CURP. Revisa que esté bien escrito.');
    }
    const mayor = mayorActivo(a.FamiliaID, alumnos);
    if (mayor && mayor.ID !== a.ID) throw new Error('Entra con el CURP de tu hijo(a) mayor inscrito: ' + mayor.Nombre.split(' ')[0] + '.');
    const fam = filas('Familias').find(f => f.ID === a.FamiliaID);
    if (!fam || fam.PIN !== clave) throw new Error('PIN incorrecto. Si lo olvidaste, pídelo en la dirección.');
    return { token: crearSesion({ rol: 'padre', fam: fam.ID }), rol: 'padre', nombre: fam.Tutor };
  }
  if (rol === 'maestro') {
    const m = filas('Maestros').find(x => x.Usuario.toLowerCase() === usuario.toLowerCase());
    if (!m || m.Contrasena !== clave) throw new Error('Usuario o contraseña incorrectos.');
    return { token: crearSesion({ rol: 'maestro', id: m.ID }), rol: 'maestro', nombre: m.Nombre + ' · ' + m.Grupo };
  }
  if (rol === 'dir') {
    if (usuario.toLowerCase() !== cfg('usuario_direccion').toLowerCase() || clave !== cfg('contrasena_direccion'))
      throw new Error('Usuario o contraseña incorrectos.');
    return { token: crearSesion({ rol: 'dir' }), rol: 'dir', nombre: 'Dirección' };
  }
  throw new Error('Tipo de usuario no válido.');
}

/* ================= PADRE ================= */

function datosPadre(token) {
  const s = sesion(token, ['padre']);
  const fam = filas('Familias').find(f => f.ID === s.fam);
  const maestros = filas('Maestros');
  const hijos = filas('Alumnos').filter(a => a.FamiliaID === s.fam && a.Activo !== 'NO')
    .sort((a, b) => grado(b.Grupo) - grado(a.Grupo))
    .map(a => ({ id: a.ID, nombre: a.Nombre, grupo: a.Grupo, qr: a.QR,
      maestro: (maestros.find(m => m.Grupo === a.Grupo) || {}).Nombre || 'Por asignar' }));
  const ids = hijos.map(h => h.id);
  return {
    tutor: fam.Tutor, telefono: fam.Telefono, correo: fam.Correo, hijos,
    personas: personasDe(s.fam),
    historial: filas('Salidas').filter(x => ids.indexOf(x.AlumnoID) >= 0).slice(-60).reverse().map(salidaCorta)
  };
}

function cambiarQR(token, alumnoId) {
  const s = sesion(token, ['padre', 'dir']);
  return conBloqueo(() => {
    const a = filas('Alumnos').find(x => x.ID === alumnoId);
    if (!a || (s.rol === 'padre' && a.FamiliaID !== s.fam)) throw new Error('Alumno no encontrado.');
    const qr = nuevoQR();
    actualizar('Alumnos', a._fila, { QR: qr });
    return qr;
  });
}

function cambiarPin(token, pin) {
  const s = sesion(token, ['padre']);
  pin = txt(pin);
  if (!/^\d{4,6}$/.test(pin)) throw new Error('El PIN debe tener de 4 a 6 números.');
  const f = filas('Familias').find(x => x.ID === s.fam);
  actualizar('Familias', f._fila, { PIN: pin });
  return true;
}

function validarCorreo(c) {
  if (c && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c)) throw new Error('El correo no parece válido: ' + c);
  return c;
}

function cambiarContacto(token, d) {
  const s = sesion(token, ['padre']);
  const f = filas('Familias').find(x => x.ID === s.fam);
  actualizar('Familias', f._fila, { Correo: validarCorreo(txt(d.correo)), Telefono: txt(d.telefono) });
  return true;
}

// Padre: su familia. Dirección: indica p.famId.
function agregarPersona(token, p) {
  const s = sesion(token, ['padre', 'dir']);
  const famId = s.rol === 'padre' ? s.fam : txt(p.famId);
  const nombre = txt(p.nombre);
  if (!famId) throw new Error('Falta la familia.');
  if (!nombre) throw new Error('Escribe el nombre completo.');
  if (!p.foto) throw new Error('Agrega una foto del rostro para que el maestro la reconozca.');
  return conBloqueo(() => {
    agregar('Personas', { ID: nuevoId(), FamiliaID: famId, Nombre: nombre, Parentesco: txt(p.parentesco), FotoID: guardarFoto(p.foto, nombre) });
    return true;
  });
}

function quitarPersona(token, personaId) {
  const s = sesion(token, ['padre', 'dir']);
  return conBloqueo(() => {
    const todas = filas('Personas');
    const p = todas.find(x => x.ID === personaId && (s.rol === 'dir' || x.FamiliaID === s.fam));
    if (!p) throw new Error('Persona no encontrada.');
    if (s.rol === 'padre' && todas.filter(x => x.FamiliaID === s.fam).length <= 1) throw new Error('Debe quedar al menos una persona autorizada.');
    hoja('Personas').deleteRow(p._fila);
    borrarFoto(p.FotoID);
    return true;
  });
}

/* ================= MAESTRO ================= */

function permisoHoy(maestroId, grupo) {
  return filas('Permisos').some(p => p.Fecha === hoy() && p.Grupo === grupo && p.MaestroID === maestroId);
}

function buscarQR(token, codigo) {
  const s = sesion(token, ['maestro', 'dir']);
  codigo = txt(codigo).toUpperCase();
  const a = filas('Alumnos').find(x => x.QR.toUpperCase() === codigo && x.Activo !== 'NO');
  if (!a) return { valido: false };
  const maestros = filas('Maestros');
  const yo = s.rol === 'maestro' ? maestros.find(m => m.ID === s.id) : null;
  const titular = maestros.find(m => m.Grupo === a.Grupo);
  let permiso = 'dir';
  if (yo) permiso = yo.Grupo === a.Grupo ? 'propio' : (permisoHoy(yo.ID, a.Grupo) ? 'autorizado' : 'no');
  const salida = filas('Salidas').filter(x => x.AlumnoID === a.ID && x.Fecha === hoy()).pop();
  return {
    valido: true,
    alumno: { id: a.ID, nombre: a.Nombre, grupo: a.Grupo, qr: a.QR, foto: fotoDataUrl(a.FotoID), titular: titular ? titular.Nombre : 'Sin maestro' },
    permiso,
    salida: salida ? { hora: salida.Hora, recogio: salida.Recogio } : null,
    personas: permiso === 'no' ? [] : personasDe(a.FamiliaID)
  };
}

function registrarSalida(token, alumnoId, personaId) {
  const s = sesion(token, ['maestro', 'dir']);
  const res = conBloqueo(() => {
    const a = filas('Alumnos').find(x => x.ID === alumnoId);
    if (!a) throw new Error('Alumno no encontrado.');
    const yo = s.rol === 'maestro' ? filas('Maestros').find(m => m.ID === s.id) : null;
    if (yo && yo.Grupo !== a.Grupo && !permisoHoy(yo.ID, a.Grupo))
      throw new Error('No tienes permiso para entregar a alumnos de ' + a.Grupo + '.');
    if (filas('Salidas').some(x => x.AlumnoID === a.ID && x.Fecha === hoy()))
      throw new Error('Este alumno ya tiene salida registrada hoy.');
    const p = filas('Personas').find(x => x.ID === personaId && x.FamiliaID === a.FamiliaID);
    if (!p) throw new Error('Esa persona no está autorizada para este alumno.');
    const reg = { Fecha: hoy(), Hora: ahora(), AlumnoID: a.ID, Alumno: a.Nombre, Grupo: a.Grupo,
      Recogio: p.Nombre, Parentesco: p.Parentesco, FotoID: p.FotoID, Entrego: yo ? yo.Nombre : 'Dirección' };
    agregar('Salidas', reg);
    return { reg, fam: filas('Familias').find(f => f.ID === a.FamiliaID) };
  });
  let correo = 'sin correo registrado';
  if (res.fam && res.fam.Correo) {
    try {
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error('cuota');
      const r = res.reg;
      MailApp.sendEmail({
        to: res.fam.Correo,
        name: cfg('nombre_escuela') || 'Escuela',
        subject: 'Salida de ' + r.Alumno.split(' ')[0] + ' – ' + r.Hora,
        htmlBody: '<p>Hola, ' + res.fam.Tutor + ':</p><p><b>' + r.Alumno + '</b> (' + r.Grupo + ') salió de la escuela hoy ' +
          r.Fecha + ' a las <b>' + r.Hora + '</b> con <b>' + r.Recogio + '</b> (' + r.Parentesco + ').</p><p>Lo entregó: ' +
          r.Entrego + '.</p><p style="color:#666">Si no reconoces esta salida, comunícate de inmediato con la escuela.</p>'
      });
      correo = 'enviado';
    } catch (e) { correo = 'no se pudo enviar'; }
  }
  return { hora: res.reg.Hora, correo };
}

function datosMaestro(token) {
  const s = sesion(token, ['maestro']);
  const maestros = filas('Maestros');
  const yo = maestros.find(m => m.ID === s.id);
  if (!yo) throw new Error('Tu usuario ya no existe. Habla con la dirección.');
  const salidasHoy = filas('Salidas').filter(x => x.Fecha === hoy());
  const permisos = filas('Permisos').filter(p => p.Fecha === hoy());
  const grupo = filas('Alumnos').filter(a => a.Grupo === yo.Grupo && a.Activo !== 'NO')
    .sort((a, b) => a.Nombre.localeCompare(b.Nombre))
    .map(a => { const x = salidasHoy.find(y => y.AlumnoID === a.ID); return { id: a.ID, nombre: a.Nombre, qr: a.QR, salio: x ? x.Hora + ' · ' + x.Recogio : '' }; });
  return {
    yo: { nombre: yo.Nombre, grupo: yo.Grupo },
    grupo,
    maestros: maestros.filter(m => m.ID !== yo.ID).map(m => ({ id: m.ID, nombre: m.Nombre, grupo: m.Grupo })),
    dados: permisos.filter(p => p.Grupo === yo.Grupo).map(p => ({ id: p.MaestroID, nombre: (maestros.find(m => m.ID === p.MaestroID) || {}).Nombre || '' })),
    recibidos: permisos.filter(p => p.MaestroID === yo.ID).map(p => ({ grupo: p.Grupo, de: (maestros.find(m => m.Grupo === p.Grupo) || {}).Nombre || '' })),
    salidas: salidasHoy.slice().reverse().map(salidaCorta)
  };
}

function darPermiso(token, maestroId) {
  const s = sesion(token, ['maestro']);
  return conBloqueo(() => {
    const yo = filas('Maestros').find(m => m.ID === s.id);
    if (!permisoHoy(maestroId, yo.Grupo)) agregar('Permisos', { Fecha: hoy(), Grupo: yo.Grupo, MaestroID: maestroId });
    return true;
  });
}

function quitarPermiso(token, maestroId) {
  const s = sesion(token, ['maestro']);
  return conBloqueo(() => {
    const yo = filas('Maestros').find(m => m.ID === s.id);
    filas('Permisos').filter(p => p.Fecha === hoy() && p.Grupo === yo.Grupo && p.MaestroID === maestroId)
      .reverse().forEach(p => hoja('Permisos').deleteRow(p._fila));
    return true;
  });
}

/* ================= DIRECCIÓN ================= */

function resumenDir(token) {
  sesion(token, ['dir']);
  const alumnos = filas('Alumnos').filter(a => a.Activo !== 'NO');
  const maestros = filas('Maestros');
  const salidasHoy = filas('Salidas').filter(x => x.Fecha === hoy());
  const conteo = {};
  GRUPOS.forEach(g => conteo[g] = { alumnos: alumnos.filter(a => a.Grupo === g).length, salidas: salidasHoy.filter(x => x.Grupo === g).length,
    maestro: (maestros.find(m => m.Grupo === g) || {}).Nombre || '' });
  return {
    grupos: GRUPOS, conteo, total: alumnos.length, salidasHoy: salidasHoy.length, maestros: maestros.length,
    ultimas: salidasHoy.slice(-8).reverse().map(salidaCorta),
    sinPersonas: (() => { const conP = {}; filas('Personas').forEach(p => conP[p.FamiliaID] = 1); return alumnos.filter(a => !conP[a.FamiliaID]).length; })(),
    correosRestantes: MailApp.getRemainingDailyQuota(),
    hojaUrl: SpreadsheetApp.getActive().getUrl()
  };
}

function listaGrupo(token, grupo) {
  sesion(token, ['dir']);
  const familias = filas('Familias');
  return filas('Alumnos').filter(a => a.Grupo === grupo && a.Activo !== 'NO')
    .sort((a, b) => a.Nombre.localeCompare(b.Nombre))
    .map(a => { const f = familias.find(x => x.ID === a.FamiliaID) || {}; return { id: a.ID, nombre: a.Nombre, grupo: a.Grupo, tutor: f.Tutor || '', pin: f.PIN || '' }; });
}

function buscarAlumnos(token, texto) {
  sesion(token, ['dir']);
  const q = txt(texto).toLowerCase();
  if (q.length < 2) return [];
  return filas('Alumnos').filter(a => a.Nombre.toLowerCase().indexOf(q) >= 0 || a.CURP.toLowerCase().indexOf(q) >= 0)
    .slice(0, 25).map(a => ({ id: a.ID, nombre: a.Nombre, grupo: a.Grupo, curp: a.CURP, activo: a.Activo !== 'NO' }));
}

function fichaAlumno(token, id) {
  sesion(token, ['dir']);
  const alumnos = filas('Alumnos');
  const a = alumnos.find(x => x.ID === id);
  if (!a) throw new Error('Alumno no encontrado.');
  const f = filas('Familias').find(x => x.ID === a.FamiliaID) || {};
  return {
    alumno: { id: a.ID, nombre: a.Nombre, curp: a.CURP, grupo: a.Grupo, qr: a.QR, activo: a.Activo !== 'NO', foto: fotoDataUrl(a.FotoID) },
    familia: { id: f.ID, tutor: f.Tutor, pin: f.PIN, correo: f.Correo, telefono: f.Telefono },
    hermanos: alumnos.filter(x => x.FamiliaID === a.FamiliaID && x.ID !== a.ID && x.Activo !== 'NO').map(x => ({ id: x.ID, nombre: x.Nombre, grupo: x.Grupo })),
    personas: personasDe(a.FamiliaID),
    salidas: filas('Salidas').filter(x => x.AlumnoID === a.ID).slice(-15).reverse().map(salidaCorta)
  };
}

function crearAlumno(d) {
  const nombre = txt(d.nombre);
  const curp = txt(d.curp).toUpperCase();
  const grupo = normalizarGrupo(d.grupo);
  if (!nombre) throw new Error('Falta el nombre.');
  if (!/^[A-Z0-9]{18}$/.test(curp)) throw new Error('El CURP debe tener 18 letras y números.');
  if (!grupo) throw new Error('Grupo no válido: "' + txt(d.grupo) + '". Usa el formato 3A o 3°A.');
  const alumnos = filas('Alumnos');
  if (alumnos.some(a => a.CURP.toUpperCase() === curp)) throw new Error('El CURP ' + curp + ' ya está registrado.');
  let famId, pin = '';
  const curpHermano = txt(d.curpHermano).toUpperCase();
  if (curpHermano) {
    const h = alumnos.find(a => a.CURP.toUpperCase() === curpHermano);
    if (!h) throw new Error('No encontramos al hermano con CURP ' + curpHermano + '.');
    famId = h.FamiliaID;
  } else {
    famId = nuevoId(); pin = nuevoPin();
    agregar('Familias', { ID: famId, Tutor: txt(d.tutor) || 'Tutor de ' + nombre.split(' ')[0], PIN: pin,
      Telefono: txt(d.telefono), Correo: validarCorreo(txt(d.correo)) });
  }
  const id = nuevoId();
  agregar('Alumnos', { ID: id, Nombre: nombre, CURP: curp, Grupo: grupo, QR: nuevoQR(), FamiliaID: famId, Activo: 'SI', FotoID: '' });
  return { id, pin, hermano: !!curpHermano };
}

function altaAlumno(token, d) {
  sesion(token, ['dir']);
  return conBloqueo(() => crearAlumno(d));
}

function editarAlumno(token, d) {
  sesion(token, ['dir']);
  return conBloqueo(() => {
    const alumnos = filas('Alumnos');
    const a = alumnos.find(x => x.ID === d.id);
    if (!a) throw new Error('Alumno no encontrado.');
    const cambios = {};
    if (d.nombre !== undefined) { if (!txt(d.nombre)) throw new Error('Falta el nombre.'); cambios.Nombre = txt(d.nombre); }
    if (d.curp !== undefined) {
      const c = txt(d.curp).toUpperCase();
      if (!/^[A-Z0-9]{18}$/.test(c)) throw new Error('El CURP debe tener 18 letras y números.');
      if (alumnos.some(x => x.ID !== a.ID && x.CURP.toUpperCase() === c)) throw new Error('Ese CURP ya es de otro alumno.');
      cambios.CURP = c;
    }
    if (d.grupo !== undefined) { const g = normalizarGrupo(d.grupo); if (!g) throw new Error('Grupo no válido.'); cambios.Grupo = g; }
    if (d.activo !== undefined) cambios.Activo = d.activo ? 'SI' : 'NO';
    actualizar('Alumnos', a._fila, cambios);
    return true;
  });
}

function fotoAlumno(token, id, dataUrl) {
  sesion(token, ['dir']);
  return conBloqueo(() => {
    const a = filas('Alumnos').find(x => x.ID === id);
    if (!a) throw new Error('Alumno no encontrado.');
    const nueva = guardarFoto(dataUrl, a.Nombre);
    if (!nueva) throw new Error('La foto no es válida.');
    borrarFoto(a.FotoID);
    actualizar('Alumnos', a._fila, { FotoID: nueva });
    return true;
  });
}

function editarFamilia(token, d) {
  sesion(token, ['dir']);
  return conBloqueo(() => {
    const f = filas('Familias').find(x => x.ID === d.id);
    if (!f) throw new Error('Familia no encontrada.');
    const pin = txt(d.pin);
    if (pin && !/^\d{4,6}$/.test(pin)) throw new Error('El PIN debe tener de 4 a 6 números.');
    actualizar('Familias', f._fila, { Tutor: txt(d.tutor) || f.Tutor, Correo: validarCorreo(txt(d.correo)), Telefono: txt(d.telefono), PIN: pin || f.PIN });
    return true;
  });
}

function restablecerPin(token, famId) {
  sesion(token, ['dir']);
  const f = filas('Familias').find(x => x.ID === famId);
  if (!f) throw new Error('Familia no encontrada.');
  const pin = nuevoPin();
  actualizar('Familias', f._fila, { PIN: pin });
  return pin;
}

function listaMaestros(token) {
  sesion(token, ['dir']);
  return filas('Maestros').sort((a, b) => GRUPOS.indexOf(a.Grupo) - GRUPOS.indexOf(b.Grupo))
    .map(m => ({ id: m.ID, nombre: m.Nombre, grupo: m.Grupo, usuario: m.Usuario, contrasena: m.Contrasena, correo: m.Correo }));
}

// Crea (sin id), edita (con id) o elimina (eliminar: true)
function guardarMaestro(token, d) {
  sesion(token, ['dir']);
  return conBloqueo(() => {
    const maestros = filas('Maestros');
    const m = d.id ? maestros.find(x => x.ID === d.id) : null;
    if (d.id && !m) throw new Error('Maestro no encontrado.');
    if (d.eliminar) { hoja('Maestros').deleteRow(m._fila); return { eliminado: true }; }
    const nombre = txt(d.nombre), usuario = txt(d.usuario).toLowerCase(), grupo = normalizarGrupo(d.grupo);
    if (!nombre || !usuario) throw new Error('Escribe nombre y usuario.');
    if (!grupo) throw new Error('Elige un grupo.');
    if (maestros.some(x => x.ID !== d.id && x.Usuario.toLowerCase() === usuario)) throw new Error('Ese usuario ya existe.');
    const otro = maestros.find(x => x.ID !== d.id && x.Grupo === grupo);
    if (otro) throw new Error('El grupo ' + grupo + ' ya lo tiene ' + otro.Nombre + '.');
    const clave = txt(d.contrasena) || (m ? m.Contrasena : String(Math.floor(100000 + Math.random() * 900000)));
    const datos = { Nombre: nombre, Grupo: grupo, Usuario: usuario, Contrasena: clave, Correo: validarCorreo(txt(d.correo)) };
    if (m) actualizar('Maestros', m._fila, datos); else agregar('Maestros', Object.assign({ ID: nuevoId() }, datos));
    return { contrasena: clave };
  });
}

// filas: [{nombre, curp, grupo, tutor, telefono, correo, curpHermano}] leídas del Excel en el teléfono
function importarFilas(token, lista) {
  sesion(token, ['dir']);
  if (!Array.isArray(lista) || !lista.length) throw new Error('El archivo no tiene alumnos.');
  // Primero los que no dependen de un hermano, luego los que sí
  const orden = lista.map((r, i) => ({ r, i })).sort((a, b) => (txt(a.r.curpHermano) ? 1 : 0) - (txt(b.r.curpHermano) ? 1 : 0));
  const res = new Array(lista.length);
  conBloqueo(() => orden.forEach(({ r, i }) => {
    try { const x = crearAlumno(r); res[i] = { ok: true, msg: x.hermano ? 'Agregado a la familia de su hermano' : 'PIN ' + x.pin }; }
    catch (e) { res[i] = { ok: false, msg: e.message }; }
  }));
  return res;
}

function reporte(token, f) {
  sesion(token, ['dir']);
  f = f || {};
  return filas('Salidas').filter(x =>
    (!f.grupo || x.Grupo === f.grupo) && (!f.desde || x.Fecha >= f.desde) && (!f.hasta || x.Fecha <= f.hasta) &&
    (!f.alumno || x.Alumno.toLowerCase().indexOf(txt(f.alumno).toLowerCase()) >= 0)
  ).reverse().map(salidaCorta);
}

function exportarExcel(token, f) {
  const rows = reporte(token, f);
  const tmp = SpreadsheetApp.create('Reporte salidas ' + hoy());
  const datos = [['Fecha', 'Hora', 'Alumno', 'Grupo', 'Recogió', 'Parentesco', 'Entregó']]
    .concat(rows.map(r => [r.fecha, r.hora, r.alumno, r.grupo, r.recogio, r.parentesco, r.entrego]));
  const sh = tmp.getSheets()[0];
  sh.getRange(1, 1, datos.length, 7).setNumberFormat('@').setValues(datos);
  sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  sh.autoResizeColumns(1, 7);
  SpreadsheetApp.flush();
  const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return { nombre: 'Salidas ' + (f && f.grupo ? f.grupo + ' ' : '') + hoy() + '.xlsx', base64: Utilities.base64Encode(blob.getBytes()) };
}

function cerrarCiclo() {
  return conBloqueo(() => {
    let subieron = 0, egresaron = 0;
    filas('Alumnos').forEach(a => {
      if (a.Activo === 'NO') return;
      const g = grado(a.Grupo);
      if (g >= 6) { actualizar('Alumnos', a._fila, { Activo: 'NO' }); egresaron++; }
      else { actualizar('Alumnos', a._fila, { Grupo: (g + 1) + '°' + a.Grupo.slice(-1) }); subieron++; }
    });
    hoja('Salidas').copyTo(SpreadsheetApp.getActive()).setName('Salidas hasta ' + hoy() + ' ' + ahora().replace(':', ''));
    ['Salidas', 'Permisos'].forEach(n => {
      const h = hoja(n);
      if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, HOJAS[n].length).clearContent();
    });
    return { subieron, egresaron };
  });
}

function cerrarCicloDir(token, confirmacion) {
  sesion(token, ['dir']);
  if (txt(confirmacion).toUpperCase() !== 'CERRAR') throw new Error('Escribe CERRAR para confirmar.');
  return cerrarCiclo();
}

function cambiarClaveDir(token, actual, nueva) {
  sesion(token, ['dir']);
  if (txt(actual) !== cfg('contrasena_direccion')) throw new Error('La contraseña actual no es correcta.');
  if (txt(nueva).length < 6) throw new Error('La nueva contraseña debe tener al menos 6 caracteres.');
  setCfg('contrasena_direccion', txt(nueva));
  return true;
}
