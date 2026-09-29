'use strict';
// Verifica Validador, MotorCalculo y Exportador de index.html contra las pruebas Given-When-Then de Elaboration I.
// Uso: node tests/elaboration1.test.js   (sin dependencias; el navegador no interviene)

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const raiz = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(raiz, 'index.html'), 'utf8');
const vupTexto = fs.readFileSync(path.join(raiz, 'VUP.md'), 'utf8');

function scriptPorId(id) {
  const coincidencia = html.match(new RegExp('<script id="' + id + '">([\\s\\S]*?)</script>'));
  if (!coincidencia) throw new Error('No se encontró <script id="' + id + '"> en index.html');
  return coincidencia[1];
}

// El script de la app no toca el DOM al cargarse fuera del navegador (document no existe).
function cargarApp(contextoExtra) {
  const contexto = vm.createContext(Object.assign({ console }, contextoExtra || {}));
  return vm.runInContext(
    scriptPorId('app') + '\n;({ Validador, MotorCalculo, MotorGraficos, Persistencia, Exportador, Vista });',
    contexto
  );
}

const app = cargarApp();
const { Validador, MotorCalculo, MotorGraficos, Persistencia, Exportador, Vista } = app;

// ---------------------------------------------------------------------------
const resultados = [];
function prueba(id, descripcion, funcion) {
  const fallos = [];
  const esperar = (condicion, detalle) => { if (!condicion) fallos.push(detalle); };
  const igual = (real, esperado, detalle) => {
    const coincide = typeof esperado === 'number' ? Math.abs(real - esperado) < 0.005 : JSON.stringify(real) === JSON.stringify(esperado);
    if (!coincide) fallos.push(detalle + ': se esperaba ' + JSON.stringify(esperado) + ' y salió ' + JSON.stringify(real));
  };
  try { funcion(esperar, igual); } catch (error) { fallos.push('excepción: ' + error.message); }
  resultados.push({ id, descripcion, ok: fallos.length === 0, fallos });
}
const fila = (matriz, datos) => {
  const validacion = Validador.validar(matriz, datos);
  if (!validacion.valido) throw new Error('la validación rechazó los datos: ' + validacion.errores.join(' | '));
};

// ---------------------------------------------------------------------------
// Historia 1: BCG
prueba('1.1', 'BCG: cuadrantes y tamaño de burbuja', (esperar, igual) => {
  const datos = { tamanoPor: 'ingresos', divisiones: [
    { nombre: 'A', ingresos: '500', utilidades: '100', participacionRelativa: '1.80', crecimiento: '15' },
    { nombre: 'B', ingresos: '300', utilidades: '30', participacionRelativa: '0.40', crecimiento: '12' },
    { nombre: 'C', ingresos: '150', utilidades: '60', participacionRelativa: '1.50', crecimiento: '4' },
    { nombre: 'D', ingresos: '50', utilidades: '10', participacionRelativa: '0.30', crecimiento: '2' }
  ] };
  fila('BCG', datos);
  const r = MotorCalculo.calcularBCG(datos.divisiones);
  igual(r.divisiones.map((d) => d.cuadrante), ['Estrella', 'Interrogante', 'Vaca lechera', 'Perro'], 'cuadrantes');
  igual(r.totalIngresos, 1000, 'ingresos totales');
  igual(r.totalUtilidades, 200, 'utilidades totales');
  r.divisiones.forEach((d, i) => {
    igual(d.porcentajeIngresos, [50, 30, 15, 5][i], '% ingresos de ' + d.nombre);
    igual(d.porcentajeUtilidades, [50, 15, 30, 5][i], '% utilidades de ' + d.nombre);
  });
  igual(r.divisiones.reduce((s, d) => s + d.porcentajeIngresos, 0), 100, 'suma % ingresos');
  igual(r.divisiones.reduce((s, d) => s + d.porcentajeUtilidades, 0), 100, 'suma % utilidades');
});

// Historia 2: EFI y EFE
const factoresEFI = () => [
  { tipo: 'fortaleza', peso: 0.20, clasificacion: 4 }, { tipo: 'fortaleza', peso: 0.15, clasificacion: 4 }, { tipo: 'fortaleza', peso: 0.10, clasificacion: 3 },
  { tipo: 'debilidad', peso: 0.25, clasificacion: 1 }, { tipo: 'debilidad', peso: 0.20, clasificacion: 2 }, { tipo: 'debilidad', peso: 0.10, clasificacion: 1 }
];
const factoresEFE = () => [
  { tipo: 'oportunidad', peso: 0.25, clasificacion: 4 }, { tipo: 'oportunidad', peso: 0.20, clasificacion: 3 }, { tipo: 'oportunidad', peso: 0.10, clasificacion: 2 },
  { tipo: 'amenaza', peso: 0.25, clasificacion: 3 }, { tipo: 'amenaza', peso: 0.15, clasificacion: 2 }, { tipo: 'amenaza', peso: 0.05, clasificacion: 1 }
];

prueba('2.1', 'EFI débil (total 2.45)', (esperar, igual) => {
  fila('EFI', { factores: factoresEFI() });
  const r = MotorCalculo.calcularEFI(factoresEFI());
  igual(r.sumaPesos, 1, 'suma de pesos');
  igual(r.factores.map((f) => f.ponderado), [0.80, 0.60, 0.30, 0.25, 0.40, 0.10], 'ponderados');
  igual(r.total, 2.45, 'total EFI');
  igual(r.diagnostico, 'posición interna débil', 'diagnóstico');
});

prueba('2.2', 'EFE que aprovecha oportunidades (total 2.90)', (esperar, igual) => {
  fila('EFE', { factores: factoresEFE() });
  const r = MotorCalculo.calcularEFE(factoresEFE());
  igual(r.sumaPesos, 1, 'suma de pesos');
  igual(r.factores.map((f) => f.ponderado), [1.00, 0.60, 0.20, 0.75, 0.30, 0.05], 'ponderados');
  igual(r.total, 2.90, 'total EFE');
  igual(r.diagnostico, 'la organización aprovecha oportunidades y evita amenazas por encima del promedio', 'diagnóstico');
});

prueba('2.3', 'Pesos que no suman 1 (0.95)', (esperar, igual) => {
  const factores = factoresEFI();
  factores[5].peso = 0.05;
  const v = Validador.validar('EFI', { factores });
  esperar(v.valido === false, 'la validación debía rechazar los datos');
  esperar(v.errores.some((e) => e.includes('0.95') && e.includes('1.00')), 'el mensaje debía mostrar 0.95 y 1.00: ' + JSON.stringify(v.errores));
  igual(MotorCalculo.calcularEFI(factores).total, 2.40, 'total aritmético (referencia)');
});

prueba('2.4', 'Tolerancia de punto flotante (0.6 + 0.3 + 0.1)', (esperar) => {
  esperar(0.6 + 0.3 + 0.1 !== 1, 'precondición: la suma en punto flotante no debe ser exactamente 1');
  const factores = [
    { tipo: 'oportunidad', peso: 0.6, clasificacion: 4 }, { tipo: 'oportunidad', peso: 0.3, clasificacion: 3 }, { tipo: 'amenaza', peso: 0.1, clasificacion: 2 }
  ];
  esperar(Validador.validarPesos(factores) === true, 'validarPesos debía aceptar la suma como 1');
  esperar(Validador.validar('EFE', { factores }).valido === true, 'validar debía aceptar el EFE');
});

// Historia 3: MPC
const factoresMPC = (pesos) => [
  { nombre: 'Participación de mercado', peso: pesos[0] }, { nombre: 'Competitividad de precios', peso: pesos[1] },
  { nombre: 'Calidad del producto', peso: pesos[2] }, { nombre: 'Lealtad del cliente', peso: pesos[3] }
];
const empresasMPC = () => [
  { nombre: 'Mi empresa', clasificaciones: [3, 2, 4, 3] }, { nombre: 'Competidor A', clasificaciones: [4, 3, 3, 2] }, { nombre: 'Competidor B', clasificaciones: [2, 4, 2, 4] }
];

prueba('3.1', 'MPC: mi empresa frente a dos competidores', (esperar, igual) => {
  const factores = factoresMPC([0.30, 0.25, 0.25, 0.20]);
  fila('MPC', { factores, empresas: empresasMPC() });
  const r = MotorCalculo.calcularMPC(factores, empresasMPC());
  igual(r.empresas.map((e) => e.total), [3.00, 3.10, 2.90], 'totales (mi empresa, A, B)');
  igual(r.empresas[0].ponderados, [0.90, 0.50, 1.00, 0.60], 'ponderados de mi empresa');
  igual(r.empresas[1].ponderados, [1.20, 0.75, 0.75, 0.40], 'ponderados de A');
  igual(r.empresas[2].ponderados, [0.60, 1.00, 0.50, 0.80], 'ponderados de B');
  igual(r.ranking.map((e) => e.nombre), ['Competidor A', 'Mi empresa', 'Competidor B'], 'orden');
  igual(r.posicionPropia, 2, 'posición de mi empresa');
  igual(r.totalEmpresas, 3, 'cantidad de empresas');
});

prueba('3.2', 'MPC: los pesos se comparten entre todas las empresas', (esperar, igual) => {
  const factores = factoresMPC([0.10, 0.50, 0.20, 0.20]);
  fila('MPC', { factores, empresas: empresasMPC() });
  const r = MotorCalculo.calcularMPC(factores, empresasMPC());
  igual(r.empresas.map((e) => e.total), [2.70, 2.90, 3.40], 'totales (mi empresa, A, B)');
  igual(r.ranking.map((e) => e.nombre), ['Competidor B', 'Competidor A', 'Mi empresa'], 'orden');
  igual(r.posicionPropia, 3, 'posición de mi empresa');
});

// Historia 4: PEYEA
function casoPEYEA(id, nombre, ejes, promedios, x, y, cuadrante) {
  prueba(id, 'PEYEA: perfil ' + nombre, (esperar, igual) => {
    fila('PEYEA', ejes);
    const r = MotorCalculo.calcularPEYEA(ejes);
    igual(r.promedios.FF, promedios[0], 'FF');
    igual(r.promedios.FI, promedios[1], 'FI');
    igual(r.promedios.VC, promedios[2], 'VC');
    igual(r.promedios.EE, promedios[3], 'EE');
    igual(r.x, x, 'X');
    igual(r.y, y, 'Y');
    igual(r.vector, { x, y }, 'vector');
    igual(r.cuadrante, cuadrante, 'cuadrante');
  });
}
casoPEYEA('4.1', 'agresivo', { FF: [5, 4, 4, 3], FI: [4, 5, 3], VC: [-2, -3, -1, -2], EE: [-3, -2, -4, -3, -3] }, [4.00, 4.00, -2.00, -3.00], 2.00, 1.00, 'agresivo');
casoPEYEA('4.2', 'conservador', { FF: [5, 5, 5], FI: [1, 2], VC: [-4, -4], EE: [-2, -1, -3] }, [5.00, 1.50, -4.00, -2.00], -2.50, 3.00, 'conservador');
casoPEYEA('4.3', 'defensivo', { FF: [1, 2], FI: [1, 1], VC: [-4, -4, -4], EE: [-5, -5] }, [1.50, 1.00, -4.00, -5.00], -3.00, -3.50, 'defensivo');
casoPEYEA('4.4', 'competitivo', { FF: [1, 2, 3], FI: [5, 5, 5], VC: [-1, -2], EE: [-5, -5, -5] }, [2.00, 5.00, -1.50, -5.00], 3.50, -3.00, 'competitivo');

// Historia 5: MIE
function casoMIE(id, descripcion, totalEFI, totalEFE, nivelEFI, nivelEFE, celda, zona) {
  prueba(id, 'MIE: ' + descripcion, (esperar, igual) => {
    fila('MIE', { totalEFI, totalEFE });
    const r = MotorCalculo.ubicarMIE(totalEFI, totalEFE);
    igual(r.nivelEFI, nivelEFI, 'nivel EFI');
    igual(r.nivelEFE, nivelEFE, 'nivel EFE');
    igual(r.celda, celda, 'celda');
    if (zona) igual(r.zona, zona, 'zona');
  });
}
prueba('5.1', 'MIE: celda central con los totales de 2.1 y 2.2', (esperar, igual) => {
  const efi = MotorCalculo.calcularEFI(factoresEFI()).total;
  const efe = MotorCalculo.calcularEFE(factoresEFE()).total;
  igual(efi, 2.45, 'total EFI derivado');
  igual(efe, 2.90, 'total EFE derivado');
  const r = MotorCalculo.ubicarMIE(efi, efe);
  igual([r.nivelEFI, r.nivelEFE, r.celda, r.zona], ['promedio', 'promedio', 'V', 'retener y mantener'], 'ubicación');
});
casoMIE('5.2', 'esquina de crecimiento', 3.20, 3.50, 'fuerte', 'fuerte', 'I', 'crecer y construir');
casoMIE('5.3', 'esquina de cosecha', 1.80, 1.50, 'débil', 'débil', 'IX', 'cosechar o desinvertir');
prueba('5.4', 'MIE: valores en el límite de rango (EFE = 3.50)', (esperar, igual) => {
  const celdas = [1.99, 2.00, 2.99, 3.00].map((efi) => MotorCalculo.ubicarMIE(efi, 3.50).celda);
  igual(celdas, ['III', 'II', 'II', 'I'], 'celdas para EFI 1.99, 2.00, 2.99 y 3.00');
});

// Historia 6: GE (sin implementar por decisión del juez)
prueba('6.1', 'GE: sin lógica de posicionamiento (falla de forma explícita, sin resultado inventado)', (esperar) => {
  let lanzoUbicar = false, lanzoDibujar = false;
  try { MotorCalculo.ubicarGE({}); } catch (error) { lanzoUbicar = /pendiente/i.test(error.message); }
  try { MotorGraficos.dibujarGE({}); } catch (error) { lanzoDibujar = /pendiente/i.test(error.message); }
  esperar(lanzoUbicar, 'ubicarGE debía lanzar un error de módulo pendiente');
  esperar(lanzoDibujar, 'dibujarGE debía lanzar un error de módulo pendiente');
  const v = Validador.validar('GE', {});
  esperar(v.valido === false && /pendiente/i.test(v.errores[0]), 'validar("GE") debía indicar que está pendiente');
});

// ---------------------------------------------------------------------------
// Comprobaciones adicionales (no son pruebas de Elaboration I)
prueba('X.1', 'Firmas: los seis objetos tienen exactamente los métodos del diagrama de clases de Elaboration II', (esperar, igual) => {
  const diagrama = vupTexto.match(/```mermaid\nclassDiagram([\s\S]*?)```/)[1];
  const esperadas = {};
  let clase = null;
  diagrama.split('\n').forEach((linea) => {
    const c = linea.match(/^\s*class (\w+) \{/);
    if (c) { clase = c[1]; esperadas[clase] = {}; return; }
    const m = linea.match(/^\s*\+(\w+)\(([^)]*)\)/);
    if (m && clase) esperadas[clase][m[1]] = m[2].trim() === '' ? 0 : m[2].split(',').length;
  });
  const reales = { Vista, Validador, MotorCalculo, MotorGraficos, Persistencia, Exportador };
  Object.keys(esperadas).forEach((nombre) => {
    const obtenidas = {};
    Object.keys(reales[nombre]).forEach((metodo) => { obtenidas[metodo] = reales[nombre][metodo].length; });
    igual(obtenidas, esperadas[nombre], 'métodos y cantidad de parámetros de ' + nombre);
  });
  igual(Object.keys(esperadas).sort(), Object.keys(reales).sort(), 'componentes');
});

prueba('X.2', 'Validación: rangos, vacíos y formatos', (esperar) => {
  const rechaza = (matriz, datos, detalle) => esperar(Validador.validar(matriz, datos).valido === false, detalle);
  const acepta = (matriz, datos, detalle) => esperar(Validador.validar(matriz, datos).valido === true, detalle + ': ' + Validador.validar(matriz, datos).errores.join(' | '));
  const peyea = (cambios) => Object.assign({ FF: [3], FI: [3], VC: [-3], EE: [-3] }, cambios);
  acepta('PEYEA', peyea({}), 'PEYEA base');
  rechaza('PEYEA', peyea({ FF: [0] }), 'FF = 0');
  rechaza('PEYEA', peyea({ FI: [7] }), 'FI = 7');
  rechaza('PEYEA', peyea({ VC: [0] }), 'VC = 0');
  rechaza('PEYEA', peyea({ EE: [-7] }), 'EE = -7');
  rechaza('PEYEA', peyea({ VC: [] }), 'eje sin factores');
  const ef = (clasificacion, peso) => ({ factores: [{ tipo: 'fortaleza', peso, clasificacion }] });
  acepta('EFI', ef(4, 1), 'un factor con peso 1');
  rechaza('EFI', ef(5, 1), 'clasificación 5');
  rechaza('EFI', ef(0, 1), 'clasificación 0');
  rechaza('EFI', ef(2.5, 1), 'clasificación 2.5');
  rechaza('EFI', ef('', 1), 'clasificación vacía');
  rechaza('EFI', ef(3, ''), 'peso vacío');
  rechaza('EFI', ef(3, 1.2), 'peso mayor que 1');
  rechaza('EFI', ef(3, 'abc'), 'peso no numérico');
  acepta('EFI', { factores: [{ tipo: 'fortaleza', peso: '0,4', clasificacion: '3' }, { tipo: 'debilidad', peso: '0,6', clasificacion: '2' }] }, 'coma decimal');
  rechaza('EFI', { factores: [] }, 'sin factores');
  rechaza('MIE', { totalEFI: 0.9, totalEFE: 2 }, 'total EFI fuera de 1 a 4');
  rechaza('MIE', { totalEFI: 2, totalEFE: 4.1 }, 'total EFE fuera de 1 a 4');
  acepta('MIE', { totalEFI: 1, totalEFE: 4 }, 'extremos 1 y 4');
  rechaza('MPC', { factores: factoresMPC([0.3, 0.25, 0.25, 0.2]), empresas: [empresasMPC()[0]] }, 'MPC sin competidores');
  rechaza('MPC', { factores: factoresMPC([0.3, 0.25, 0.25, 0.2]), empresas: [{ nombre: 'Yo', clasificaciones: [3, 2] }, empresasMPC()[1]] }, 'clasificaciones incompletas');
  const bcg = (d) => ({ tamanoPor: 'ingresos', divisiones: [Object.assign({ nombre: '', ingresos: '10', utilidades: '2', participacionRelativa: '1', crecimiento: '5' }, d)] });
  acepta('BCG', bcg({}), 'BCG base con nombre vacío (nombre opcional)');
  rechaza('BCG', bcg({ ingresos: '' }), 'ingresos vacíos');
  rechaza('BCG', bcg({ participacionRelativa: '0' }), 'participación relativa 0');
  rechaza('BCG', bcg({ utilidades: '-5' }), 'utilidades totales negativas');
  rechaza('BCG', { tamanoPor: 'ingresos', divisiones: [] }, 'sin divisiones');
});

prueba('X.3', 'Límites documentados como [VERIFICAR]: la Vista puede avisarlos', (esperar, igual) => {
  const efi25 = MotorCalculo.calcularEFI([{ tipo: 'fortaleza', peso: 0.5, clasificacion: 3 }, { tipo: 'debilidad', peso: 0.5, clasificacion: 2 }]);
  igual(efi25.total, 2.5, 'total exactamente 2.5');
  esperar(efi25.enLimite === true, 'enLimite debía ser true con total 2.5');
  igual(efi25.diagnostico, 'posición interna fuerte', 'caso más común: 2.5 no es débil');
  esperar(MotorCalculo.calcularPEYEA({ FF: [3], FI: [3], VC: [-3], EE: [-3] }).enEje === true, 'PEYEA con X = 0 debía avisar enEje');
  igual(MotorCalculo.ubicarMIE(1.995, 2).celda, 'VI', 'MIE con EFI 1.995 y EFE 2.0 se trata como EFI débil y EFE promedio (< 2.0)');
  const bcgLimite = MotorCalculo.calcularBCG([{ nombre: 'L', ingresos: 1, utilidades: 1, participacionRelativa: 1, crecimiento: 10 }]);
  esperar(bcgLimite.divisiones[0].cuadrante === 'Estrella' && bcgLimite.divisiones[0].enLimite === true, 'BCG en el umbral: Estrella y enLimite');
});

// Exportador: se prueba con la copia de SheetJS que está embebida en index.html.
prueba('X.4', 'Exportador: genera las hojas "Datos" y "Resultados" con SheetJS embebido', (esperar, igual) => {
  const contexto = vm.createContext({ console });
  vm.runInContext(scriptPorId('sheetjs'), contexto);
  esperar(contexto.XLSX && contexto.XLSX.version === '0.20.3', 'SheetJS 0.20.3 debía cargarse desde el script embebido');
  let capturado = null;
  const original = contexto.XLSX.writeFile;
  contexto.XLSX.writeFile = (libro, nombre) => { capturado = { libro, nombre }; };
  const appConXLSX = cargarApp({ XLSX: contexto.XLSX });
  const datos = { factores: factoresEFI().map((f, i) => Object.assign({ nombre: 'Factor ' + (i + 1) }, f)) };
  const resultado = MotorCalculo.calcularEFI(datos.factores);
  const nombre = appConXLSX.Exportador.exportarXLSX('EFI', datos, resultado);
  igual(nombre, 'Mtx-EFI.xlsx', 'nombre del archivo');
  esperar(capturado !== null && capturado.nombre === 'Mtx-EFI.xlsx', 'debía llamar a XLSX.writeFile con el nombre del archivo');
  igual(capturado.libro.SheetNames, ['Datos', 'Resultados'], 'hojas');
  const hojaDatos = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Datos, { header: 1 });
  const hojaResultados = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Resultados, { header: 1 });
  igual(hojaDatos[0], ['Tipo', 'Factor', 'Peso', 'Clasificación'], 'encabezado de Datos');
  igual(hojaDatos.length, 7, 'filas de Datos (encabezado + 6 factores)');
  const total = hojaResultados.find((f) => f[0] === 'Total EFI');
  esperar(total && Math.abs(total[1] - 2.45) < 1e-9, 'la hoja Resultados debía traer Total EFI = 2.45');
  const diagnostico = hojaResultados.find((f) => f[0] === 'Diagnóstico');
  igual(diagnostico[1], 'Posición interna débil', 'diagnóstico en Resultados');
  contexto.XLSX.writeFile = original;
  let sinLibreria = false;
  try { app.Exportador.exportarXLSX('EFI', datos, resultado); } catch (error) { sinLibreria = /librería/i.test(error.message); }
  esperar(sinLibreria, 'sin XLSX debía fallar con un mensaje claro');
});

// Persistencia con un localStorage simulado.
prueba('X.5', 'Persistencia: guardar, cargar, estado corrupto y almacenamiento inaccesible', (esperar, igual) => {
  const almacen = {};
  const localStorageSimulado = {
    getItem: (k) => (k in almacen ? almacen[k] : null), setItem: (k, v) => { almacen[k] = String(v); }, removeItem: (k) => { delete almacen[k]; }
  };
  const p = cargarApp({ localStorage: localStorageSimulado }).Persistencia;
  const vacio = p.cargar();
  igual(vacio.matrizActiva, 'BCG', 'estado vacío al inicio');
  vacio.matrizActiva = 'PEYEA';
  vacio.datos.PEYEA.FF[0].valor = '4';
  esperar(p.guardar(vacio) === true, 'guardar debía devolver true');
  const recuperado = p.cargar();
  igual(recuperado.matrizActiva, 'PEYEA', 'matriz activa recuperada');
  igual(recuperado.datos.PEYEA.FF[0].valor, '4', 'dato recuperado');
  almacen['mtx.estado'] = '{no es json';
  igual(p.cargar().matrizActiva, 'BCG', 'JSON corrupto devuelve estado vacío');
  almacen['mtx.estado'] = JSON.stringify({ version: 1, matrizActiva: 'BCG', datos: {} });
  igual(p.cargar().matrizActiva, 'BCG', 'estructura inválida devuelve estado vacío');
  p.limpiar();
  esperar(!('mtx.estado' in almacen), 'limpiar debía borrar la clave');
  const roto = { getItem() { throw new Error('bloqueado'); }, setItem() { throw new Error('bloqueado'); }, removeItem() { throw new Error('bloqueado'); } };
  const pRoto = cargarApp({ localStorage: roto }).Persistencia;
  igual(pRoto.cargar().matrizActiva, 'BCG', 'almacenamiento inaccesible devuelve estado vacío');
  esperar(pRoto.guardar(vacio) === false, 'guardar sin acceso debía devolver false y no lanzar');
  pRoto.limpiar();
});

// ---------------------------------------------------------------------------
let fallidas = 0;
resultados.forEach((r) => {
  console.log((r.ok ? 'OK    ' : 'FALLA ') + r.id.padEnd(4) + ' ' + r.descripcion);
  r.fallos.forEach((f) => console.log('        - ' + f));
  if (!r.ok) fallidas++;
});
console.log('\n' + (resultados.length - fallidas) + ' de ' + resultados.length + ' comprobaciones correctas.');
process.exit(fallidas === 0 ? 0 : 1);
