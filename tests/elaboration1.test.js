'use strict';
// Verifica Validador, MotorCalculo y Exportador de index.html contra las pruebas Given-When-Then de Elaboration I
// (Módulo 1: pruebas 1.1 a 6.1; Módulo 2, Análisis Estructural: pruebas AE.1 a AE.10; Módulo 3, Radar Estratégico: pruebas RE.1 a RE.11).
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
// Además de los seis componentes devuelve evaluarMatriz y fijarEstado, para probar el estado "vacía", "inválida" u "ok" de una
// matriz con el mismo evaluarMatriz de la aplicación (usa la variable de módulo "estado", que la rutina de arranque no llena en Node).
function cargarApp(contextoExtra) {
  const contexto = vm.createContext(Object.assign({ console }, contextoExtra || {}));
  return vm.runInContext(
    scriptPorId('app') + '\n;({ Validador, MotorCalculo, MotorGraficos, Persistencia, Exportador, Vista, evaluarMatriz, fijarEstado: (nuevo) => { estado = nuevo; } });',
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

// Historia: Análisis Estructural (Módulo 2). El arnés verifica el cálculo, la validación y la persistencia; el dibujo de dibujarAE,
// la diagonal bloqueada en pantalla y el orden de la hoja Validadas se verifican en Construction III, en un navegador.
const { evaluarMatriz, fijarEstado } = app;
const nombresAE = (n) => Array.from({ length: n }, (_, i) => 'V' + (i + 1));
// Como el formulario: las calificaciones se guardan como texto y la diagonal es null.
const enTextoAE = (matriz) => matriz.map((fila) => fila.map((v) => (v === null ? null : String(v))));
const datosAE = (matriz, marcas) => ({ variables: nombresAE(matriz.length), matriz: enTextoAE(matriz), marcas: marcas || matriz.map(() => null) });
const matrizAE1 = () => [[null, 4, 0, 4], [1, null, 2, 4], [0, 1, null, 0], [0, 2, 0, null]];
const evaluarAE = (datos) => { fijarEstado({ datos: { AE: datos } }); return evaluarMatriz('AE'); };
// "igual" compara las listas con JSON exacto; para listas de números (proyecciones con decimales) se compara elemento a elemento, con la tolerancia de dos decimales.
const igualLista = (igual, reales, esperados, detalle) => {
  igual(reales.length, esperados.length, detalle + ' (cantidad)');
  esperados.forEach((esperado, i) => igual(reales[i], esperado, detalle + ' [' + i + ']'));
};

function casoAE(id, descripcion, matriz, esperado) {
  prueba(id, 'AE: ' + descripcion, (esperar, igual) => {
    const datos = datosAE(matriz);
    fila('AE', datos);
    const r = MotorCalculo.calcularAE(datos.variables, datos.matriz);
    igual(r.corteY, esperado.corteY, 'corteY');
    igual(r.corteX, esperado.corteX, 'corteX');
    igual(r.variables.map((v) => v.motricidad), esperado.motricidad, 'motricidades');
    igual(r.variables.map((v) => v.dependencia), esperado.dependencia, 'dependencias');
    igual(r.variables.map((v) => v.cuadrante), esperado.cuadrantes, 'cuadrantes');
    if (esperado.proyeccionX) {
      igualLista(igual, r.variables.map((v) => v.proyeccion.x), esperado.proyeccionX, 'proyección x');
      igualLista(igual, r.variables.map((v) => v.proyeccion.y), esperado.proyeccionY, 'proyección y');
      igualLista(igual, r.variables.map((v) => v.puntoProyeccion), esperado.punto, 'punto de proyección sobre la diagonal');
    }
  });
}

casoAE('AE.1', 'cuatro variables que cubren los cuatro cuadrantes', matrizAE1(), {
  corteY: 4, corteX: 4, motricidad: [8, 7, 1, 2], dependencia: [1, 7, 2, 8],
  cuadrantes: ['INDEPENDIENTES', 'AMBIGUAS', 'AUTONOMAS', 'DEPENDIENTES'],
  proyeccionX: [3.5, 0, -0.5, -3], proyeccionY: [4.9497, 0, 0.7071, 4.2426], punto: [4.5, 7, 1.5, 5]
});
casoAE('AE.2', 'motricidad exactamente en el corte cuenta como alta', [[null, 2, 2], [1, null, 1], [0, 0, null]], {
  corteY: 2, corteX: 1.5, motricidad: [4, 2, 0], dependencia: [1, 2, 3], cuadrantes: ['INDEPENDIENTES', 'AMBIGUAS', 'DEPENDIENTES']
});
casoAE('AE.3', 'dependencia exactamente en el corte cuenta como alta', [[null, 1, 0], [2, null, 0], [2, 1, null]], {
  corteY: 1.5, corteX: 2, motricidad: [1, 2, 3], dependencia: [4, 2, 0], cuadrantes: ['DEPENDIENTES', 'AMBIGUAS', 'INDEPENDIENTES']
});
casoAE('AE.4', 'los cortes están en la mitad del máximo, no en el promedio',
  [[null, 4, 4, 4], [3, null, 2, 2], [3, 2, null, 2], [3, 2, 2, null]], {
    corteY: 6, corteX: 4.5, motricidad: [12, 7, 7, 7], dependencia: [9, 8, 8, 8], cuadrantes: ['AMBIGUAS', 'AMBIGUAS', 'AMBIGUAS', 'AMBIGUAS']
  });

prueba('AE.5', 'AE: la diagonal no cuenta como celda sin calificar (el bloqueo en pantalla se verifica en Construction III)', (esperar) => {
  // Seis celdas fuera de la diagonal completas y la diagonal en null: la matriz está completa.
  const completa = datosAE([[null, 1, 2], [3, null, 4], [0, 1, null]]);
  const v = Validador.validar('AE', completa);
  esperar(v.valido === true, 'la matriz completa debía validar: ' + v.errores.join(' | '));
  esperar(evaluarAE(completa).estado === 'ok', 'evaluarMatriz debía calcular con la matriz completa');
  // El cero es una calificación hecha, no una celda vacía.
  esperar(Validador.validar('AE', datosAE([[null, 0], [0, null]])).valido === true, 'una matriz de ceros debía ser válida');
});

prueba('AE.6', 'AE: celdas sin calificar y valores fuera de rango', (esperar, igual) => {
  const vacia = datosAE([[null, '', ''], ['', null, ''], ['', '', null]]);
  const e0 = evaluarAE(vacia);
  igual(e0.estado, 'vacia', 'ninguna celda con valor: estado');
  igual(e0.errores, [], 'ninguna celda con valor: sin errores');
  esperar(e0.resultado === undefined, 'ninguna celda con valor: no debía haber resultado');

  const incompleta = datosAE([[null, 1, 2], [3, null, 4], [0, '', null]]);
  const e1 = evaluarAE(incompleta);
  igual(e1.estado, 'invalida', 'cinco celdas con valor y una en blanco: estado');
  esperar(e1.errores.some((m) => /campos vac/i.test(m)), 'debía mostrar el error de campos vacíos: ' + JSON.stringify(e1.errores));
  esperar(e1.resultado === undefined, 'no debía calcular con la matriz incompleta');

  [5, -1].forEach((valor) => {
    const fuera = datosAE([[null, 1, 2], [3, null, 4], [0, valor, null]]);
    const e2 = evaluarAE(fuera);
    igual(e2.estado, 'invalida', 'valor ' + valor + ': estado');
    esperar(e2.errores.some((m) => m.includes('0') && m.includes('4')), 'el error debía indicar el rango de 0 a 4: ' + JSON.stringify(e2.errores));
    esperar(e2.resultado === undefined, 'valor ' + valor + ': no debía calcular');
  });
  // Los datos escritos no se pierden al validar.
  igual(incompleta.matriz[2], ['0', '', null], 'la matriz ingresada se conserva tal cual');
});

prueba('AE.7', 'AE: mínimo de dos variables', (esperar, igual) => {
  const una = { variables: ['V1'], matriz: [[null]], marcas: [null] };
  igual(evaluarAE(una).estado, 'vacia', 'una sola variable: estado');
  igual(evaluarAE(una).errores, [], 'una sola variable: sin errores');
  const v = Validador.validar('AE', una);
  esperar(v.valido === false && v.errores.some((m) => /al menos dos variables/i.test(m)), 'validar debía pedir al menos dos variables');

  const dos = datosAE([[null, 3], [1, null]]);
  fila('AE', dos);
  const r = MotorCalculo.calcularAE(dos.variables, dos.matriz);
  igual(r.corteY, 1.5, 'corteY');
  igual(r.corteX, 1.5, 'corteX');
  igual(r.variables.map((x) => x.cuadrante), ['INDEPENDIENTES', 'DEPENDIENTES'], 'cuadrantes');
  igualLista(igual, r.variables.map((x) => x.proyeccion.x), [1, -1], 'proyección x');
  igualLista(igual, r.variables.map((x) => x.proyeccion.y), [1.4142, 1.4142], 'proyección y');
});

prueba('AE.8', 'AE: el resultado conserva el orden de carga y trae las marcas (la hoja Validadas se verifica en Construction III)', (esperar, igual) => {
  const datos = datosAE(matrizAE1(), ['SI', 'NO', 'NO', 'SI']);
  const e = evaluarAE(datos);
  igual(e.estado, 'ok', 'estado');
  igual(e.resultado.variables.map((v) => v.nombre), ['V1', 'V2', 'V3', 'V4'], 'orden de carga');
  igual(e.resultado.variables.map((v) => v.marca), ['SI', 'NO', 'NO', 'SI'], 'marcas en el mismo orden');
  // Un orden por motricidad (V1, V2, V4, V3) o por dependencia (V1, V3, V2, V4) daría otro resultado.
  igual(e.resultado.variables.map((v) => v.motricidad), [8, 7, 1, 2], 'motricidades en orden de carga, no ordenadas');
  // Marcar V4 antes que V1 no cambia el orden.
  const alReves = datosAE(matrizAE1(), ['SI', null, null, 'SI']);
  igual(evaluarAE(alReves).resultado.variables.map((v) => v.marca), ['SI', null, null, 'SI'], 'marca sin definir se conserva como null');
});

prueba('AE.9', 'AE: ninguna variable marcada con SÍ', (esperar, igual) => {
  [[null, null, null, null], ['NO', 'NO', 'NO', 'NO']].forEach((marcas) => {
    const e = evaluarAE(datosAE(matrizAE1(), marcas));
    igual(e.estado, 'ok', 'el módulo sigue calculando con marcas ' + JSON.stringify(marcas));
    esperar(e.resultado.variables.every((v) => v.marca !== 'SI'), 'ninguna variable debía quedar con SÍ');
  });
});

prueba('AE.10', 'AE: los datos siguen tras recargar (persistencia y recálculo)', (esperar, igual) => {
  const almacen = {};
  const localStorageSimulado = {
    getItem: (k) => (k in almacen ? almacen[k] : null), setItem: (k, v) => { almacen[k] = String(v); }, removeItem: (k) => { delete almacen[k]; }
  };
  const app2 = cargarApp({ localStorage: localStorageSimulado });
  const p = app2.Persistencia;
  const inicial = p.cargar();
  igual(inicial.datos.AE, { variables: [], matriz: [], marcas: [] }, 'AE vacío en un estado nuevo');
  const datos = datosAE(matrizAE1(), ['SI', 'NO', 'NO', 'SI']);
  inicial.matrizActiva = 'AE';
  inicial.datos.AE = datos;
  esperar(p.guardar(inicial) === true, 'guardar debía devolver true');
  // "Recargar": otra instancia de la aplicación lee lo guardado.
  const p2 = cargarApp({ localStorage: localStorageSimulado }).Persistencia;
  const recuperado = p2.cargar();
  igual(recuperado.matrizActiva, 'AE', 'matriz activa recuperada');
  igual(recuperado.datos.AE, datos, 'variables, matriz (diagonal null) y marcas recuperadas');
  app2.fijarEstado(recuperado);
  const e = app2.evaluarMatriz('AE');
  igual(e.estado, 'ok', 'estado tras recargar');
  igual(e.resultado.variables.map((v) => v.cuadrante), ['INDEPENDIENTES', 'AMBIGUAS', 'AUTONOMAS', 'DEPENDIENTES'], 'mismos cuadrantes de AE.1');
  igual([e.resultado.corteY, e.resultado.corteX], [4, 4], 'mismos cortes de AE.1');
  igual(e.resultado.variables.map((v) => v.marca), ['SI', 'NO', 'NO', 'SI'], 'mismas marcas de AE.8');

  // Un estado guardado antes del módulo AE no trae datos.AE: se completa con AE vacío y los demás datos sobreviven.
  const viejo = p2.cargar();
  viejo.datos.BCG.divisiones[0].ingresos = '500';
  delete viejo.datos.AE;
  almacen['mtx.estado'] = JSON.stringify(viejo);
  const migrado = p2.cargar();
  igual(migrado.datos.AE, { variables: [], matriz: [], marcas: [] }, 'estado anterior sin AE: AE vacío');
  igual(migrado.datos.BCG.divisiones[0].ingresos, '500', 'estado anterior sin AE: los demás datos se conservan');
  // Un AE dañado (matriz que no es cuadrada, o diagonal con valor) también se reemplaza por AE vacío.
  const danado = p2.cargar();
  danado.datos.AE = { variables: ['V1', 'V2'], matriz: [[null, '1']], marcas: [null, null] };
  almacen['mtx.estado'] = JSON.stringify(danado);
  igual(p2.cargar().datos.AE, { variables: [], matriz: [], marcas: [] }, 'AE dañado: AE vacío');
});

// Historia: Radar Estratégico (Módulo 3). El arnés verifica la validación, el cálculo, el flujo de evaluarMatriz y la persistencia con los
// mismos datos de Elaboration I. Para lo visual (RE.4 y RE.5 en el dibujo, RE.6 y RE.7 en el formulario) no hay navegador: se usa un DOM
// simulado mínimo y se comprueba la estructura de lo que generan dibujarRadar y formularioRadar (cuántos puntos, a qué distancia del centro,
// qué rótulos, cuántos campos). Cómo se ve en pantalla (colores, tamaños, si los rótulos se pisan) se verifica en Construction III, en un navegador.
// La tabla de índices de abajo se transcribe de Elaboration I a propósito, sin derivarla de la aplicación, para comprobar a la aplicación contra ella.
const COMPONENTES_RADAR = [
  ['Movilización 1', 0, 4], ['Movilización 2', 4, 4], ['Movilización 3', 8, 4],
  ['Traducción 1', 12, 5], ['Traducción 2', 17, 4], ['Traducción 3', 21, 3],
  ['Alineamiento 1', 24, 4], ['Alineamiento 2', 28, 4],
  ['Motivación 1', 32, 4], ['Motivación 2', 36, 4], ['Motivación 3', 40, 4],
  ['Gestión 1', 44, 4], ['Gestión 2', 48, 4], ['Gestión 3', 52, 4]
];
const NOMBRES_RADAR = COMPONENTES_RADAR.map((c) => c[0]);
// Como el formulario: las calificaciones se guardan como texto. "asignaciones" = { 'Movilización 1': [0, 1, 2, 3], ... }; el resto queda en null.
function datosRadar(asignaciones) {
  const calificaciones = new Array(56).fill(null);
  Object.entries(asignaciones || {}).forEach(([nombre, valores]) => {
    const componente = COMPONENTES_RADAR.find((c) => c[0] === nombre);
    if (!componente || valores.length > componente[2]) throw new Error('componente de prueba inválido: ' + nombre);
    valores.forEach((v, k) => { calificaciones[componente[1] + k] = v === null ? null : String(v); });
  });
  return { calificaciones };
}
const casoRE1 = () => ({ 'Movilización 1': [0, 1, 2, 3], 'Traducción 1': [0, 0, 5, 5, 5], 'Traducción 3': [0, 0, 0], 'Gestión 3': [5, 5, 5, 5] });
const evaluarRadar = (datos) => { fijarEstado({ datos: { RADAR: datos } }); return evaluarMatriz('RADAR'); };
const estadosRadar = (resultado) => resultado.map((c) => c.estado);
const completosRadar = (resultado) => resultado.map((c, k) => (c.estado === 'completo' ? k : -1)).filter((k) => k >= 0);

// DOM simulado mínimo (solo lo que usan svg(), texto(), el() y tabla()): sin navegador, y solo para inspeccionar la estructura generada.
function nodoSimulado(etiqueta) {
  return {
    etiqueta, atributos: {}, hijos: [], className: '', textContent: '',
    setAttribute(clave, valor) { this.atributos[clave] = String(valor); },
    append(...nodos) { this.hijos.push(...nodos); },
    replaceChildren(...nodos) { this.hijos = nodos; }
  };
}
const recorrerNodos = (nodo, funcion) => { if (typeof nodo === 'string') return; funcion(nodo); nodo.hijos.forEach((h) => recorrerNodos(h, funcion)); };
const aplanarNodos = (raiz) => { const lista = []; recorrerNodos(raiz, (n) => lista.push(n)); return lista; };
const textoDeNodo = (nodo) => (typeof nodo === 'string' ? nodo : (nodo.textContent || '') + ' ' + nodo.hijos.map(textoDeNodo).join(' '));
function cargarAppConDocumento() {
  const contenedor = nodoSimulado('div');
  const contexto = vm.createContext({ console });
  const interfaz = vm.runInContext(scriptPorId('app') + '\n;({ MotorGraficos, formularioRadar, resultadosRadar, ESTRUCTURA_RADAR });', contexto);
  // document se asigna después de cargar el script: así la rutina de arranque (que solo corre si document existe al cargarse) no se ejecuta.
  contexto.document = { createElementNS: (espacio, etiqueta) => nodoSimulado(etiqueta), createElement: nodoSimulado, querySelector: () => contenedor };
  return Object.assign({ contenedor }, interfaz);
}
const appDOM = cargarAppConDocumento();
function trazadoRadar(resultado) {
  appDOM.MotorGraficos.dibujarRadar(resultado);
  const nodos = aplanarNodos(appDOM.contenedor.hijos[0]);
  const num = (n, clave) => Number(n.atributos[clave]);
  const anillo = nodos.find((n) => n.etiqueta === 'circle' && n.atributos.class === 'g-eje');
  const centro = { x: num(anillo, 'cx'), y: num(anillo, 'cy') };
  const radio = num(anillo, 'r');
  return {
    rayos: nodos.filter((n) => n.etiqueta === 'line' && n.atributos.class === 'g-rejilla').map((l) => ({ dx: num(l, 'x2') - centro.x, dy: num(l, 'y2') - centro.y })),
    puntos: nodos.filter((n) => n.etiqueta === 'circle' && n.atributos.class === 'punto-radar')
      .map((p) => ({ dx: num(p, 'cx') - centro.x, dy: num(p, 'cy') - centro.y, fraccion: Math.hypot(num(p, 'cx') - centro.x, num(p, 'cy') - centro.y) / radio })),
    rotulos: nodos.filter((n) => n.etiqueta === 'text').map((t) => t.hijos[0]).filter((t) => NOMBRES_RADAR.includes(t))
  };
}

prueba('RE.1', 'RADAR: cuatro componentes completos y diez vacíos (promedio con divisores 4, 5, 3 y 4)', (esperar, igual) => {
  const e = evaluarRadar(datosRadar(casoRE1()));
  igual(e.estado, 'ok', 'estado');
  igual(e.errores, [], 'sin errores');
  igual(e.resultado.length, 14, 'un resultado por componente');
  igual(completosRadar(e.resultado), [0, 3, 5, 13], 'componentes completos (Movilización 1, Traducción 1, Traducción 3, Gestión 3)');
  igual([0, 3, 5, 13].map((k) => e.resultado[k].puntaje), [1.5, 3.0, 0.0, 5.0], 'puntajes');
  esperar(e.resultado[0].puntaje === 1.5, 'el puntaje no se redondea internamente');
  igual(e.resultado.filter((c, k) => ![0, 3, 5, 13].includes(k)).map((c) => c.estado), new Array(10).fill('vacío'), 'los otros diez componentes están vacíos');
  esperar(e.resultado.every((c) => c.estado === 'completo' || c.puntaje === null), 'solo los completos traen puntaje');
  // El mismo caso, con las calificaciones como números (no como texto) y sin pasar por evaluarMatriz.
  const directo = MotorCalculo.calcularRadar(datosRadar(casoRE1()).calificaciones.map((c) => (c === null ? null : Number(c))), []);
  igual([0, 3, 5, 13].map((k) => directo[k].puntaje), [1.5, 3.0, 0.0, 5.0], 'calcularRadar directo');
  // Traducción 3 en 0.00 es un componente completo en el ideal; no es lo mismo que un componente vacío.
  igual(e.resultado[5], { estado: 'completo', puntaje: 0 }, 'Traducción 3 completo en 0.00');
  igual(e.resultado[4], { estado: 'vacío', puntaje: null }, 'Traducción 2 vacío');
});

prueba('RE.2', 'RADAR: un componente incompleto no se calcula, ni da error, y se calcula al completarlo', (esperar, igual) => {
  const parcial = datosRadar(Object.assign(casoRE1(), { 'Movilización 2': [1, 3] }));
  const e = evaluarRadar(parcial);
  igual(e.estado, 'ok', 'el módulo sigue calculando');
  igual(e.errores, [], 'sin ningún mensaje de error');
  igual(e.resultado[1], { estado: 'incompleto', puntaje: null }, 'Movilización 2 incompleto, sin puntaje (no 2.00)');
  igual(completosRadar(e.resultado), [0, 3, 5, 13], 'los mismos cuatro puntos de RE.1');
  igual(estadosRadar(e.resultado).filter((s) => s === 'vacío').length, 9, 'los otros nueve siguen vacíos');
  esperar(Validador.validar('RADAR', parcial).valido === true, 'el Validador no rechaza un componente incompleto');
  const completo = evaluarRadar(datosRadar(Object.assign(casoRE1(), { 'Movilización 2': [1, 3, 2, 4] })));
  igual(completo.resultado[1], { estado: 'completo', puntaje: 2.5 }, 'Movilización 2 completo: (1 + 3 + 2 + 4) / 4');
  igual(completosRadar(completo.resultado), [0, 1, 3, 5, 13], 'cinco puntos');
  igual([0, 3, 5, 13].map((k) => completo.resultado[k].puntaje), [1.5, 3.0, 0.0, 5.0], 'los cuatro anteriores no cambian');
});

prueba('RE.3', 'RADAR: el módulo entero vacío (sin calificar ninguna de las 56 características)', (esperar, igual) => {
  [datosRadar(), { calificaciones: new Array(56).fill('') }].forEach((datos, caso) => {
    const e = evaluarRadar(datos);
    igual(e.estado, 'vacia', 'caso ' + caso + ': estado');
    igual(e.errores, [], 'caso ' + caso + ': sin errores');
    esperar(e.resultado === undefined, 'caso ' + caso + ': no debía haber resultado ni gráfico');
  });
  const v = Validador.validar('RADAR', datosRadar());
  esperar(v.valido === true && v.errores.length === 0, 'el Validador no marca error por falta de calificaciones: ' + JSON.stringify(v.errores));
  igual(v.indicesInvalidos, [], 'ningún índice inválido');
  // Lo que calcularRadar devuelve si se le pasa el módulo vacío: 14 componentes vacíos.
  igual(estadosRadar(MotorCalculo.calcularRadar(datosRadar().calificaciones, [])), new Array(14).fill('vacío'), 'catorce componentes vacíos');
});

prueba('RE.4', 'RADAR: eje fijo de 0 a 5, sin ajustarse a los datos (distancias comprobadas en el SVG; lo visual, en Construction III)', (esperar, igual) => {
  // Un solo componente completo, con 1.50: debe quedar al 30 % del radio (1.5 / 5), no estirarse hasta el borde por ser el único.
  const solo = evaluarRadar(datosRadar({ 'Movilización 1': [0, 1, 2, 3] }));
  igual(completosRadar(solo.resultado), [0], 'un solo componente completo');
  const t = trazadoRadar(solo.resultado);
  igual(t.puntos.length, 1, 'un solo punto');
  igual(t.puntos[0].fraccion, 0.3, 'distancia al centro como fracción del radio');
  // Con el caso de RE.1: 0.30, 0.60, 0.00 (centro) y 1.00 (borde).
  const t1 = trazadoRadar(evaluarRadar(datosRadar(casoRE1())).resultado);
  igualLista(igual, t1.puntos.map((p) => p.fraccion), [0.3, 0.6, 0, 1], 'distancias del caso RE.1');
});

prueba('RE.5', 'RADAR: catorce puntas rotuladas con "<Etapa> <posición>" y puntos solo en los componentes completos (estructura del SVG; lo visual, en Construction III)', (esperar, igual) => {
  const resultado = evaluarRadar(datosRadar(casoRE1())).resultado;
  const t = trazadoRadar(resultado);
  igual(t.rayos.length, 14, 'catorce puntas');
  igual(t.rotulos, NOMBRES_RADAR, 'rótulos en el orden de las puntas 1 a 14');
  igual(t.puntos.length, 4, 'cuatro puntos, aunque las 14 puntas estén rotuladas');
  // Cada punto cae sobre la punta de su componente: Movilización 1, Traducción 1, Traducción 3 y Gestión 3.
  const sobreLaPunta = (p, k) => Math.abs(p.dx * t.rayos[k].dy - p.dy * t.rayos[k].dx) < 1e-6 * (1 + Math.hypot(t.rayos[k].dx, t.rayos[k].dy));
  [[0, 0], [1, 3], [3, 13]].forEach(([i, k]) => esperar(sobreLaPunta(t.puntos[i], k), 'el punto ' + (i + 1) + ' debía caer sobre la punta ' + (k + 1) + ' (' + NOMBRES_RADAR[k] + ')'));
  esperar(Math.hypot(t.puntos[2].dx, t.puntos[2].dy) < 1e-9, 'Traducción 3 (0.00) debía caer en el centro');
  // Un componente que no está completo no recibe punto.
  const conIncompleto = trazadoRadar(evaluarRadar(datosRadar(Object.assign(casoRE1(), { 'Movilización 2': [1, 3] }))).resultado);
  igual(conIncompleto.puntos.length, 4, 'un componente incompleto no recibe punto');
  const conInvalido = MotorCalculo.calcularRadar(datosRadar(casoRE1()).calificaciones, [0]);
  igual(trazadoRadar(conInvalido).puntos.length, 3, 'un componente inválido no recibe punto');
});

prueba('RE.6', 'RADAR: la lectura invertida está explícita (contenido del formulario y del SVG; la redacción y el aspecto, en Construction III)', (esperar) => {
  const formulario = appDOM.formularioRadar(datosRadar()).map(textoDeNodo).join(' ');
  ['Estoy completamente de acuerdo', 'Estoy bastante de acuerdo', 'Estoy algo de acuerdo', 'No estoy muy de acuerdo', 'No estoy casi nada de acuerdo', 'Estoy en completo desacuerdo']
    .forEach((nivel, n) => esperar(formulario.includes(n + ': ' + nivel), 'el formulario debía mostrar el nivel ' + n + ' de la escala'));
  esperar(/invertida/i.test(formulario) && /0 es el mejor/i.test(formulario) && /5 el peor/i.test(formulario), 'el formulario debía avisar que 0 es el mejor valor y 5 el peor');
  appDOM.MotorGraficos.dibujarRadar(evaluarRadar(datosRadar(casoRE1())).resultado);
  const grafico = textoDeNodo(appDOM.contenedor.hijos[0]);
  esperar(/invertida/i.test(grafico) && /se aleja del centro/i.test(grafico) && /problema/i.test(grafico), 'el gráfico debía avisar que un punto que se aleja del centro señala un problema');
});

prueba('RE.7', 'RADAR: la estructura es fija (56 campos, 14 componentes, 5 etapas, sin botones; lo visual, en Construction III)', (esperar, igual) => {
  const nodos = appDOM.formularioRadar(datosRadar()).flatMap(aplanarNodos);
  const campos = nodos.filter((n) => n.etiqueta === 'input').map((n) => n.atributos['data-campo']);
  igual(campos, Array.from({ length: 56 }, (_, i) => 'calificaciones.' + i), 'los 56 campos, en orden, con índice plano');
  igual(nodos.filter((n) => n.etiqueta === 'button' || n.atributos['data-accion'] !== undefined).length, 0, 'ningún botón ni acción para agregar o quitar');
  igual(nodos.filter((n) => n.etiqueta === 'select').length, 0, 'ningún selector que altere la estructura');
  const etapas = nodos.filter((n) => (n.className || '').includes('etapa-radar'));
  igual(etapas.map((e) => aplanarNodos(e).filter((n) => n.className === 'componente-radar').length), [3, 3, 2, 3, 3], 'componentes por etapa');
  const componentes = nodos.filter((n) => n.className === 'componente-radar');
  igual(componentes.map((c) => aplanarNodos(c).filter((n) => n.etiqueta === 'input').length), [4, 4, 4, 5, 4, 3, 4, 4, 4, 4, 4, 4, 4, 4], 'características por componente');
  // El encabezado de cada componente es "<nombre>: <título del profesor>"; la primera columna de cada tabla es "Afirmación" y cada fila "<n>. <texto>".
  const encabezados = componentes.map((c) => c.hijos[0].textContent);
  igual(encabezados.map((h) => h.slice(0, h.indexOf(':'))), NOMBRES_RADAR, 'nombre de cada componente al inicio del encabezado');
  esperar(encabezados.every((h) => h.slice(h.indexOf(':') + 1).trim().length > 0), 'los 14 componentes debían mostrar su título');
  igual(encabezados[0], 'Movilización 1: LA VISION, MISION Y ESTRATEGIA ESTÁN CLARAMENTE DEFINIDAS', 'encabezado de Movilización 1');
  igual(encabezados[13], 'Gestión 3: LA EMPRESA REALIZA UN SEGUIMIENTO SISTEMÁTICO DE LA GESTION ESTRATÉGICA', 'encabezado de Gestión 3');
  // Los dos componentes de Alineamiento comparten título en el documento del profesor ([VERIFICAR] de la Inception): aquí solo se comprueba que el
  // formulario los muestre distinguidos por su nombre.
  igual(encabezados[6].slice(encabezados[6].indexOf(':')), encabezados[7].slice(encabezados[7].indexOf(':')), 'Alineamiento 1 y 2 repiten el título del Excel');
  esperar(encabezados[6] !== encabezados[7], 'aun así sus encabezados debían diferir por el nombre');
  igual(componentes.map((c) => aplanarNodos(c).filter((n) => n.etiqueta === 'th').map((th) => th.textContent)), new Array(14).fill(['Afirmación', 'Calificación (0 a 5)']), 'encabezados de columna de cada tabla');
  const filasAfirmaciones = componentes.map((c) => aplanarNodos(c).filter((n) => n.etiqueta === 'tbody')[0].hijos.map((tr) => tr.hijos[0].hijos[0]));
  igual(filasAfirmaciones.map((f) => f.length), [4, 4, 4, 5, 4, 3, 4, 4, 4, 4, 4, 4, 4, 4], 'afirmaciones mostradas por componente');
  igual(filasAfirmaciones.flat().length, 56, 'las 56 afirmaciones están en el formulario');
  esperar(filasAfirmaciones.every((filas) => filas.every((texto, k) => texto.startsWith((k + 1) + '. ') && texto.length > 4)), 'cada fila debía empezar con su número dentro del componente y traer el texto');
  const estructura = appDOM.ESTRUCTURA_RADAR;
  igual(estructura.map((c) => c.afirmaciones.length), estructura.map((c) => c.caracteristicas), 'afirmaciones.length coincide con caracteristicas en los 14 componentes');
  igual(estructura.map((c) => c.afirmaciones.length), [4, 4, 4, 5, 4, 3, 4, 4, 4, 4, 4, 4, 4, 4], 'cantidad de afirmaciones por componente');
  // Texto exacto del profesor, transcrito sin corregir.
  igual(filasAfirmaciones[0][0], '1. La Estrategia está definida y formalizada por escrito', 'primera afirmación de Movilización 1');
  igual(filasAfirmaciones[3][4], '5. La Empresa tiene definidos el despliegue de sus objetivos a los niveles inferiores de la organizacion', 'quinta afirmación de Traducción 1');
  igual(filasAfirmaciones[7][0], '1. Los Gerentes programan reuniones periodicas para evaluar la información necesaria con sus unidades de soporte', 'primera afirmación de Alineamiento 2');
  igual(filasAfirmaciones[13][3], '4. La empresa tiene una reunión anual de redefinición del la Estrategia', 'cuarta afirmación de Gestión 3 (con la errata del Excel)');
  const tablaResultados = appDOM.resultadosRadar(MotorCalculo.calcularRadar(datosRadar().calificaciones, [])).flatMap(aplanarNodos);
  igual(tablaResultados.filter((n) => n.etiqueta === 'tr').length, 15, 'la tabla de resultados tiene el encabezado y 14 filas');
});

prueba('RE.8', 'RADAR: los dos componentes de Alineamiento (mismo título en el documento) son independientes', (esperar, igual) => {
  const soloSegundo = evaluarRadar(datosRadar({ 'Alineamiento 2': [3, 3, 3, 3] }));
  igual(soloSegundo.resultado[7], { estado: 'completo', puntaje: 3 }, 'Alineamiento 2 completo en 3.00 (punta 8)');
  igual(soloSegundo.resultado[6], { estado: 'vacío', puntaje: null }, 'Alineamiento 1 vacío (punta 7)');
  igual(completosRadar(soloSegundo.resultado), [7], 'solo la punta 8 tiene punto');
  const ambos = evaluarRadar(datosRadar({ 'Alineamiento 2': [3, 3, 3, 3], 'Alineamiento 1': [1, 1, 1, 1] }));
  igual(ambos.resultado[6], { estado: 'completo', puntaje: 1 }, 'Alineamiento 1 completo en 1.00');
  igual(ambos.resultado[7], { estado: 'completo', puntaje: 3 }, 'Alineamiento 2 sigue en 3.00');
});

prueba('RE.9', 'RADAR: editar o borrar una calificación recalcula o devuelve el componente a incompleto', (esperar, igual) => {
  const datos = datosRadar(casoRE1());
  datos.calificaciones[3] = '5';
  const editado = evaluarRadar(datos);
  igual(editado.resultado[0], { estado: 'completo', puntaje: 2 }, 'Movilización 1: (0 + 1 + 2 + 5) / 4');
  igual([3, 5, 13].map((k) => editado.resultado[k].puntaje), [3.0, 0.0, 5.0], 'los otros tres puntos no cambian');
  [null, ''].forEach((sinValor) => {
    const borrado = datosRadar(casoRE1());
    borrado.calificaciones[0] = sinValor;
    const e = evaluarRadar(borrado);
    const caso = 'borrar con ' + JSON.stringify(sinValor);
    igual(e.estado, 'ok', caso + ': el módulo sigue calculando');
    igual(e.errores, [], caso + ': sin errores');
    igual(e.resultado[0], { estado: 'incompleto', puntaje: null }, caso + ': Movilización 1 incompleto');
    igual(completosRadar(e.resultado), [3, 5, 13], caso + ': los otros tres puntos');
    igual(borrado.calificaciones.slice(1, 4), ['1', '2', '3'], caso + ': las otras tres calificaciones se conservan');
  });
});

prueba('RE.10', 'RADAR: calificación fuera de rango (el componente queda inválido y los demás no se ven afectados)', (esperar, igual) => {
  [6, -1].forEach((valor) => {
    const datos = datosRadar(casoRE1());
    datos.calificaciones[3] = String(valor);
    const v = Validador.validar('RADAR', datos);
    const caso = 'valor ' + valor;
    esperar(v.valido === false, caso + ': el Validador debía rechazar la calificación');
    igual(v.indicesInvalidos, [3], caso + ': índice plano de la característica fuera de rango');
    igual(v.errores.length, 1, caso + ': un solo mensaje, por característica');
    const e = evaluarRadar(datos);
    igual(e.estado, 'ok', caso + ': el módulo no se bloquea');
    esperar(e.errores.length === 1 && e.errores[0].startsWith('Movilización 1, característica 4:') && e.errores[0].includes('0') && e.errores[0].includes('5'),
      caso + ': el error debía nombrar el componente y la posición e indicar el rango de 0 a 5: ' + JSON.stringify(e.errores));
    igual(e.resultado[0], { estado: 'inválido', puntaje: null }, caso + ': Movilización 1 inválido, sin puntaje');
    igual([3, 5, 13].map((k) => e.resultado[k].puntaje), [3.0, 0.0, 5.0], caso + ': los otros tres componentes siguen igual');
    igual(datos.calificaciones[3], String(valor), caso + ': el valor escrito se conserva');
  });
  // Un valor fuera de rango con el resto del componente sin calificar sigue siendo inválido (no incompleto).
  igual(evaluarRadar(datosRadar({ 'Movilización 1': [9] })).resultado[0].estado, 'inválido', 'inválido aunque el resto del componente esté sin calificar');
  // Un valor no numérico es una entrada inválida, no una celda vacía.
  const texto = datosRadar(casoRE1());
  texto.calificaciones[3] = 'abc';
  igual(evaluarRadar(texto).resultado[0].estado, 'inválido', 'texto no numérico');
  // Los dos extremos de la escala son válidos.
  esperar(Validador.validar('RADAR', datosRadar(casoRE1())).valido === true, '0 y 5 son calificaciones válidas');
});

prueba('RE.11', 'RADAR: los datos siguen tras recargar (persistencia y recálculo)', (esperar, igual) => {
  const almacen = {};
  const localStorageSimulado = {
    getItem: (k) => (k in almacen ? almacen[k] : null), setItem: (k, v) => { almacen[k] = String(v); }, removeItem: (k) => { delete almacen[k]; }
  };
  const app2 = cargarApp({ localStorage: localStorageSimulado });
  const p = app2.Persistencia;
  const inicial = p.cargar();
  igual(inicial.datos.RADAR, { calificaciones: new Array(56).fill(null) }, 'RADAR con 56 null en un estado nuevo');
  const datos = datosRadar(Object.assign(casoRE1(), { 'Movilización 2': [1, 3] }));
  inicial.matrizActiva = 'RADAR';
  inicial.datos.RADAR = datos;
  esperar(p.guardar(inicial) === true, 'guardar debía devolver true');
  // "Recargar": otra instancia de la aplicación lee lo guardado.
  const app3 = cargarApp({ localStorage: localStorageSimulado });
  const recuperado = app3.Persistencia.cargar();
  igual(recuperado.matrizActiva, 'RADAR', 'matriz activa recuperada');
  igual(recuperado.datos.RADAR, datos, 'las 56 calificaciones recuperadas, incluidas las dos del componente incompleto');
  app3.fijarEstado(recuperado);
  const e = app3.evaluarMatriz('RADAR');
  igual(e.estado, 'ok', 'estado tras recargar');
  igual(completosRadar(e.resultado), [0, 3, 5, 13], 'los mismos cuatro puntos de RE.1');
  igual([0, 3, 5, 13].map((k) => e.resultado[k].puntaje), [1.5, 3.0, 0.0, 5.0], 'los mismos puntajes de RE.1');
  igual(e.resultado[1], { estado: 'incompleto', puntaje: null }, 'Movilización 2 sigue incompleto, sin punto');

  // Un estado guardado antes del módulo no trae datos.RADAR: se completa con 56 null y los demás datos sobreviven, sin cambiar la versión.
  const viejo = app3.Persistencia.cargar();
  viejo.datos.BCG.divisiones[0].ingresos = '500';
  delete viejo.datos.RADAR;
  almacen['mtx.estado'] = JSON.stringify(viejo);
  const migrado = app3.Persistencia.cargar();
  igual(migrado.datos.RADAR, { calificaciones: new Array(56).fill(null) }, 'estado anterior sin RADAR: 56 null');
  igual(migrado.datos.BCG.divisiones[0].ingresos, '500', 'estado anterior sin RADAR: los demás datos se conservan');
  igual(migrado.version, 1, 'la versión del formato guardado no cambia');
  // Un RADAR dañado (cantidad distinta de 56) también se reemplaza por 56 null.
  const danado = app3.Persistencia.cargar();
  danado.datos.RADAR = { calificaciones: ['1', '2'] };
  almacen['mtx.estado'] = JSON.stringify(danado);
  igual(app3.Persistencia.cargar().datos.RADAR, { calificaciones: new Array(56).fill(null) }, 'RADAR dañado: 56 null');
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

// Exportador de AE (decisión 3 de Construction II del Módulo 2). Estructura provisional [VERIFICAR]: no se pudo comparar con el Excel original.
prueba('X.6', 'Exportador AE: hojas "Datos" (matriz completa) y "Resultados" (una fila por variable y los cortes)', (esperar, igual) => {
  const contexto = vm.createContext({ console });
  vm.runInContext(scriptPorId('sheetjs'), contexto);
  let capturado = null;
  contexto.XLSX.writeFile = (libro, nombre) => { capturado = { libro, nombre }; };
  const appConXLSX = cargarApp({ XLSX: contexto.XLSX });
  const datos = datosAE(matrizAE1(), ['SI', 'NO', 'NO', 'SI']);
  const resultado = MotorCalculo.calcularAE(datos.variables.map((nombre, i) => ({ nombre, marca: datos.marcas[i] })), datos.matriz);
  igual(appConXLSX.Exportador.exportarXLSX('AE', datos, resultado), 'Mtx-AE.xlsx', 'nombre del archivo');
  esperar(capturado !== null && capturado.nombre === 'Mtx-AE.xlsx', 'debía llamar a XLSX.writeFile con Mtx-AE.xlsx');
  igual(capturado.libro.SheetNames, ['Datos', 'Resultados'], 'hojas');
  const hojaDatos = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Datos, { header: 1 });
  const hojaResultados = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Resultados, { header: 1 });
  igual(hojaDatos[0], ['Variable', 'V1', 'V2', 'V3', 'V4'], 'encabezado de columnas con los nombres de variable');
  igual(hojaDatos.length, 5, 'filas de Datos (encabezado + 4 variables)');
  igual(hojaDatos.map((f) => f[0]).slice(1), ['V1', 'V2', 'V3', 'V4'], 'encabezado de filas con los nombres de variable');
  matrizAE1().forEach((filaEsperada, i) => filaEsperada.forEach((valor, j) => {
    const real = hojaDatos[i + 1][j + 1];
    if (valor === null) esperar(real === undefined || real === null, 'la diagonal debía quedar vacía en (' + i + ', ' + j + ')');
    else igual(real, valor, 'celda (' + i + ', ' + j + ')');
  }));
  igual(hojaResultados[0], ['Variable', 'Motricidad', 'Dependencia', 'Cuadrante', 'Proyección x', 'Proyección y', 'Marca'], 'encabezado de Resultados');
  igual(hojaResultados[1].slice(0, 4), ['V1', 8, 1, 'INDEPENDIENTES'], 'fila de V1');
  igual(hojaResultados[1][6], 'SÍ', 'marca de V1');
  igual(hojaResultados.slice(1, 5).map((f) => f[0]), ['V1', 'V2', 'V3', 'V4'], 'variables en orden de carga');
  igual(hojaResultados.slice(1, 5).map((f) => f[6]), ['SÍ', 'NO', 'NO', 'SÍ'], 'marcas');
  igualLista(igual, hojaResultados.slice(1, 5).map((f) => f[4]), [3.5, 0, -0.5, -3], 'proyección x');
  const corteY = hojaResultados.find((f) => f[0] === 'Corte Y');
  const corteX = hojaResultados.find((f) => f[0] === 'Corte X');
  esperar(corteY && corteY[1] === 4 && corteX && corteX[1] === 4, 'las filas finales debían traer Corte Y = 4 y Corte X = 4');
});

// Exportador de RADAR (decisiones F y G de Construction II del Módulo 3): "Datos" con las 56 calificaciones y "Resultados" con los 14 componentes
// y su estado real, y se exporta aunque haya componentes incompletos, vacíos o inválidos.
prueba('X.7', 'Exportador RADAR: hojas "Datos" (56 filas) y "Resultados" (14 filas con el estado real, inválido incluido)', (esperar, igual) => {
  const contexto = vm.createContext({ console });
  vm.runInContext(scriptPorId('sheetjs'), contexto);
  let capturado = null;
  contexto.XLSX.writeFile = (libro, nombre) => { capturado = { libro, nombre }; };
  const appConXLSX = cargarApp({ XLSX: contexto.XLSX });
  // Movilización 1 inválido (un 6), Movilización 2 incompleto, Traducción 1 y Gestión 3 completos, el resto vacío.
  const datos = datosRadar({ 'Movilización 1': [0, 1, 2, 6], 'Movilización 2': [1, 3], 'Traducción 1': [0, 0, 5, 5, 5], 'Gestión 3': [5, 5, 5, 5] });
  const e = evaluarRadar(datos);
  igual(e.estado, 'ok', 'la exportación no se bloquea por componentes incompletos o inválidos');
  igual(appConXLSX.Exportador.exportarXLSX('RADAR', e.datos, e.resultado), 'Mtx-RADAR.xlsx', 'nombre del archivo');
  esperar(capturado !== null && capturado.nombre === 'Mtx-RADAR.xlsx', 'debía llamar a XLSX.writeFile con Mtx-RADAR.xlsx');
  igual(capturado.libro.SheetNames, ['Datos', 'Resultados'], 'hojas');
  const hojaDatos = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Datos, { header: 1 });
  const hojaResultados = contexto.XLSX.utils.sheet_to_json(capturado.libro.Sheets.Resultados, { header: 1 });
  igual(hojaDatos[0], ['Etapa', 'Componente', 'Característica', 'Calificación'], 'encabezado de Datos');
  igual(hojaDatos.length, 57, 'filas de Datos (encabezado + 56 características)');
  igual(hojaDatos[1], ['Movilización', 'Movilización 1', 'Característica 1', 0], 'primera fila de Datos');
  igual(hojaDatos[4][3], 6, 'el valor fuera de rango se exporta tal cual');
  igual(hojaDatos[7][3], undefined, 'una característica sin calificar queda vacía');
  igual(hojaDatos[56], ['Gestión', 'Gestión 3', 'Característica 4', 5], 'última fila de Datos');
  igual(hojaResultados[0], ['Etapa', 'Componente', 'Estado', 'Puntaje'], 'encabezado de Resultados');
  igual(hojaResultados.length, 15, 'filas de Resultados (encabezado + 14 componentes)');
  igual(hojaResultados[1].slice(0, 3), ['Movilización', 'Movilización 1', 'Inválido'], 'Movilización 1 inválido');
  igual(hojaResultados[2].slice(0, 3), ['Movilización', 'Movilización 2', 'Incompleto'], 'Movilización 2 incompleto');
  igual(hojaResultados[3].slice(0, 3), ['Movilización', 'Movilización 3', 'Vacío'], 'Movilización 3 vacío');
  igual(hojaResultados[4], ['Traducción', 'Traducción 1', 'Completo', 3], 'Traducción 1 completo con su puntaje');
  igual(hojaResultados[14], ['Gestión', 'Gestión 3', 'Completo', 5], 'Gestión 3 completo con su puntaje');
  igual(hojaResultados.slice(1).map((f) => f[3]).filter((p) => p !== undefined).length, 2, 'solo los componentes completos traen puntaje');
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
