# Mtx: Documento VUP (Vibe Unified Process)

Estado: borrador de la fase Inception, pendiente de confirmación de Gerson antes de pasar a Elaboration I.
Referencia del proceso: https://vibeprocess.org/ (Vibe Unified Process, Dr. Michael Dorin, University of St. Thomas), adaptado de Jacobson, Booch y Rumbaugh (1999), El Proceso Unificado de Desarrollo de Software.

Flujo de trabajo acordado: este documento y el código viven en esta carpeta local. Claude Code trabaja aquí como constructor (fases de Elaboration II en adelante). La sesión de Cowork del Project "Prompt Engineering" en claude.ai actúa como juez, lee este mismo documento y el código a través del puente con la computadora, aplica los prompts de revisión de rol de VUP (Product Owner, QA, Arquitecto, Desarrollador) y documenta qué se acepta o se rechaza antes de desbloquear la siguiente fase.

## Inception

### Nombre del proyecto
Mtx

### Declaración de visión

Para los estudiantes y profesores del curso de Gestión Estratégica, que necesitan aplicar las matrices clásicas de planeamiento estratégico sin depender de una instalación frágil de Excel con macros VBA y bases de datos Access, Mtx es una suite de calculadoras estratégicas interactivas que corre íntegramente en el navegador. Calcula y grafica cada matriz al instante, sin instalación ni configuración previa. A diferencia de las plantillas actuales en Excel VBA más Access, Mtx elimina toda dependencia de software de escritorio, macros bloqueadas por seguridad de Windows, y bases de datos que se rompen al mover la carpeta a OneDrive o SharePoint.

### Fuera de alcance (v1)

- Replicar el modelo PE-BSC completo (34 hojas: misión, visión, EFI, EFE, objetivos estratégicos, análisis FLOR, ADN de misión y visión, Balanced Scorecard) queda para una fase posterior.
- Análisis Estructural (estilo MICMAC) ya se inició como Módulo 2 (ver la sección "Inception — Módulo 2: Análisis Estructural" al final de este documento). Radar Estratégico ya se inició como Módulo 3 (ver la sección "Inception — Módulo 3: Radar Estratégico" al final de este documento). Solo Priorización de Iniciativas sigue pendiente para fases posteriores.
- Persistencia compartida entre varios integrantes de un mismo grupo trabajando a la vez no se resuelve en esta primera versión.
- Generación de conclusiones asistida por IA, al estilo de Praxio, se evalúa en una fase posterior, no en el primer entregable.

### Justificación

El sistema actual depende de Excel VBA más Access, lo que genera fallas documentadas en el propio manual técnico del profesor: macros bloqueadas por seguridad la primera vez que el archivo llega por internet o WhatsApp, rutas rotas si la carpeta se mueve a OneDrive o SharePoint, necesidad de tener instalado el motor de Access del mismo bitness que Office, y solo un usuario a la vez por archivo porque Access no soporta escritura simultánea. Migrar a una suite HTML autónoma elimina esa fricción sin costo de licencias, mejora la experiencia de los alumnos del curso de Gestión Estratégica, y sirve como caso de estudio propio de ingeniería de software asistida por IA.

### Historias de usuario (v1, módulo Matrices de Combinación)

1. Como estudiante, quiero cargar los ingresos y utilidades de cada división de mi empresa y obtener automáticamente el gráfico BCG, para no depender de una hoja de Excel con macros.
2. Como estudiante, quiero calificar mis factores internos y externos con peso y clasificación, y obtener el puntaje ponderado de EFI y EFE automáticamente, para no sumar a mano.
3. Como estudiante, quiero comparar mi empresa contra varios competidores en una matriz de perfil competitivo (MPC), para visualizar mi posición relativa.
4. Como estudiante, quiero ingresar los valores de fuerza financiera, ventaja competitiva, fuerza de la industria y estabilidad del entorno, para obtener automáticamente el vector y el cuadrante de la matriz PEYEA.
5. Como estudiante, quiero que el sistema ubique automáticamente mi empresa en la matriz interna-externa (MIE) a partir de mis totales de EFI y EFE, para saber en qué celda cae.
6. Como estudiante, quiero ver la matriz de la gran estrategia con mi posición marcada, para identificar qué tipo de estrategias me corresponden.
7. Como profesor, quiero que mis alumnos puedan usar estas siete matrices sin instalar nada ni depender de un laboratorio con Office y Access configurados.

### Requisitos no funcionales

- Todos los cálculos se ejecutan en el navegador, sin backend ni servidor propio.
- El archivo abre directamente haciendo doble clic o desde un enlace, sin instalación.
- Los resultados y gráficos deben coincidir exactamente con las fórmulas clásicas de cada matriz (BCG, PEYEA, EFI, EFE, MPC, MIE, GE), verificables contra un caso de bibliografía conocido.
- La interfaz debe ser usable por un estudiante sin conocimientos técnicos, replicando o mejorando la simplicidad guiada del Excel original (qué se escribe a mano y qué sale solo).
- El progreso del usuario se guarda automáticamente en el navegador (localStorage) para no perderse si cierra la pestaña.

### Riesgos de desarrollo

- Terminología ambigua en el modelo PE-BSC original (análisis FLOR, ADN de misión y visión) que no corresponde a bibliografía estándar reconocible y requiere confirmación del profesor antes de replicarla en fases posteriores.
- Las bases de datos Access originales del profesor fueron reconstruidas y arrancan vacías, sin datos de ejemplo ya resueltos contra los cuales validar los cálculos; hay que construir casos de prueba propios a partir de las fórmulas documentadas y de la bibliografía clásica de gestión estratégica.
- Fidelidad a la terminología exacta de D'Alessio (EFI, EFE, PEYEA, MPC, MIE, GE), dado que probablemente se usa para calificar en el curso.

### URLs de referencia

- https://github.com/hcornejovillena-cmd/praxio (referencia arquitectónica: aplicación HTML única, sin backend, flujo por módulo de carga, validación, análisis, resultados, exportación y conclusiones asistidas por IA)

## Elaboration I

Objetivo de la fase: definir, para cada historia de usuario de Inception, al menos una prueba de aceptación concreta en formato Given-When-Then (Dado / Cuando / Entonces), con valores numéricos y resultado esperado exacto. Todos los resultados esperados fueron recalculados con un script aparte antes de escribirse aquí.

Convenciones generales de las pruebas:

- Los valores numéricos se comparan con dos decimales, salvo donde se indique otra cosa.
- Las marcas [VERIFICAR] señalan una convención que las fórmulas estándar no fijan de forma única y que debe confirmarse con el profesor o el material del curso antes de la fase de construcción. Se consolidan al final de esta sección.
- Las pruebas de las historias 2, 3 y 5 comparten datos a propósito: el EFI (2.45) y el EFE (2.90) de las pruebas 2.1 y 2.2 son las entradas de la prueba 5.1, para verificar la coherencia entre matrices.

### Historia 1: BCG

Hallazgo sobre la historia: el texto solo menciona "ingresos y utilidades", pero el eje X (participación relativa de mercado) y el eje Y (tasa de crecimiento del mercado) necesitan sus propios datos de entrada. La prueba los incluye como datos que el estudiante ingresa. Si el curso pide que la herramienta calcule la participación relativa a partir de ventas de competidores, la historia debe ampliarse [VERIFICAR].

**Prueba 1.1: cuadrantes y tamaño de burbuja**

- Dado que el estudiante carga cuatro divisiones con estos datos (ingresos y utilidades en millones):

  | División | Ingresos | Utilidades | Participación relativa (X) | Crecimiento del mercado (Y) |
  |---|---|---|---|---|
  | A | 500 | 100 | 1.80 | 15 % |
  | B | 300 | 30 | 0.40 | 12 % |
  | C | 150 | 60 | 1.50 | 4 % |
  | D | 50 | 10 | 0.30 | 2 % |

- Cuando calcula el gráfico BCG usando los umbrales participación relativa = 1.0 y crecimiento = 10 % [VERIFICAR]
- Entonces se cumple todo lo siguiente:
  - A cae en Estrella, B en Interrogante, C en Vaca lechera y D en Perro.
  - Los ingresos totales son 1 000 y las utilidades totales son 200.
  - El tamaño de burbuja por % de ingresos es A = 50 %, B = 30 %, C = 15 %, D = 5 % (suma 100 %).
  - El tamaño de burbuja por % de utilidades es A = 50 %, B = 15 %, C = 30 %, D = 5 % (suma 100 %).

Puntos sin definir en la historia, no probados hasta que se decidan [VERIFICAR]:

- Qué hace la burbuja cuando una división tiene utilidad negativa o cero (el tamaño no puede ser negativo).
- A qué cuadrante pertenece una división con participación relativa exactamente 1.0 o con crecimiento exactamente igual al umbral.
- Si el umbral de crecimiento es 10 % fijo, es configurable, o sigue la convención de D'Alessio, que centra el eje en 0 %. Lo mismo para si el eje X va en escala logarítmica.

### Historia 2: EFI y EFE

**Prueba 2.1: EFI débil**

- Dado que el estudiante ingresa estos factores internos (peso, clasificación):
  - Fortalezas: F1 (0.20, 4), F2 (0.15, 4), F3 (0.10, 3).
  - Debilidades: D1 (0.25, 1), D2 (0.20, 2), D3 (0.10, 1).
- Cuando el sistema calcula el EFI
- Entonces se cumple todo lo siguiente:
  - La suma de pesos es 1.00.
  - Los ponderados son 0.80, 0.60, 0.30, 0.25, 0.40 y 0.10.
  - El total EFI es 2.45.
  - Como 2.45 < 2.5, el diagnóstico es "posición interna débil".

**Prueba 2.2: EFE que aprovecha oportunidades**

- Dado que el estudiante ingresa estos factores externos (peso, clasificación):
  - Oportunidades: O1 (0.25, 4), O2 (0.20, 3), O3 (0.10, 2).
  - Amenazas: A1 (0.25, 3), A2 (0.15, 2), A3 (0.05, 1).
- Cuando el sistema calcula el EFE
- Entonces se cumple todo lo siguiente:
  - La suma de pesos es 1.00.
  - Los ponderados son 1.00, 0.60, 0.20, 0.75, 0.30 y 0.05.
  - El total EFE es 2.90.
  - Como 2.90 > 2.5, el diagnóstico es "la organización aprovecha oportunidades y evita amenazas por encima del promedio".

**Prueba 2.3: pesos que no suman 1**

- Dado el EFI de la prueba 2.1 en el que se cambia el peso de D3 de 0.10 a 0.05, de modo que la suma de pesos pasa a 0.95
- Cuando el estudiante intenta ver el resultado
- Entonces el sistema muestra que la suma es 0.95, indica que debe ser 1.00 y no presenta el total ni el diagnóstico como resultado válido. (El total aritmético sería 2.40. La redacción exacta del mensaje se define en el diseño.)

**Prueba 2.4: tolerancia de punto flotante en la suma de pesos**

- Dado un EFE con tres factores de pesos 0.6, 0.3 y 0.1 (en JavaScript, 0.6 + 0.3 + 0.1 da 0.9999999999999999)
- Cuando el sistema valida que los pesos suman 1
- Entonces la validación acepta la suma como 1.00. Hace falta una tolerancia, por ejemplo |suma − 1| < 0.001, y no una comparación con igualdad estricta.

Puntos sin definir [VERIFICAR]:

- Qué diagnóstico se muestra cuando el total es exactamente 2.5 (la regla dada solo cubre menor y mayor).
- Si el curso restringe la clasificación por tipo de factor. D'Alessio usa 3 o 4 para fortalezas y 1 o 2 para debilidades en el EFI, y 1 a 4 libre en el EFE. La regla dada dice 1 a 4 para todos; la prueba 2.1 cumple ambas lecturas.
- Si se rechaza una clasificación fuera de 1 a 4 (por ejemplo 5) o un peso fuera de 0 a 1. Se asume que sí, pero no se fija con una prueba numérica.

### Historia 3: MPC

**Prueba 3.1: empresa propia frente a dos competidores**

- Dado que los factores críticos y sus pesos son los mismos para todos: Participación de mercado 0.30, Competitividad de precios 0.25, Calidad del producto 0.25, Lealtad del cliente 0.20 (suma 1.00), y estas clasificaciones (en el orden de los factores):
  - Mi empresa: 3, 2, 4, 3.
  - Competidor A: 4, 3, 3, 2.
  - Competidor B: 2, 4, 2, 4.
- Cuando el sistema calcula el MPC
- Entonces se cumple todo lo siguiente:
  - Mi empresa obtiene 0.90 + 0.50 + 1.00 + 0.60 = 3.00.
  - El Competidor A obtiene 1.20 + 0.75 + 0.75 + 0.40 = 3.10.
  - El Competidor B obtiene 0.60 + 1.00 + 0.50 + 0.80 = 2.90.
  - El orden de mayor a menor es A (3.10), Mi empresa (3.00), B (2.90), por lo que mi empresa queda en 2.º lugar de 3.

**Prueba 3.2: los pesos se comparten entre todas las empresas**

- Dado el MPC de la prueba 3.1
- Cuando el estudiante cambia los pesos a 0.10, 0.50, 0.20 y 0.20 (suma 1.00) una sola vez
- Entonces las tres columnas se recalculan con los nuevos pesos, sin que el estudiante tenga que reescribirlos por empresa: Mi empresa 2.70, Competidor A 2.90 y Competidor B 3.40. El orden pasa a ser B, A, Mi empresa, por lo que mi empresa queda en 3.er lugar.

Puntos sin definir [VERIFICAR]:

- Cantidad mínima y máxima de competidores admitidos (el Excel original podría limitarla).
- Si el curso también valida la clasificación 1 a 4 en el MPC como en EFI y EFE.
- Cómo se muestran los empates (ninguna prueba los cubre).

### Historia 4: PEYEA

Convención de cálculo usada: cada eje es el promedio simple de sus factores. Fuerza financiera (FF) y fuerza de la industria (FI) se califican de 1 (peor) a 6 (mejor); ventaja competitiva (VC) y estabilidad del entorno (EE) se califican de −1 (mejor) a −6 (peor). Ninguna de las cuatro variables admite el valor 0, siguiendo la convención de D'Alessio y Rowe. Eje X = ventaja competitiva + fuerza de la industria. Eje Y = estabilidad del entorno + fuerza financiera. Cuadrantes: agresivo (X>0, Y>0), conservador (X<0, Y>0), defensivo (X<0, Y<0), competitivo (X>0, Y<0).

**Prueba 4.1: perfil agresivo**

- Dado que el estudiante califica: Fuerza financiera (FF) = 5, 4, 4, 3; Fuerza de la industria (FI) = 4, 5, 3; Ventaja competitiva (VC) = −2, −3, −1, −2; Estabilidad del entorno (EE) = −3, −2, −4, −3, −3
- Cuando el sistema calcula la matriz PEYEA
- Entonces los promedios son FF = 4.00, FI = 4.00, VC = −2.00, EE = −3.00. Por tanto X = −2.00 + 4.00 = 2.00 e Y = −3.00 + 4.00 = 1.00. El vector termina en (2.00, 1.00) y el cuadrante es agresivo.

**Prueba 4.2: perfil conservador**

- Dado FF = 5, 5, 5; FI = 1, 2; VC = −4, −4; EE = −2, −1, −3
- Cuando el sistema calcula la matriz PEYEA
- Entonces FF = 5.00, FI = 1.50, VC = −4.00, EE = −2.00, X = −2.50, Y = 3.00 y el cuadrante es conservador.

**Prueba 4.3: perfil defensivo**

- Dado FF = 1, 2; FI = 1, 1; VC = −4, −4, −4; EE = −5, −5
- Cuando el sistema calcula la matriz PEYEA
- Entonces FF = 1.50, FI = 1.00, VC = −4.00, EE = −5.00, X = −3.00, Y = −3.50 y el cuadrante es defensivo.

**Prueba 4.4: perfil competitivo**

- Dado FF = 1, 2, 3; FI = 5, 5, 5; VC = −1, −2; EE = −5, −5, −5
- Cuando el sistema calcula la matriz PEYEA
- Entonces FF = 2.00, FI = 5.00, VC = −1.50, EE = −5.00, X = 3.50, Y = −3.00 y el cuadrante es competitivo.

Puntos sin definir [VERIFICAR]:

- Vector: la historia pide "el vector". Se asume que el vector es el segmento del origen a (X, Y). Si el curso exige magnitud y ángulo, para la prueba 4.1 serían √5 = 2.24 y 26.57° (atan2(1, 2)).
- Cuadrante cuando X = 0 o Y = 0: no está definido por la regla dada.
- Cantidad de factores por eje (las pruebas usan de 2 a 5 factores por eje, sin asumir un número fijo).

### Historia 5: MIE

Convención: el total de EFI (eje X) y el total de EFE (eje Y) se clasifican como débil (1.0 a menos de 2.0), promedio (2.0 a menos de 3.0) o fuerte (3.0 a 4.0). Numeración de celdas en romano, de I a IX, con EFE fuerte arriba y EFI fuerte a la izquierda [VERIFICAR]. Zonas: crecer y construir = I, II, IV; retener y mantener = III, V, VII; cosechar o desinvertir = VI, VIII, IX.

**Prueba 5.1: celda central con los totales de las pruebas 2.1 y 2.2**

- Dado EFI = 2.45 y EFE = 2.90
- Cuando el sistema ubica la empresa en la MIE
- Entonces el EFI es "promedio", el EFE es "promedio", la celda es V y la zona es "retener y mantener".

**Prueba 5.2: esquina de crecimiento**

- Dado EFI = 3.20 y EFE = 3.50
- Cuando el sistema ubica la empresa en la MIE
- Entonces ambos son "fuerte", la celda es I y la zona es "crecer y construir".

**Prueba 5.3: esquina de cosecha**

- Dado EFI = 1.80 y EFE = 1.50
- Cuando el sistema ubica la empresa en la MIE
- Entonces ambos son "débil", la celda es IX y la zona es "cosechar o desinvertir".

**Prueba 5.4: valores en el límite de rango**

- Dado un EFE fijo de 3.50 (fuerte) y estos valores de EFI: 1.99, 2.00, 2.99 y 3.00
- Cuando el sistema ubica la empresa en la MIE con cada uno
- Entonces la celda es III para EFI 1.99 (débil), II para 2.00 (promedio), II para 2.99 (promedio) e I para 3.00 (fuerte).

Puntos sin definir [VERIFICAR]:

- Qué se hace con un valor entre 1.99 y 2.00, por ejemplo 1.995. Los rangos dados dejan un hueco. La prueba 5.4 asume cortes en < 2.0 y < 3.0, lo que equivale a los rangos dados con valores de dos decimales.
- Numeración de celdas y asignación de zonas: usadas aquí según David y D'Alessio, sin confirmación del curso.

### Historia 6: Gran Estrategia (GE)

La historia no admite una prueba de extremo a extremo con la información actual. Pide "ver la matriz con mi posición marcada", pero ni la historia ni el resto de VUP.md definen de dónde salen los dos valores: si el estudiante los ingresa directamente, en qué escala, cuál es el punto que separa rápido de lento y fuerte de débil, o si se derivan de otras matrices (por ejemplo del EFE y del EFI, o de los ejes de la PEYEA). Sin eso, cualquier valor de entrada con resultado "exacto" sería inventado.

Lo que sí se puede verificar hoy es la correspondencia de cada cuadrante con sus estrategias sugeridas, según la Matriz de la Gran Estrategia clásica [VERIFICAR contra el material del curso, porque las listas varían entre autores].

**Prueba 6.1: correspondencia cuadrante y estrategias (tabla de búsqueda)**

- Dado que el sistema clasifica la empresa en un cuadrante
- Cuando el estudiante consulta el resultado
- Entonces el cuadrante determina el tipo de estrategia sugerida:
  - Crecimiento rápido con posición fuerte (cuadrante I): estrategias intensivas, integración y diversificación concéntrica.
  - Crecimiento rápido con posición débil (cuadrante II): estrategias intensivas, integración horizontal, desinversión, liquidación.
  - Crecimiento lento con posición débil (cuadrante III): reducción, diversificación conglomerada, desinversión, liquidación.
  - Crecimiento lento con posición fuerte (cuadrante IV): diversificación concéntrica, horizontal y conglomerada, empresas conjuntas.

Decisiones necesarias antes de escribir la prueba de ubicación (bloquean la prueba de extremo a extremo de la historia 6):

1. ¿Los dos ejes los ingresa el estudiante o se derivan de otras matrices?
2. ¿En qué escala están, y cuál es el punto que separa rápido de lento y fuerte de débil?
3. ¿Qué pasa cuando un valor cae exactamente en el punto de corte?

### Historia 7: uso sin instalación (profesor)

Esta historia es un criterio de aceptación de despliegue, no de cálculo, y por eso su prueba no lleva números de negocio. Es concreta y comprobable, pero necesita un equipo real para ejecutarse.

**Prueba 7.1: abrir y usar las siete matrices en un equipo limpio**

- Dado un equipo con navegador moderno, sin Microsoft Office, sin motor de Access y sin conexión a un servidor propio (por ejemplo, la carpeta copiada a un pendrive o descargada de un enlace)
- Cuando el profesor o un alumno abre el archivo principal con doble clic (protocolo `file://`) y usa las siete matrices
- Entonces se cumple todo lo siguiente:
  - No aparece ningún aviso de macros, permisos ni instalación de componentes.
  - Las siete matrices (BCG, EFI, EFE, MPC, PEYEA, MIE, GE) están accesibles desde la interfaz.
  - Los datos de las pruebas 1.1, 2.1, 2.2, 3.1, 4.1, 5.1 y 5.2 producen exactamente los resultados esperados descritos arriba.
  - La consola del navegador no registra errores.
  - La pestaña de red no registra solicitudes a un backend propio.
  - Al recargar la página, los datos ingresados siguen ahí (localStorage).

Puntos sin definir [VERIFICAR]:

- Navegadores y versiones mínimas que el curso debe soportar.
- Resuelto en Elaboration II: todo el código, incluida cualquier librería como la del Exportador, va embebido en el mismo archivo, sin depender de un CDN, para que funcione sin conexión a internet.

### Resumen de la fase

| Historia | Pruebas | Estado |
|---|---|---|
| 1. BCG | 1.1 | Con prueba concreta |
| 2. EFI y EFE | 2.1, 2.2, 2.3, 2.4 | Con prueba concreta |
| 3. MPC | 3.1, 3.2 | Con prueba concreta |
| 4. PEYEA | 4.1, 4.2, 4.3, 4.4 | Con prueba concreta |
| 5. MIE | 5.1, 5.2, 5.3, 5.4 | Con prueba concreta |
| 6. GE | 6.1 (solo tabla de búsqueda) | Sin prueba de ubicación: faltan definiciones de entrada |
| 7. Uso sin instalación | 7.1 | Con prueba concreta, ejecutable solo en equipo real |

Total: 17 pruebas Given-When-Then.

Lista consolidada de puntos [VERIFICAR] para confirmar con el profesor:

1. BCG: umbrales de participación relativa y crecimiento, escala del eje X, si la herramienta calcula la participación relativa, y tratamiento de utilidad negativa o cero.
2. EFI y EFE: diagnóstico en total exactamente 2.5, y restricción de clasificación por tipo de factor.
3. MPC: cantidad de competidores admitidos, validación de clasificación 1 a 4 y manejo de empates.
4. PEYEA: definición del vector (segmento o magnitud y ángulo), cuadrante en el eje cero.
5. MIE: numeración de celdas, asignación de zonas, y huecos entre rangos.
6. GE: origen, escala y punto de corte de los dos ejes, y listas de estrategias por cuadrante.
7. Historia 7: navegadores soportados y uso de CDN.

## Elaboration II

Objetivo de la fase: definir la arquitectura del framework (componentes, colaboraciones y diagramas) sin escribir todavía código de la aplicación. Las convenciones de cálculo, clasificación y validación ya fijadas en Elaboration I no se reinterpretan aquí; los componentes solo las aplican. Las decisiones que Inception y Elaboration I no resuelven se marcan [VERIFICAR] y se consolidan al final de esta sección.

Decisión de esta fase que afecta a Elaboration I: el Exportador usa una librería de generación de .xlsx embebida en el propio archivo, sin CDN externa, para funcionar sin conexión a internet. Esto responde al punto [VERIFICAR] de la Historia 7 sobre el uso de CDN. La lista de Elaboration I no se modificó en esta fase.

Principios de la arquitectura:

- La Vista es el único orquestador. Recibe los eventos del usuario y llama a los otros cinco componentes en el orden que corresponda. Ningún otro componente llama a otro, lo que evita dependencias cruzadas.
- El Validador y el Motor de Cálculo no dependen del navegador ni de la interfaz (entran datos, salen resultados). Por eso las pruebas de Elaboration I se pueden ejecutar contra ellos sin abrir una página.
- La fuente de verdad es lo que el usuario ingresó (factores, divisiones, calificaciones). Los totales y clasificaciones se recalculan cuando hacen falta, no se guardan como datos independientes, así que nunca quedan desactualizados respecto de los datos.
- Los datos ingresados se guardan en cada cambio, sean válidos o no, para no perder progreso (requisito no funcional de persistencia). Los resultados solo se calculan y muestran cuando la validación pasa.

### 1. Componentes del framework

| Componente | Responsabilidad |
|---|---|
| Vista | Renderiza los formularios de entrada y los resultados de cada matriz, y despacha los eventos del usuario hacia los demás componentes. |
| Validador | Verifica que los datos ingresados cumplan las reglas de cada matriz antes de calcular (pesos que suman 1, rangos numéricos válidos según la convención de cada matriz, campos no vacíos). |
| Motor de Cálculo | Contiene la lógica de las siete matrices (BCG, EFI, EFE, MPC, PEYEA, MIE, GE), aplicando exactamente las fórmulas y clasificaciones ya definidas y probadas en Elaboration I. |
| Motor de Gráficos | Dibuja las representaciones visuales de cada matriz (cuadrantes del BCG, vector del PEYEA, cuadrícula de la MIE, y las demás según corresponda). |
| Persistencia | Guarda y recupera el estado de la sesión en el almacenamiento local del navegador, para que los datos no se pierdan al recargar. |
| Exportador | Genera un archivo Excel (.xlsx) descargable con los datos ingresados y los resultados calculados de la matriz activa, usando una librería de generación de Excel embebida en el propio archivo (sin CDN externa), para mantener el funcionamiento sin conexión a internet. |

Relación con las pruebas de Elaboration I:

| Componente | Pruebas de Elaboration I que lo ejercitan |
|---|---|
| Validador | 2.3 (pesos suman 0.95), 2.4 (tolerancia de punto flotante) |
| Motor de Cálculo | 1.1, 2.1, 2.2, 3.1, 3.2, 4.1 a 4.4, 5.1 a 5.4, y 6.1 en lo que se pueda ubicar |
| Motor de Gráficos | 1.1 (cuadrantes y burbujas), 4.1 a 4.4 (cuadrante del vector), 5.1 a 5.4 (celda de la MIE) |
| Persistencia | 7.1 (los datos siguen tras recargar) |
| Vista y Exportador | 7.1 (las siete matrices accesibles); el Exportador no tiene prueba propia en Elaboration I [VERIFICAR] |

### 2. Escenarios (plays)

**Escenario 1: cálculo simple de una matriz nueva (BCG)**

El usuario abre el módulo BCG e ingresa las cuatro divisiones de la prueba 1.1. Con cada cambio, la Vista pide a la Persistencia que guarde el estado. Luego la Vista pasa los datos al Validador, que confirma que cada división tenga ingresos, utilidades, participación relativa y crecimiento numéricos. La Vista pasa los datos válidos al Motor de Cálculo, que devuelve el cuadrante de cada división (A Estrella, B Interrogante, C Vaca lechera, D Perro) y los porcentajes de ingresos y utilidades. La Vista entrega ese resultado al Motor de Gráficos, que dibuja las burbujas en el plano de cuatro cuadrantes, y muestra el resultado al usuario.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.

**Escenario 2: una matriz reutiliza resultados de otras (MIE con EFI y EFE)**

El usuario abre el módulo MIE. La Vista pide a la Persistencia los factores guardados de EFI y EFE. El Validador comprueba ambos conjuntos de factores. Si son válidos, el Motor de Cálculo recalcula los totales (2.45 y 2.90 con los datos de las pruebas 2.1 y 2.2) y con ellos ubica la empresa en la celda V, zona "retener y mantener", como en la prueba 5.1. La Vista entrega el resultado al Motor de Gráficos, que dibuja la cuadrícula de nueve celdas con la posición marcada. Si falta EFI o EFE, o alguno no pasa la validación, la Vista indica cuál matriz debe completarse y no calcula la MIE.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
Decisión tomada: la MIE reutiliza los datos de EFI y EFE y recalcula los totales, no lee totales guardados. [VERIFICAR] si el curso permite ingresar los totales a mano cuando el estudiante no ha llenado EFI y EFE (la Historia 5 dice "a partir de mis totales de EFI y EFE", sin aclarar el origen).

**Escenario 3: rechazo por datos inválidos**

El usuario, en el EFI de la prueba 2.1, cambia el peso de D3 de 0.10 a 0.05. La Vista guarda el estado en la Persistencia y pasa los datos al Validador, que detecta que los pesos suman 0.95 en lugar de 1.00 y devuelve el error. La Vista muestra el error y no muestra total ni diagnóstico. El Motor de Cálculo no se invoca, y el Motor de Gráficos tampoco.
Componentes: Vista, Persistencia, Validador.
Este escenario cubre la prueba 2.3. Muestra un patrón distinto al del escenario 1: el flujo se corta en la validación.

**Escenario 4: exportación a Excel de la matriz activa**

Con el EFI válido de la prueba 2.1 abierto, el usuario pulsa "Exportar". La Vista pasa los datos al Validador, que confirma que sean válidos. La Vista pide el resultado al Motor de Cálculo (total 2.45 y diagnóstico "posición interna débil") y entrega al Exportador la matriz, los datos ingresados y el resultado. El Exportador genera el archivo .xlsx con la librería embebida y la Vista dispara la descarga en el navegador. Todo ocurre sin conexión a internet. Si los datos no son válidos, la Vista muestra los errores y no llama al Exportador.
Componentes: Vista, Validador, Motor de Cálculo, Exportador.
[VERIFICAR] si se permite exportar datos incompletos o inválidos (solo con los datos ingresados y sin resultados), en lugar de bloquear la exportación.

**Escenario 5: recuperación de una sesión previa al recargar**

El usuario recarga la página o vuelve a abrir el archivo. La Vista pide a la Persistencia el estado guardado. Si existe, la Vista vuelve a llenar los formularios con esos datos y el Validador los revisa. Para cada matriz que pase la validación, el Motor de Cálculo recalcula el resultado y el Motor de Gráficos lo dibuja de nuevo. Las matrices incompletas se muestran con sus datos pero sin resultado. Si el almacenamiento está vacío, corrupto o inaccesible, la Persistencia devuelve un estado vacío y la Vista muestra los formularios en blanco.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
[VERIFICAR]: si una matriz incompleta tras la recarga muestra mensajes de error de inmediato o solo después de que el usuario la edite; y si el usuario recibe un aviso cuando se descarta un estado corrupto.

Cobertura de componentes por escenario:

| Componente | Esc. 1 | Esc. 2 | Esc. 3 | Esc. 4 | Esc. 5 |
|---|---|---|---|---|---|
| Vista | sí | sí | sí | sí | sí |
| Validador | sí | sí | sí | sí | sí |
| Motor de Cálculo | sí | sí | no | sí | sí |
| Motor de Gráficos | sí | sí | no | no | sí |
| Persistencia | sí | sí | sí | no | sí |
| Exportador | no | no | no | sí | no |

### 3. Diagramas

**Diagrama de clases**

Los nombres de las clases van sin acentos ni espacios por compatibilidad con Mermaid: `MotorCalculo` es el Motor de Cálculo y `MotorGraficos` es el Motor de Gráficos. La Vista es la única clase que depende de las demás, coherente con el principio de orquestación único.

```mermaid
classDiagram
    class Vista {
        +renderFormulario(matriz)
        +renderResultados(matriz, resultado)
        +mostrarErrores(errores)
        +despacharEvento(evento)
    }
    class Validador {
        +validar(matriz, datos) ResultadoValidacion
        +validarPesos(factores) bool
        +validarRango(valor, min, max) bool
        +validarCamposVacios(datos) bool
    }
    class MotorCalculo {
        +calcularBCG(divisiones) ResultadoBCG
        +calcularEFI(factores) ResultadoPonderado
        +calcularEFE(factores) ResultadoPonderado
        +calcularMPC(factores, empresas) ResultadoMPC
        +calcularPEYEA(ejes) ResultadoPEYEA
        +ubicarMIE(totalEFI, totalEFE) ResultadoMIE
        +ubicarGE(ejes) ResultadoGE
        +calcularAE(variables, matriz) ResultadoAE
        +calcularRadar(calificaciones, errores) ResultadoRadar
    }
    class MotorGraficos {
        +dibujarBCG(resultado)
        +dibujarPEYEA(resultado)
        +dibujarMIE(resultado)
        +dibujarGE(resultado)
        +dibujarAE(resultado)
        +dibujarRadar(resultado)
    }
    class Persistencia {
        +guardar(estado)
        +cargar() Estado
        +limpiar()
    }
    class Exportador {
        +exportarXLSX(matriz, datos, resultado) ArchivoXLSX
    }
    Vista ..> Validador : usa
    Vista ..> MotorCalculo : usa
    Vista ..> MotorGraficos : usa
    Vista ..> Persistencia : usa
    Vista ..> Exportador : usa
```

Notas del diagrama de clases:

- `calcularAE` y `dibujarAE` se agregaron en la Elaboration II del Módulo 2 (Análisis Estructural). `calcularRadar` y `dibujarRadar` se agregaron en la Elaboration II del Módulo 3 (Radar Estratégico). Las otras firmas no cambian.
- Las firmas de `ubicarGE` y `dibujarGE` son provisionales: dependen de las decisiones pendientes de la Historia 6 (origen y escala de los ejes) [VERIFICAR].
- El Motor de Gráficos solo lista los métodos de las matrices con representación visual definida en Elaboration I (BCG, PEYEA, MIE, GE). [VERIFICAR] si EFI, EFE y MPC llevan también una representación gráfica (por ejemplo barras en el MPC) o solo tabla.

**Diagrama de secuencia del escenario 1: cálculo simple (BCG)**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: ingresa las divisiones del BCG
    V->>P: guardar(estado)
    V->>VA: validar(BCG, divisiones)
    VA-->>V: datos válidos
    V->>MC: calcularBCG(divisiones)
    MC-->>V: cuadrantes y porcentajes de burbuja
    V->>MG: dibujarBCG(resultado)
    MG-->>V: gráfico de cuatro cuadrantes
    V-->>U: muestra resultados y gráfico
```

**Diagrama de secuencia del escenario 2: MIE reutiliza EFI y EFE**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: abre el módulo MIE
    V->>P: cargar() factores de EFI y EFE
    P-->>V: factores guardados
    V->>VA: validar(EFI, factores)
    VA-->>V: resultado de validación EFI
    V->>VA: validar(EFE, factores)
    VA-->>V: resultado de validación EFE
    alt EFI y EFE válidos
        V->>MC: calcularEFI(factores)
        MC-->>V: total EFI
        V->>MC: calcularEFE(factores)
        MC-->>V: total EFE
        V->>MC: ubicarMIE(totalEFI, totalEFE)
        MC-->>V: celda y zona
        V->>MG: dibujarMIE(resultado)
        MG-->>V: cuadrícula con la posición marcada
        V-->>U: muestra celda y zona
    else falta EFI o EFE, o alguno es inválido
        V-->>U: indica cuál matriz debe completar
    end
```

**Diagrama de secuencia del escenario 3: rechazo por datos inválidos**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    U->>V: cambia el peso de D3 de 0.10 a 0.05
    V->>P: guardar(estado)
    V->>VA: validar(EFI, datos)
    VA-->>V: error, los pesos suman 0.95 y deben sumar 1.00
    V-->>U: muestra el error sin total ni diagnóstico
```

**Diagrama de secuencia del escenario 4: exportación a Excel**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant E as Exportador
    U->>V: pulsa Exportar en la matriz activa
    V->>VA: validar(EFI, datos)
    alt datos válidos
        VA-->>V: datos válidos
        V->>MC: calcularEFI(datos)
        MC-->>V: total y diagnóstico
        V->>E: exportarXLSX(EFI, datos, resultado)
        E-->>V: archivo .xlsx generado con la librería embebida
        V-->>U: descarga el archivo
    else datos inválidos
        VA-->>V: errores
        V-->>U: muestra los errores, no exporta
    end
```

**Diagrama de secuencia del escenario 5: recuperación al recargar**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: recarga la página
    V->>P: cargar()
    alt hay estado guardado
        P-->>V: estado de la sesión previa
        V->>V: vuelve a llenar los formularios
        V->>VA: validar(matriz, datos recuperados)
        VA-->>V: resultado de validación
        opt la matriz es válida
            V->>MC: calcular la matriz
            MC-->>V: resultado
            V->>MG: dibujar(resultado)
            MG-->>V: gráfico
        end
        V-->>U: formularios con sus datos y resultados de las matrices completas
    else vacío, corrupto o inaccesible
        P-->>V: estado vacío
        V-->>U: formularios en blanco
    end
```

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Estructura del Excel exportado: hojas, nombre del archivo, valores fijos o fórmulas vivas, y si incluye una imagen del gráfico. Solo está decidido que se exporta la matriz activa, con datos ingresados y resultados.
2. Librería concreta de generación de .xlsx que se embebe: peso, licencia y forma de incluirla en el archivo. Se decide al comienzo de Construction.
3. Representación gráfica de EFI, EFE y MPC (gráfico o solo tabla), y firmas definitivas de `ubicarGE` y `dibujarGE` (dependen de las decisiones pendientes de la Historia 6).
4. Origen de los totales en la MIE cuando el estudiante no ha llenado EFI y EFE: si se permite ingresarlos a mano.
5. Exportación con datos incompletos o inválidos: bloquear (como está descrito) o permitir solo con los datos ingresados.
6. Recuperación de sesión: mensajes de error para matrices incompletas tras la recarga, aviso al descartar un estado corrupto, y detalles de almacenamiento (clave, estructura, versión del formato guardado, si el estado es por matriz o global).
7. El Exportador no tiene prueba de aceptación en Elaboration I. Decisión del juez: no se agrega ahora. Se define un caso de prueba concreto en el plan de pruebas de Construction III, cuando ya exista una implementación real que genere el archivo .xlsx.

## Construction I

Objetivo de la fase: dejar decidido el stack y documentado el esqueleto de la estructura del proyecto, sin comportamiento real. Los cuerpos de los métodos están vacíos o llevan un comentario de marcador de posición. Ninguna lógica de cálculo, validación, graficado, persistencia ni exportación se implementa en esta fase. Esta fase no crea ningún archivo nuevo del proyecto: el esqueleto vive dentro de este documento y se copiará a un archivo HTML al empezar Construction II.

### 1. Stack tecnológico y decisiones de arquitectura

| Decisión | Justificación (requisito no funcional de Inception) |
|---|---|
| Stack: HTML, CSS y JavaScript sin framework ni backend. | "Todos los cálculos se ejecutan en el navegador, sin backend ni servidor propio" y "el archivo abre directamente haciendo doble clic, sin instalación". Un framework agregaría un paso de compilación o una dependencia externa. |
| Almacenamiento persistente: en el navegador (localStorage), sin base de datos ni servidor. | "El progreso del usuario se guarda automáticamente en el navegador (localStorage) para no perderse si cierra la pestaña". Además evita la fragilidad de Access que motivó el proyecto. |
| Alcance de autenticación: ninguno. | Sin backend no hay dónde validar identidades, y la persistencia compartida entre integrantes está fuera del alcance de v1. Los datos quedan en el navegador de cada usuario. |
| Alcance de interfaz: interfaz web mínima integrada en el mismo archivo, sin páginas separadas. | El archivo debe abrir con doble clic y funcionar sin conexión. Como se decidió en Elaboration II, no hay CDN: HTML, estilos, script y librerías van en un solo archivo. |

Consecuencias de estas decisiones que conviene tener presentes:

- Los datos viven en un solo navegador y equipo. Otro navegador o dispositivo no los ve.
- Borrar los datos del sitio en el navegador borra el progreso. Este comportamiento coincide con lo ya declarado como fuera de alcance en Inception.

### 2. Esqueleto del proyecto

Estructura de archivos prevista. El archivo HTML no existe todavía [VERIFICAR: nombre del archivo, se propone `index.html`].

```text
Mtx/
├── VUP.md            documento del proceso (este archivo)
└── index.html        (por crear en Construction II) todo el producto en un solo archivo
    ├── <style>       estilos
    ├── <body>        barra de navegación, un contenedor por matriz y botón de exportar
    └── <script>      seis objetos: Validador, MotorCalculo, MotorGraficos,
                      Persistencia, Exportador y Vista
```

Esqueleto del archivo. Las firmas de los métodos son las del diagrama de clases de Elaboration II. Cada cuerpo queda vacío con un marcador de posición.

```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Mtx</title>
  <style>
    /* Construction II: estilos base, sin diseño visual definido todavía */
    .matriz[hidden] { display: none; }
  </style>
</head>
<body>
  <header>
    <h1>Mtx</h1>
    <!-- Construction II: navegación entre las nueve matrices [VERIFICAR] -->
    <nav id="navegacion"></nav>
    <button id="btn-exportar" type="button">Exportar a Excel</button>
  </header>

  <main>
    <section id="matriz-bcg" class="matriz" data-matriz="BCG">
      <h2>BCG</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-efi" class="matriz" data-matriz="EFI" hidden>
      <h2>EFI</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-efe" class="matriz" data-matriz="EFE" hidden>
      <h2>EFE</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-mpc" class="matriz" data-matriz="MPC" hidden>
      <h2>MPC</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-peyea" class="matriz" data-matriz="PEYEA" hidden>
      <h2>PEYEA</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-mie" class="matriz" data-matriz="MIE" hidden>
      <h2>MIE</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-ge" class="matriz" data-matriz="GE" hidden>
      <h2>GE</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-ae" class="matriz" data-matriz="AE" hidden>
      <h2>AE</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>

    <section id="matriz-radar" class="matriz" data-matriz="RADAR" hidden>
      <h2>RADAR</h2>
      <div class="formulario"></div>
      <div class="errores"></div>
      <div class="resultados"></div>
      <div class="grafico"></div>
    </section>
  </main>

  <!-- Construction III: librería de generación de .xlsx embebida aquí, sin CDN [VERIFICAR] -->

  <script>
    'use strict';

    // Auxiliar privada del Validador para AE (Construction II/III): aplica el mínimo de dos variables y excluye la diagonal del conteo de campos vacíos
    function erroresAE(datos, errores) {
      // Construction II/III
    }

    // Auxiliar privada del Validador para RADAR (Construction II/III): revisa solo el rango (0 a 5, con validarRango) de las calificaciones que tengan valor
    // y devuelve los errores por característica (índice plano 0 a 55). Una calificación sin valor (null) nunca es un error
    function erroresRadar(datos, errores) {
      // Construction II/III
    }

    const Validador = {
      validar(matriz, datos) {
        // Construction II/III
      },
      validarPesos(factores) {
        // Construction II/III
      },
      validarRango(valor, min, max) {
        // Construction II/III
      },
      validarCamposVacios(datos) {
        // Construction II/III
      }
    };

    const MotorCalculo = {
      calcularBCG(divisiones) {
        // Construction II/III
      },
      calcularEFI(factores) {
        // Construction II/III
      },
      calcularEFE(factores) {
        // Construction II/III
      },
      calcularMPC(factores, empresas) {
        // Construction II/III
      },
      calcularPEYEA(ejes) {
        // Construction II/III
      },
      ubicarMIE(totalEFI, totalEFE) {
        // Construction II/III
      },
      ubicarGE(ejes) {
        // Construction II/III (firma provisional, depende de la Historia 6)
      },
      calcularAE(variables, matriz) {
        // Construction II/III
      },
      calcularRadar(calificaciones, errores) {
        // Construction II/III (devuelve un arreglo de 14 posiciones { estado, puntaje }; ver "Forma de ResultadoRadar" en Construction I del Módulo 3)
      }
    };

    const MotorGraficos = {
      dibujarBCG(resultado) {
        // Construction II/III
      },
      dibujarPEYEA(resultado) {
        // Construction II/III
      },
      dibujarMIE(resultado) {
        // Construction II/III
      },
      dibujarGE(resultado) {
        // Construction II/III (firma provisional, depende de la Historia 6)
      },
      dibujarAE(resultado) {
        // Construction II/III
      },
      dibujarRadar(resultado) {
        // Construction II/III (solo recorre el arreglo de ResultadoRadar y dibuja un punto donde estado sea "completo")
      }
    };

    const Persistencia = {
      guardar(estado) {
        // Construction II/III
      },
      cargar() {
        // Construction II/III (si el estado guardado no trae datos.AE, devolver AE vacío; sin cambio de versión del formato)
        // Construction II/III (igual para RADAR: si el estado guardado no trae datos.RADAR, devolver RADAR con 56 null)
      },
      limpiar() {
        // Construction II/III
      }
    };

    const Exportador = {
      exportarXLSX(matriz, datos, resultado) {
        // Construction III
      }
    };

    // Auxiliares privadas de la Vista para AE (Construction II/III). formularioAE dibuja la matriz NxN (celda con data-campo "matriz.i.j",
    // diagonal bloqueada); resultadosAE construye la tabla de ranking y la lista de solo lectura de la hoja Validadas
    function formularioAE(datos) {
      // Construction II/III
    }

    function resultadosAE(resultado) {
      // Construction II/III
    }

    // Auxiliares privadas de la Vista para RADAR (Construction II/III). formularioRadar dibuja las 56 calificaciones agrupadas en cinco etapas
    // (celda con data-campo "calificaciones.i", índice plano 0 a 55, sin controles para agregar ni quitar); resultadosRadar construye los puntajes por componente
    function formularioRadar(datos) {
      // Construction II/III
    }

    function resultadosRadar(resultado) {
      // Construction II/III
    }

    // La Vista es el único orquestador: llama a los otros cinco objetos.
    const Vista = {
      renderFormulario(matriz) {
        // Construction II/III
      },
      renderResultados(matriz, resultado) {
        // Construction II/III
      },
      mostrarErrores(errores) {
        // Construction II/III
      },
      despacharEvento(evento) {
        // Construction II/III
      }
    };

    // Construction II: rutina de arranque (ver decisión del juez en el punto 4 de "Puntos nuevos marcados [VERIFICAR] en esta fase"), no un método de la Vista
  </script>
</body>
</html>
```

Correspondencia con Elaboration II:

- Los seis objetos del script son los seis componentes del diagrama de clases, con los mismos nombres y las mismas firmas. `MotorCalculo` y `MotorGraficos` van sin acentos por compatibilidad, igual que en los diagramas.
- El script no agrega métodos ni componentes nuevos. No hay métodos auxiliares, constantes de configuración ni estado global, en el esqueleto de este primer módulo (la Construction I del Módulo 2, Análisis Estructural, sí agrega auxiliares privadas: ver esa sección más abajo).
- Cada `<section class="matriz">` es el contenedor de una de las matrices (siete del Módulo 1 y, desde la Construction I del Módulo 2, una octava para Análisis Estructural y, desde la Construction I del Módulo 3, una novena para Radar Estratégico). Sus cuatro `<div>` son las zonas que la Vista llena: `formulario` (`renderFormulario`), `errores` (`mostrarErrores`), `resultados` (`renderResultados`) y `grafico` (donde dibuja el Motor de Gráficos).
- Solo la sección BCG arranca visible. Es un valor inicial provisional del esqueleto, no una decisión de qué matriz se muestra primero [VERIFICAR].

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Resuelto: el archivo se llama `index.html`.
2. Resuelto: se implementó en Construction II (un botón por matriz dentro del `<nav id="navegacion">`; la Vista muestra la sección elegida y oculta las demás con `hidden`; BCG se muestra al abrir).
3. Resuelto: el Motor de Gráficos usa SVG, no canvas. Los gráficos de Mtx son formas simples en dos dimensiones (cuadrantes, burbujas, un vector, una cuadrícula de nueve celdas), no requieren dibujar grandes volúmenes de píxeles, y SVG permite inspeccionar, probar y dar estilo con CSS a cada elemento como un nodo del DOM, sin una librería adicional. Los contenedores `grafico` siguen siendo `<div>` en el esqueleto; el `<svg>` se agrega dentro de cada uno en Construction II.
4. Resuelto: se agrega una rutina de arranque fuera de los seis componentes (no un método nuevo en la Vista, para no modificar las firmas ya aprobadas en Elaboration II), que al cargar la página llama a `Persistencia.cargar()` y, con el resultado, a los métodos ya existentes de la Vista para repoblar cada matriz. El código de esta rutina se agrega en Construction II, junto con el resto del comportamiento.
5. Resuelto: se implementó en Construction II (la librería va embebida en un `<script id="sheetjs">` del propio `index.html`, antes del script de la aplicación, donde estaba el comentario marcador).
6. Resuelto: `matriz` es siempre un string con una de las nueve siglas ya usadas en Elaboration I y en el atributo `data-matriz` del esqueleto (`"BCG"`, `"EFI"`, `"EFE"`, `"MPC"`, `"PEYEA"`, `"MIE"`, `"GE"`, desde Elaboration II del Módulo 2, `"AE"`, y desde Elaboration II del Módulo 3, `"RADAR"`), no un objeto ni un código numérico.

## Construction II

Objetivo de la fase: implementar la lógica real de las seis matrices con fórmulas ya probadas en Elaboration I, sobre el esqueleto de Construction I, más la navegación entre matrices y la exportación a Excel. Es la primera fase que escribe código de la aplicación.

### 1. Resumen de lo implementado y lo pendiente

Archivos creados o modificados en esta fase:

| Archivo | Cambio |
|---|---|
| `index.html` | Nuevo. Un solo archivo con estilos, HTML, la librería SheetJS embebida y el script de la aplicación (unas 1 190 líneas de script propio). |
| `tests/elaboration1.test.js` | Nuevo. Script de Node sin dependencias que verifica `Validador`, `MotorCalculo`, `Exportador` y `Persistencia` contra las pruebas de Elaboration I. No estaba pedido explícitamente; se agregó para que el juez pueda repetir la verificación con `node tests/elaboration1.test.js` [VERIFICAR: si se conserva en el repositorio]. |
| `VUP.md` | Esta sección, y los puntos 2 y 5 de Construction I marcados como resueltos. |

| Componente o matriz | Estado |
|---|---|
| BCG, EFI, EFE, MPC, PEYEA, MIE | Implementados: formulario, validación, cálculo y resultados en pantalla. BCG, PEYEA y MIE además dibujan su gráfico en SVG. |
| GE | No implementada, como decidió el juez. `ubicarGE` y `dibujarGE` lanzan un error explícito de módulo pendiente y la interfaz nunca los llama. La sección GE muestra un mensaje claro de que el módulo está pendiente de definir con el curso (Historia 6), y su botón de exportar queda deshabilitado. |
| Navegación | Un botón por matriz dentro de `<nav id="navegacion">`. La Vista muestra la sección elegida y oculta las demás con `hidden`. BCG se muestra al abrir. |
| Persistencia | Guarda en localStorage en cada cambio, sea válido o no. Al abrir la página recupera la sesión previa (matriz activa y datos) y recalcula. |
| Exportador | Genera un `.xlsx` de la matriz activa con SheetJS embebido. |
| Gráficos de EFI, EFE y MPC | Sin gráfico, solo tabla de resultados (queda abierto el punto 3 de Elaboration II). |

Decisiones de implementación que no cambian ninguna firma aprobada:

- Los seis objetos (`Vista`, `Validador`, `MotorCalculo`, `MotorGraficos`, `Persistencia`, `Exportador`) tienen exactamente los métodos y el número de parámetros del diagrama de clases de Elaboration II. La prueba X.1 lo comprueba leyendo el diagrama de este documento.
- Fuera de esos objetos hay constantes y funciones auxiliares privadas (por ejemplo `aNumero`, `redondear`, `evaluarMatriz`, `actualizarMatriz`) y dos variables de módulo (`estado` y `contextoMatriz`). La rutina `arrancar()` es la rutina de arranque decidida por el juez en Construction I: no es un método de la Vista.
- La Vista sigue siendo el único orquestador. Ningún componente llama a otro.
- Los números se escriben como texto y se aceptan con coma o punto decimal (`1,80` o `1.80`). Los resultados se redondean a 6 decimales antes de clasificar, para que el ruido de punto flotante no cambie un diagnóstico.
- El script de la aplicación no toca el DOM al cargarse fuera del navegador, por eso `Validador` y `MotorCalculo` se prueban en Node.

### 2. Verificación prueba por prueba contra Elaboration I

Comando para repetirla:

```bash
node tests/elaboration1.test.js
```

Resultado de la última ejecución: 21 de 21 comprobaciones correctas (16 pruebas de Elaboration I y 5 comprobaciones adicionales). Además se comprobó que las pruebas detectan errores: con cuatro fallos introducidos a propósito en una copia fuera del repositorio, 6 comprobaciones fallaron.

| Prueba | Qué verifica | Resultado |
|---|---|---|
| 1.1 | Cuadrantes A Estrella, B Interrogante, C Vaca lechera, D Perro; totales 1 000 y 200; % de ingresos 50/30/15/5 y de utilidades 50/15/30/5 | Coincide |
| 2.1 | EFI: suma de pesos 1.00, ponderados 0.80, 0.60, 0.30, 0.25, 0.40, 0.10, total 2.45, "posición interna débil" | Coincide |
| 2.2 | EFE: suma de pesos 1.00, ponderados 1.00, 0.60, 0.20, 0.75, 0.30, 0.05, total 2.90, diagnóstico de aprovechamiento por encima del promedio | Coincide |
| 2.3 | Pesos que suman 0.95: se rechazan y el mensaje muestra 0.95 y 1.00; el total aritmético de referencia sería 2.40 | Coincide |
| 2.4 | Pesos 0.6, 0.3 y 0.1 (suma 0.9999999999999999 en JavaScript) se aceptan con tolerancia de 0.001 | Coincide |
| 3.1 | MPC: totales 3.00, 3.10, 2.90; orden A, mi empresa, B; mi empresa en 2.º lugar de 3 | Coincide |
| 3.2 | Nuevos pesos 0.10, 0.50, 0.20, 0.20: totales 2.70, 2.90, 3.40; orden B, A, mi empresa; mi empresa en 3.er lugar | Coincide |
| 4.1 | PEYEA agresivo: promedios 4.00, 4.00, −2.00, −3.00; X = 2.00, Y = 1.00 | Coincide |
| 4.2 | PEYEA conservador: X = −2.50, Y = 3.00 | Coincide |
| 4.3 | PEYEA defensivo: X = −3.00, Y = −3.50 | Coincide |
| 4.4 | PEYEA competitivo (versión corregida): FF = 2.00, X = 3.50, Y = −3.00 | Coincide |
| 5.1 | MIE con los totales 2.45 y 2.90 derivados de las pruebas 2.1 y 2.2: promedio, promedio, celda V, retener y mantener | Coincide |
| 5.2 | MIE con 3.20 y 3.50: celda I, crecer y construir | Coincide |
| 5.3 | MIE con 1.80 y 1.50: celda IX, cosechar o desinvertir | Coincide |
| 5.4 | MIE con EFE 3.50 y EFI 1.99, 2.00, 2.99, 3.00: celdas III, II, II, I | Coincide |
| 6.1 | GE: no hay lógica de ubicación. `ubicarGE`, `dibujarGE` y `validar("GE")` indican que el módulo está pendiente y no devuelven ningún resultado inventado | No aplica (decisión del juez); se comprueba que falla de forma explícita |
| 7.1 | Uso sin instalación | Verificada en parte, ver más abajo |

Comprobaciones adicionales del script (no son pruebas de Elaboration I):

| Comprobación | Qué verifica |
|---|---|
| X.1 | Los seis objetos tienen exactamente los métodos y parámetros del diagrama de clases |
| X.2 | Validación de rangos, vacíos y formatos: PEYEA rechaza el 0 en las cuatro variables y valores fuera de rango, EFI rechaza clasificación 0, 5, 2.5 o vacía, MPC exige mi empresa más un competidor, la coma decimal se acepta |
| X.3 | Los casos límite marcados [VERIFICAR] se detectan (total 2.5, X = 0, umbrales del BCG, 1.995 en la MIE) |
| X.4 | El Exportador, con la copia de SheetJS que está dentro de `index.html`, genera las hojas "Datos" y "Resultados" con los valores esperados |
| X.5 | Persistencia: guardar, cargar, JSON corrupto, estructura inválida y almacenamiento inaccesible devuelven un estado vacío sin lanzar errores |

Verificación de la interfaz (Chrome y el navegador integrado de la aplicación; no está automatizada en el repositorio):

- Con el archivo abierto desde el disco (`file://`) en Chrome sin interfaz gráfica, los scripts corren, BCG queda visible, el formulario se dibuja y no aparecen errores de consola de la página.
- En el navegador integrado, con datos escritos con el teclado (incluido `1,80`) y con eventos reales, se reprodujeron en pantalla los resultados de las pruebas 1.1, 2.1, 2.2, 3.1, 4.1, 5.1 y 2.3, con los mismos números de la tabla.
- Tras recargar la página, la sesión se recuperó completa: volvió a la matriz activa, conservó lo escrito y recalculó todas las matrices (escenario 5 de Elaboration II).
- El botón de exportar disparó la descarga de `Mtx-EFI.xlsx` y el contenido de `Mtx-MPC.xlsx`, leído de vuelta con SheetJS, coincide con los datos y resultados de la prueba 3.1.
- La sección GE muestra el mensaje de módulo pendiente y el botón de exportar queda deshabilitado en ella.
- El script de la aplicación no contiene solicitudes de red (`fetch`, `XMLHttpRequest`, `WebSocket`); no hay CDN ni recursos externos.
- No se verificó en Edge, Firefox ni Safari, ni con un archivo descargado de internet (marca de la web). La prueba 7.1 queda por completar en un equipo real, como ya decía Elaboration I.

### 3. Librería de exportación: SheetJS Community Edition

| Dato | Valor |
|---|---|
| Versión embebida | 0.20.3. Es la versión que la página oficial de instalación standalone (docs.sheetjs.com) declaraba como vigente el 2026-09-29. |
| Archivo | `xlsx.full.min.js`, build completa (951 904 bytes). |
| Fuente | `https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js` |
| SHA-256 | `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41` |
| Licencia | Apache-2.0, según el `package.json` y el `LICENSE` del paquete de esa versión. |
| Cómo se incluye | Texto completo, sin modificar, dentro de `<script id="sheetjs">` de `index.html`, precedido por un comentario con versión, fuente y licencia. No se carga desde un archivo aparte ni desde un CDN. |
| API usada | `XLSX.utils.book_new()`, `XLSX.utils.aoa_to_sheet()`, `XLSX.utils.book_append_sheet()` y `XLSX.writeFile()`. |

Nota técnica: el texto de la librería contiene la secuencia `<!--` cinco veces, todas dentro de cadenas de JavaScript, sin ninguna etiqueta `<script` ni `</script`, así que no interfiere con el análisis del HTML. Se comprobó que carga sin errores en el navegador. `index.html` pesa ahora un poco más de 1 MB por esta librería.

### 4. Criterios provisionales aplicados donde Elaboration I dejó un [VERIFICAR]

Cada uno tiene un comentario en el código que cita el [VERIFICAR] correspondiente, y la Vista avisa en pantalla de los casos límite con un texto que dice que la regla está pendiente de confirmar.

| Punto abierto de Elaboration I | Criterio implementado |
|---|---|
| EFI y EFE con total exactamente 2.5 | Menor que 2.5 es débil y 2.5 o más es fuerte; se muestra el aviso. |
| BCG con valor exactamente en el umbral | Participación relativa 1.0 o más es alta y crecimiento 10 % o más es alto; se muestra el aviso. |
| BCG con umbrales y escala del eje X | Umbrales 1.0 y 10 %. El eje X es logarítmico, con alta participación a la izquierda y el umbral en el centro; el eje Y es lineal con el umbral en el centro. |
| BCG con utilidad negativa o cero | La burbuja se dibuja con el tamaño mínimo. Se exige que la suma de utilidades sea mayor que 0. |
| PEYEA con X = 0 o Y = 0 | El 0 se toma como positivo; se muestra el aviso. |
| MIE con un valor entre 1.99 y 2.0 | Cortes en menos de 2.0 y menos de 3.0. |
| MPC con empates | Los empates comparten posición. |
| MPC con cantidad de competidores | Mi empresa más al menos un competidor, sin máximo. |
| EFI con clasificación por tipo de factor | Se acepta de 1 a 4 para todos los factores. |
| PEYEA con vector | Se muestra el vector (X, Y); no se calculan magnitud ni ángulo. |

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Nombres opcionales: el nombre de una división, factor o empresa no es obligatorio; se rellenan con "División 1", "Factor 2", etc. Falta confirmar si el curso exige nombres.
2. Utilidades del BCG: se rechaza un conjunto de divisiones cuya suma de utilidades sea cero o negativa, porque no se puede sacar el porcentaje de utilidades. Puede bloquear el caso legítimo de una empresa con pérdidas totales. Falta decidir si se debe permitir y qué mostrar entonces.
3. Rango del crecimiento del mercado en el BCG: se acepta cualquier número mayor o igual a −100 (%). Es un valor razonable, no fijado por el curso.
4. PEYEA con decimales: los factores de PEYEA se aceptan con decimales dentro del rango (por ejemplo 3.5). D'Alessio califica con enteros. Las clasificaciones de EFI, EFE y MPC sí exigen enteros de 1 a 4.
5. MIE: se calcula solo a partir de EFI y EFE ya completados y válidos; no se permite escribir los totales a mano (punto 4 de Elaboration II, sigue abierto). Si falta alguno, la sección indica cuál matriz completar.
6. Exportación: estructura provisional de dos hojas ("Datos" y "Resultados"), solo valores, sin imagen del gráfico y con el nombre `Mtx-<SIGLA>.xlsx`. Sin datos válidos no se exporta. Los puntos 1 y 5 de Elaboration II siguen abiertos.
7. Persistencia: una sola clave (`mtx.estado`), versión 1 del formato y estado global; un estado corrupto se descarta sin aviso. Además, una matriz sin ningún dato no muestra errores hasta que el usuario escribe. El punto 6 de Elaboration II sigue abierto.
8. Navegadores: la verificación se hizo solo en Chrome. La lista de navegadores soportados (Historia 7) sigue abierta.
9. Carpeta `tests/`: falta decidir si se conserva en el repositorio o si la verificación debe hacerse fuera de él.
10. Diseño visual: los estilos son mínimos y funcionales; el diseño visual definitivo sigue sin definirse.

## Construction III

Objetivo de la fase: dejar un plan de pruebas manual, pensado para que una persona lo siga con el navegador real, el mouse y el teclado. Esta fase no escribe ni modifica código de la aplicación. No cubre GE más allá de confirmar que está pendiente, porque esa historia no tiene lógica.

Estado del plan: ningún caso ha sido ejecutado por una persona, por eso todas las casillas "¿Pasó?" están en blanco. Los textos y números esperados de los bloques 1, 2, 3 y 6 se obtuvieron simulando esos mismos pasos sobre `index.html` con jsdom (un navegador simulado), y los de exportación y persistencia se comprobaron antes en Construction II. Eso no reemplaza la ejecución real, sobre todo en Edge y Firefox y con el archivo descargado de internet, que aún no se han probado.

### 0. Preparación y convenciones

Registro de ejecución (completar antes de empezar):

| Dato | Valor |
|---|---|
| Persona que ejecuta | |
| Fecha | |
| Equipo y versión de Windows | |
| Navegador y versión | |
| Ruta del archivo `index.html` probado | |
| Resultado global | |

Convenciones:

- **Abrir la aplicación:** con doble clic en `index.html`. Sale en el navegador predeterminado; para probar otro, clic derecho, "Abrir con".
- **Reinicio de datos (RD):** la aplicación guarda todo en el navegador y no tiene un botón para borrar datos, así que para volver a un estado limpio hay que hacer lo siguiente:
  1. Pulsar F12 y abrir la pestaña "Consola".
  2. Escribir a mano (si el navegador pide escribir `allow pasting`, hágalo) `localStorage.removeItem('mtx.estado'); location.reload()` y pulsar Enter.
  3. Cerrar las herramientas con F12.
  Salvo que un caso diga lo contrario, cada caso parte de un RD.
- **Estado limpio:** BCG visible; 2 divisiones en BCG; 2 fortalezas y 2 debilidades en EFI; 2 oportunidades y 2 amenazas en EFE; 2 factores, "Mi empresa" y "Competidor 1" en MPC; 2 factores por eje en PEYEA.
- **Agregar y quitar filas:** con los botones "+ Agregar división", "+ Agregar fortaleza", "+ Agregar debilidad", "+ Agregar oportunidad", "+ Agregar amenaza", "+ Agregar factor" y "+ Agregar competidor". El botón ✕ quita una fila. En MPC no hay ✕ para "Mi empresa".
- **Escribir datos:** con la tecla Tab se salta de un campo al siguiente. Los campos numéricos aceptan punto o coma decimal.
- **Cómo leer los resultados esperados:** los textos entre comillas se comparan tal como aparecen en pantalla. Los números salen con dos decimales. En EFI y EFE la tabla de resultados lista los factores en el orden en que fueron creados (las filas agregadas con "+" quedan al final), por eso conviene escribir el nombre de cada factor.
- **Gráfico del BCG:** alta participación relativa a la izquierda, alto crecimiento arriba; Estrella arriba a la izquierda, Interrogante arriba a la derecha, Vaca lechera abajo a la izquierda, Perro abajo a la derecha.
- **Aviso ámbar:** los casos límite muestran un recuadro amarillo con un texto que dice que la regla está pendiente de confirmar. Los errores de validación salen en un recuadro rojo.
- **¿Pasó?:** marque con una X ☐ Sí o ☐ No. Si marca No, anote qué vio en el margen o en un informe aparte.

### 1. Datos de prueba

Son los mismos números de las pruebas de Elaboration I.

**D1: BCG (prueba 1.1)**, en el orden Nombre, Ingresos, Utilidades, Participación relativa, Crecimiento:

| Fila | Nombre | Ingresos | Utilidades | Part. relativa | Crecimiento |
|---|---|---|---|---|---|
| 1 | A | 500 | 100 | 1.80 | 15 |
| 2 | B | 300 | 30 | 0.40 | 12 |
| 3 | C | 150 | 60 | 1.50 | 4 |
| 4 | D | 50 | 10 | 0.30 | 2 |

**D2: EFI (prueba 2.1)**, en el orden Nombre, Peso, Clasificación:

| Grupo | Nombre | Peso | Clasificación |
|---|---|---|---|
| Fortaleza | F1 | 0.20 | 4 |
| Fortaleza | F2 | 0.15 | 4 |
| Fortaleza | F3 | 0.10 | 3 |
| Debilidad | D1 | 0.25 | 1 |
| Debilidad | D2 | 0.20 | 2 |
| Debilidad | D3 | 0.10 | 1 |

**D3: EFE (prueba 2.2)**, en el orden Nombre, Peso, Clasificación:

| Grupo | Nombre | Peso | Clasificación |
|---|---|---|---|
| Oportunidad | O1 | 0.25 | 4 |
| Oportunidad | O2 | 0.20 | 3 |
| Oportunidad | O3 | 0.10 | 2 |
| Amenaza | A1 | 0.25 | 3 |
| Amenaza | A2 | 0.15 | 2 |
| Amenaza | A3 | 0.05 | 1 |

**D4: MPC (prueba 3.1)**: cuatro factores y tres empresas. En la primera fila de la tabla, cambie los nombres a "Competidor A" y "Competidor B".

| Factor crítico | Peso | Mi empresa | Competidor A | Competidor B |
|---|---|---|---|---|
| Participación de mercado | 0.30 | 3 | 4 | 2 |
| Competitividad de precios | 0.25 | 2 | 3 | 4 |
| Calidad del producto | 0.25 | 4 | 3 | 2 |
| Lealtad del cliente | 0.20 | 3 | 2 | 4 |

**D5: PEYEA (pruebas 4.1 a 4.4)**: valores de cada eje. La cantidad de filas de cada eje es la cantidad de valores.

| Caso | FF (fuerza financiera) | FI (fuerza de la industria) | VC (ventaja competitiva) | EE (estabilidad del entorno) |
|---|---|---|---|---|
| 4.1 agresivo | 5, 4, 4, 3 | 4, 5, 3 | −2, −3, −1, −2 | −3, −2, −4, −3, −3 |
| 4.2 conservador | 5, 5, 5 | 1, 2 | −4, −4 | −2, −1, −3 |
| 4.3 defensivo | 1, 2 | 1, 1 | −4, −4, −4 | −5, −5 |
| 4.4 competidor | 1, 2, 3 | 5, 5, 5 | −1, −2 | −5, −5, −5 |

**D6: recetas para obtener un total exacto de EFI o EFE.** La MIE no tiene campos propios: usa los totales de EFI y EFE. Para lograr un total exacto, rellene las cuatro filas del estado limpio en este orden (fila 1 y 2 son fortalezas u oportunidades; fila 3 y 4 son debilidades o amenazas), escribiendo "Peso" y "Clasificación" en cada una:

| Total buscado | Matriz | Fila 1 | Fila 2 | Fila 3 | Fila 4 |
|---|---|---|---|---|---|
| 3.20 | EFI | 0.8 y 3 | 0.2 y 4 | 0 y 1 | 0 y 1 |
| 1.80 | EFI | 0.8 y 2 | 0.2 y 1 | 0 y 1 | 0 y 1 |
| 1.99 | EFI | 0.99 y 2 | 0.01 y 1 | 0 y 1 | 0 y 1 |
| 2.00 | EFI | 1 y 2 | 0 y 1 | 0 y 1 | 0 y 1 |
| 2.99 | EFI | 0.99 y 3 | 0.01 y 2 | 0 y 1 | 0 y 1 |
| 3.00 | EFI | 1 y 3 | 0 y 1 | 0 y 1 | 0 y 1 |
| 1.995 | EFI | 0.995 y 2 | 0.005 y 1 | 0 y 1 | 0 y 1 |
| 3.50 | EFE | 0.5 y 4 | 0.5 y 3 | 0 y 1 | 0 y 1 |
| 1.50 | EFE | 0.5 y 2 | 0.5 y 1 | 0 y 1 | 0 y 1 |

**PH: prueba de humo** (se usa en los casos de despliegue), partiendo de un RD:

1. BCG con D1: los cuadrantes son A Estrella, B Interrogante, C Vaca lechera y D Perro, y salen "Total de ingresos: 1000.00. Total de utilidades: 200.00."
2. EFI con D2: sale "Total EFI: 2.45. Diagnóstico: Posición interna débil."
3. EFE con D3: sale "Total EFE: 2.90. Diagnóstico: La organización aprovecha oportunidades y evita amenazas por encima del promedio."
4. MIE: sale celda V, "Retener y mantener".
5. Recargar con F5: los datos y resultados siguen ahí.
6. Con EFI activo, pulsar "Exportar a Excel": se descarga `Mtx-EFI.xlsx` y se abre con "Total EFI" 2.45.

### 2. Bloque 1: pruebas de Elaboration I implementadas (CP-01 a CP-15)

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-01 | Prueba 1.1 (BCG) | RD. BCG visible. | 1. Pulse "+ Agregar división" dos veces (4 filas).<br>2. Escriba los datos D1.<br>3. Observe los resultados y el gráfico.<br>4. En "Tamaño de la burbuja según" elija "% de utilidades". | Sin recuadro rojo. Resultados: A Estrella, 50.00 % de ingresos y 50.00 % de utilidades; B Interrogante, 30.00 % y 15.00 %; C Vaca lechera, 15.00 % y 30.00 %; D Perro, 5.00 % y 5.00 %. Texto "Total de ingresos: 1000.00. Total de utilidades: 200.00." Gráfico: A (verde) en el cuadrante superior izquierdo "Estrella", B (amarillo) arriba a la derecha "Interrogante", C (azul) abajo a la izquierda "Vaca lechera", D (rojo) abajo a la derecha "Perro"; con % de ingresos las burbujas van de mayor a menor A, B, C, D. Con % de utilidades la leyenda dice "Tamaño de la burbuja: % de utilidades" y el orden es A, C, B, D (C es mayor que B). | ☐ Sí<br>☐ No |
| CP-02 | Prueba 2.1 (EFI) | RD. | 1. Vaya a EFI.<br>2. Pulse "+ Agregar fortaleza" y "+ Agregar debilidad" una vez cada uno (3 y 3).<br>3. Escriba los datos D2. | Sin recuadro rojo ni ámbar. Suma de pesos 1.00. Ponderados: F1 0.80, F2 0.60, F3 0.30, D1 0.25, D2 0.40, D3 0.10. Fila de total con 1.00 en Peso y 2.45 en Ponderado. Texto "Total EFI: 2.45. Diagnóstico: Posición interna débil." Sin gráfico. | ☐ Sí<br>☐ No |
| CP-03 | Prueba 2.2 (EFE) | RD. | 1. Vaya a EFE.<br>2. Pulse "+ Agregar oportunidad" y "+ Agregar amenaza" una vez cada uno (3 y 3).<br>3. Escriba los datos D3. | Sin recuadro rojo ni ámbar. Suma de pesos 1.00. Ponderados: O1 1.00, O2 0.60, O3 0.20, A1 0.75, A2 0.30, A3 0.05. Total 2.90. Texto "Total EFE: 2.90. Diagnóstico: La organización aprovecha oportunidades y evita amenazas por encima del promedio." Sin gráfico. | ☐ Sí<br>☐ No |
| CP-04 | Prueba 2.3 (EFI, pesos que no suman 1) | Datos de CP-02 cargados en EFI. | 1. En EFI, cambie el peso de D3 de 0.10 a 0.05.<br>2. Observe.<br>3. Vuelva a escribir 0.10. | Con 0.05: recuadro rojo con "La suma de los pesos es 0.95 y debe ser 1.00.", desaparecen la tabla de resultados, el total y el diagnóstico, y en su lugar dice "Complete o corrija los datos para ver el resultado." Con 0.10: el error desaparece y vuelve "Total EFI: 2.45. Diagnóstico: Posición interna débil." | ☐ Sí<br>☐ No |
| CP-05 | Prueba 2.4 (tolerancia en pesos 0.6, 0.3, 0.1) | RD. | 1. Vaya a EFE.<br>2. Pulse ✕ en la segunda amenaza (quedan 3 filas: 2 oportunidades y 1 amenaza).<br>3. Escriba peso y clasificación: 0.6 y 4; 0.3 y 3; 0.1 y 2. | Sin recuadro rojo (la suma 0.6 + 0.3 + 0.1 se acepta como 1). Suma de pesos 1.00. Ponderados 2.40, 0.90 y 0.20. Texto "Total EFE: 3.50. Diagnóstico: La organización aprovecha oportunidades y evita amenazas por encima del promedio." | ☐ Sí<br>☐ No |
| CP-06 | Prueba 3.1 (MPC) | RD. | 1. Vaya a MPC.<br>2. Pulse "+ Agregar factor" dos veces (4 factores) y "+ Agregar competidor" una vez (3 empresas).<br>3. Escriba los datos D4 (incluidos los nombres de factores y de empresas). | Sin recuadro rojo. Total: peso 1.00; Mi empresa 3.00; Competidor A 3.10; Competidor B 2.90. Ponderados de Mi empresa 0.90, 0.50, 1.00, 0.60; de A 1.20, 0.75, 0.75, 0.40; de B 0.60, 1.00, 0.50, 0.80. Texto "Mi empresa queda en la posición 2 de 3." Ranking: 1.º Competidor A 3.10, 2.º Mi empresa 3.00, 3.º Competidor B 2.90. Sin gráfico. | ☐ Sí<br>☐ No |
| CP-07 | Prueba 3.2 (MPC, pesos compartidos) | Datos de CP-06 cargados en MPC. | 1. Cambie los cuatro pesos a 0.10, 0.50, 0.20 y 0.20 (solo en la columna Peso). | Los tres totales se recalculan sin tocar nada más: Mi empresa 2.70, Competidor A 2.90, Competidor B 3.40. Texto "Mi empresa queda en la posición 3 de 3." Ranking: 1.º Competidor B 3.40, 2.º Competidor A 2.90, 3.º Mi empresa 2.70. | ☐ Sí<br>☐ No |
| CP-08 | Prueba 4.1 (PEYEA agresivo) | RD. | 1. Vaya a PEYEA.<br>2. Agregue filas hasta tener FF 4, FI 3, VC 4 y EE 5.<br>3. Escriba los valores de D5, fila 4.1. | Sin recuadro rojo ni ámbar. Promedios: FF 4.00, FI 4.00, VC -2.00, EE -3.00. Texto "X = VC + FI = 2.00. Y = EE + FF = 1.00. Vector: (2.00, 1.00)." y "Cuadrante: agresivo." Gráfico: flecha desde el origen hasta el punto (2.00, 1.00) en el cuadrante superior derecho "Agresivo". | ☐ Sí<br>☐ No |
| CP-09 | Prueba 4.2 (PEYEA conservador) | RD. | 1. En PEYEA, agregue o quite filas hasta tener FF 3, FI 2, VC 2 y EE 3.<br>2. Escriba los valores de D5, fila 4.2. | Promedios: FF 5.00, FI 1.50, VC -4.00, EE -2.00. "X = VC + FI = -2.50. Y = EE + FF = 3.00. Vector: (-2.50, 3.00)." "Cuadrante: conservador." Punto en el cuadrante superior izquierdo. | ☐ Sí<br>☐ No |
| CP-10 | Prueba 4.3 (PEYEA defensivo) | RD. | 1. En PEYEA, agregue o quite filas hasta tener FF 2, FI 2, VC 3 y EE 2.<br>2. Escriba los valores de D5, fila 4.3. | Promedios: FF 1.50, FI 1.00, VC -4.00, EE -5.00. "X = VC + FI = -3.00. Y = EE + FF = -3.50. Vector: (-3.00, -3.50)." "Cuadrante: defensivo." Punto en el cuadrante inferior izquierdo. | ☐ Sí<br>☐ No |
| CP-11 | Prueba 4.4 (PEYEA competitivo, versión corregida) | RD. | 1. En PEYEA, agregue o quite filas hasta tener FF 3, FI 3, VC 2 y EE 3.<br>2. Escriba los valores de D5, fila 4.4. | Promedios: FF 2.00, FI 5.00, VC -1.50, EE -5.00. "X = VC + FI = 3.50. Y = EE + FF = -3.00. Vector: (3.50, -3.00)." "Cuadrante: competitivo." Punto en el cuadrante inferior derecho. | ☐ Sí<br>☐ No |
| CP-12 | Prueba 5.1 (MIE con los totales de 2.1 y 2.2) | EFI con D2 y EFE con D3 cargados (CP-02 y CP-03). | 1. Pulse "MIE". | Sin recuadro rojo. Datos: "Total EFI" 2.45 (promedio), "Total EFE" 2.90 (promedio), "Celda" V, "Zona" Retener y mantener. Texto "Celda V: retener y mantener." Gráfico: cuadrícula de nueve celdas con la celda V (centro, amarilla) resaltada y un círculo "Ud." en ella. | ☐ Sí<br>☐ No |
| CP-13 | Prueba 5.2 (MIE esquina de crecimiento) | RD. | 1. En EFI escriba la receta D6 de 3.20.<br>2. En EFE escriba la receta D6 de 3.50.<br>3. Pulse "MIE". | Total EFI 3.20 (fuerte), Total EFE 3.50 (fuerte), celda I, zona "Crecer y construir". Celda I (arriba a la izquierda, verde) resaltada. | ☐ Sí<br>☐ No |
| CP-14 | Prueba 5.3 (MIE esquina de cosecha) | RD. | 1. En EFI escriba la receta D6 de 1.80.<br>2. En EFE escriba la receta D6 de 1.50.<br>3. Pulse "MIE". | Total EFI 1.80 (débil), Total EFE 1.50 (débil), celda IX, zona "Cosechar o desinvertir". Celda IX (abajo a la derecha, roja) resaltada. | ☐ Sí<br>☐ No |
| CP-15 | Prueba 5.4 (MIE en los límites de rango) | RD. En EFE la receta de 3.50 escrita. | 1. Repita cuatro veces: en EFI escriba la receta D6 de 1.99, luego 2.00, luego 2.99 y luego 3.00.<br>2. Después de cada una, pulse "MIE" y anote celda y zona. | 1.99: "(débil)", celda III, "Retener y mantener". 2.00: "(promedio)", celda II, "Crecer y construir". 2.99: "(promedio)", celda II, "Crecer y construir". 3.00: "(fuerte)", celda I, "Crecer y construir". En los cuatro, Total EFE 3.50 (fuerte). | ☐ Sí<br>☐ No |

### 3. Bloque 2: criterios provisionales implementados en Construction II (CP-16 a CP-30)

En estos casos el sistema debe aplicar el criterio documentado en la sección 4 de Construction II y mostrar el aviso correspondiente. Los avisos ámbar dicen que la regla está pendiente de confirmar con el curso: es el comportamiento esperado.

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-16 | EFI con total exactamente 2.5 | RD. | 1. Vaya a EFI.<br>2. Escriba peso y clasificación en las cuatro filas: 0.25 y 3; 0.25 y 3; 0.25 y 2; 0.25 y 2. | Sin recuadro rojo. Ponderados 0.75, 0.75, 0.50, 0.50. "Total EFI: 2.50. Diagnóstico: Posición interna fuerte." Recuadro ámbar: "El total es exactamente 2.5. La regla del curso para este caso está pendiente de confirmar." | ☐ Sí<br>☐ No |
| CP-17 | EFE con total exactamente 2.5 | RD. | 1. Vaya a EFE.<br>2. Escriba peso y clasificación en las cuatro filas: 0.25 y 3; 0.25 y 3; 0.25 y 2; 0.25 y 2. | "Total EFE: 2.50. Diagnóstico: La organización aprovecha oportunidades y evita amenazas por encima del promedio." Recuadro ámbar con el mismo texto que en CP-16. | ☐ Sí<br>☐ No |
| CP-18 | BCG con valor exactamente en el umbral | RD. | 1. En BCG, deje los nombres vacíos y escriba: fila 1 con 100, 20, 1, 10; fila 2 con 100, 20, 0.5, 5.<br>2. En la fila 1, cambie a participación 0.99 y crecimiento 10.<br>3. Cambie a participación 1 y crecimiento 9.99.<br>4. Cambie a participación 1.01 y crecimiento 10.01. | Paso 1: "División 1" Estrella y "División 2" Perro, cada una con 50.00 % y 50.00 %; recuadro ámbar "Alguna división está exactamente en un umbral (participación relativa 1.0 o crecimiento 10 %). Su cuadrante depende de la convención del curso, que está pendiente de confirmar." Paso 2: División 1 Interrogante, con aviso. Paso 3: División 1 Vaca lechera, con aviso. Paso 4: División 1 Estrella, sin aviso. | ☐ Sí<br>☐ No |
| CP-19 | BCG con suma de utilidades cero | RD. | 1. En BCG, fila 1 con 100, 10, 1.5, 15 y fila 2 con 100, -10, 0.5, 5.<br>2. Cambie la utilidad de la fila 2 a -20. | Recuadro rojo con "La suma de las utilidades debe ser mayor que 0 para calcular su porcentaje." en ambos pasos. No hay tabla ni gráfico; en su lugar dice "Complete o corrija los datos para ver el resultado." | ☐ Sí<br>☐ No |
| CP-20 | BCG con una utilidad negativa y suma positiva | RD. | 1. En BCG, fila 1 con 100, 50, 1.5, 15 y fila 2 con 100, -10, 0.5, 5.<br>2. En "Tamaño de la burbuja según" elija "% de utilidades". | Sin recuadro rojo. División 1 Estrella con 50.00 % de ingresos y 125.00 % de utilidades; División 2 Perro con 50.00 % y -25.00 %. "Total de ingresos: 200.00. Total de utilidades: 40.00." En el gráfico, la burbuja de División 2 se ve con el tamaño mínimo (muy pequeña) y la de División 1 es grande. | ☐ Sí<br>☐ No |
| CP-21 | PEYEA con X = 0 y Y = 0 | RD. | 1. En PEYEA, con las dos filas de cada eje: FF 3 y 3; FI 3 y 3; VC -3 y -3; EE -3 y -3. | "X = VC + FI = 0.00. Y = EE + FF = 0.00. Vector: (0.00, 0.00)." "Cuadrante: agresivo." Recuadro ámbar "El vector queda sobre un eje. La asignación de cuadrante en ese caso está pendiente de confirmar con el curso." En el gráfico solo se ve el punto en el origen, sin flecha. | ☐ Sí<br>☐ No |
| CP-22 | PEYEA con X = 0 y Y negativo | RD. | 1. En PEYEA, con las dos filas de cada eje: FF 1 y 1; FI 3 y 3; VC -3 y -3; EE -3 y -3. | "X = VC + FI = 0.00. Y = EE + FF = -2.00. Vector: (0.00, -2.00)." "Cuadrante: competitivo." Recuadro ámbar igual que en CP-21. Flecha hacia abajo sobre el eje vertical. | ☐ Sí<br>☐ No |
| CP-23 | PEYEA con Y = 0 y X negativo | RD. | 1. En PEYEA, con las dos filas de cada eje: FF 3 y 3; FI 1 y 1; VC -4 y -4; EE -3 y -3. | "X = VC + FI = -3.00. Y = EE + FF = 0.00. Vector: (-3.00, 0.00)." "Cuadrante: conservador." Recuadro ámbar igual que en CP-21. Flecha hacia la izquierda sobre el eje horizontal. | ☐ Sí<br>☐ No |
| CP-24 | MIE con un total como 1.995 | RD. | 1. En EFI escriba la receta D6 de 1.995.<br>2. En EFE escriba la receta D6 de 3.50.<br>3. Observe EFI y luego pulse "MIE". | EFI: "Total EFI: 2.00. Diagnóstico: Posición interna débil." (la pantalla redondea a dos decimales, por eso el 1.995 se ve como 2.00 y los pesos 0.995 y 0.005 se ven como 0.99 y 0.01). MIE: "Total EFI" 2.00 (débil), "Total EFE" 3.50 (fuerte), celda III, "Retener y mantener". Lo que se comprueba es el criterio documentado: 1.995 es menor que 2.0, así que cuenta como débil. Anote en observaciones que la pantalla muestra 2.00 junto a "débil". | ☐ Sí<br>☐ No |
| CP-25 | MPC con empate | RD. | 1. Vaya a MPC.<br>2. Escriba pesos 0.5 y 0.5 en los dos factores.<br>3. Escriba 3 y 3 para Mi empresa y 3 y 3 para Competidor 1. | Totales 3.00 y 3.00. "Mi empresa queda en la posición 1 de 2." En el ranking, "Mi empresa" y "Competidor 1" aparecen ambas con posición 1 y total 3.00. | ☐ Sí<br>☐ No |
| CP-26 | MPC sin competidores | Datos de CP-25 cargados. | 1. Pulse ✕ junto a "Competidor 1" (primera fila de la tabla).<br>2. Busque un ✕ junto a "Mi empresa". | Recuadro rojo con "Ingrese mi empresa y al menos un competidor." y sin tabla de resultados ("Complete o corrija los datos para ver el resultado."). No existe ✕ junto a "Mi empresa". | ☐ Sí<br>☐ No |
| CP-27 | EFI sin restricción de clasificación por tipo de factor | RD. | 1. En EFI escriba peso y clasificación: fortalezas 0.4 y 1, 0.1 y 2; debilidades 0.3 y 1, 0.2 y 2. | Sin recuadro rojo (una fortaleza con clasificación 1 se acepta). Ponderados 0.40, 0.20, 0.30, 0.40. "Total EFI: 1.30. Diagnóstico: Posición interna débil." | ☐ Sí<br>☐ No |
| CP-28 | Matriz sin ningún dato no muestra errores | RD. | 1. Pulse uno por uno BCG, EFI, EFE, MPC, PEYEA y MIE, sin escribir nada.<br>2. En BCG, escriba solo 100 en Ingresos de la fila 1. | BCG, EFI, EFE, MPC y PEYEA: sin recuadro rojo y con el texto "Complete o corrija los datos para ver el resultado." MIE: recuadro rojo con "Complete la matriz EFI para ubicar la empresa en la MIE." y "Complete la matriz EFE para ubicar la empresa en la MIE." Al escribir el 100 en BCG aparece el recuadro rojo "Hay campos vacíos: complete todos los campos antes de calcular." | ☐ Sí<br>☐ No |
| CP-29 | Nombres opcionales | RD. | 1. En EFI escriba los datos de CP-27 sin escribir ningún nombre.<br>2. Escriba "Marca" como nombre de la primera fila. | Las filas de resultados se llaman "Factor 1", "Factor 2", "Factor 3" y "Factor 4" (sin recuadro rojo por falta de nombres). Al escribir el nombre, la primera fila pasa a llamarse "Marca". | ☐ Sí<br>☐ No |
| CP-30 | Coma decimal | RD. | 1. En EFI escriba pesos con coma: 0,20; 0,30; 0,25; 0,25.<br>2. Escriba clasificaciones 4, 3, 2 y 1. | Sin recuadro rojo. En la tabla de resultados los pesos salen con punto (0.20, 0.30, 0.25, 0.25). Ponderados 0.80, 0.90, 0.50, 0.25. "Total EFI: 2.45. Diagnóstico: Posición interna débil." En los campos se sigue viendo lo que se escribió (con coma). | ☐ Sí<br>☐ No |

### 4. Bloque 3: Gran Estrategia (CP-31)

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-31 | GE pendiente (decisión del juez, Historia 6) | RD. Cualquier otra matriz activa. | 1. Pulse "GE".<br>2. Intente pulsar "Exportar a Excel".<br>3. Recargue con F5 (con GE activa).<br>4. Pulse otra matriz, por ejemplo "BCG". | Paso 1: aparece el mensaje "Módulo pendiente de definir con el curso: la Gran Estrategia (GE) todavía no tiene definidos el origen de sus dos ejes, su escala ni su punto de corte (Historia 6). Por ahora esta sección no calcula ni dibuja nada." No hay formulario, ni tabla, ni gráfico, ni número, ni cuadrante, ni recuadro rojo. Paso 2: el botón "Exportar a Excel" está deshabilitado (atenuado) y no descarga nada. Paso 3: tras recargar sigue en GE con el mismo mensaje y el botón deshabilitado. Paso 4: el botón "Exportar a Excel" vuelve a estar habilitado. | ☐ Sí<br>☐ No |

### 5. Bloque 4: navegación, persistencia y exportación (CP-32 a CP-46)

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-32 | Navegación: estado inicial | RD. | 1. Observe la página recién cargada. | Hay siete botones: BCG, EFI, EFE, MPC, PEYEA, MIE y GE. El botón BCG está resaltado (fondo azul, texto blanco) y solo se ve la sección BCG con 2 filas. "Exportar a Excel" está habilitado. | ☐ Sí<br>☐ No |
| CP-33 | Navegación: un botón por matriz | RD. | 1. Pulse en orden BCG, EFI, EFE, MPC, PEYEA, MIE, GE y otra vez BCG.<br>2. Después de cada clic, observe el título de la sección y el botón resaltado. | Después de cada clic solo se ve la sección elegida, con su título (BCG, EFI, etc.), y solo ese botón aparece resaltado. Ninguna otra sección se ve. | ☐ Sí<br>☐ No |
| CP-34 | Navegación: no se pierden datos | RD. | 1. En BCG escriba 100 en Ingresos de la fila 1.<br>2. Pulse EFI y escriba 0.5 en el peso de la primera fila.<br>3. Vuelva a BCG y luego a EFI. | En BCG sigue el 100 y en EFI sigue el 0.5. | ☐ Sí<br>☐ No |
| CP-35 | Reutilización: la MIE sigue a EFI y EFE | EFI con D2 y EFE con D3 cargados; MIE mostrada con celda V (CP-12). | 1. Vaya a EFI y cambie la clasificación de F1 de 4 a 1.<br>2. Pulse "MIE".<br>3. Vuelva a EFI, devuelva la clasificación de F1 a 4 y pulse "MIE". | Paso 2: "Total EFI" 1.85 (débil), "Total EFE" 2.90 (promedio), celda VI, zona "Cosechar o desinvertir". Paso 3: vuelve a 2.45 (promedio), celda V, "Retener y mantener". | ☐ Sí<br>☐ No |
| CP-36 | Persistencia: recargar conserva datos y resultados | RD. | 1. Escriba los datos D1 en BCG, D2 en EFI (3 y 3 filas), D3 en EFE (3 y 3 filas) y D4 en MPC.<br>2. Deje EFI como matriz activa.<br>3. Pulse F5.<br>4. Revise EFI, BCG, EFE, MPC y MIE. | Después de F5 la matriz activa sigue siendo EFI, con los mismos datos y "Total EFI: 2.45. Diagnóstico: Posición interna débil." Al pulsar los demás botones, BCG, EFE y MPC conservan sus datos y sus resultados (los mismos de CP-01, CP-03 y CP-06) y MIE muestra celda V. | ☐ Sí<br>☐ No |
| CP-37 | Persistencia: datos incompletos | RD. | 1. En BCG escriba 100 en Ingresos de la fila 1 y nada más.<br>2. Pulse F5. | Tras recargar, el 100 sigue en su campo y aparece el recuadro rojo "Hay campos vacíos: complete todos los campos antes de calcular." | ☐ Sí<br>☐ No |
| CP-38 | Persistencia: cerrar la pestaña y volver | RD. | 1. Escriba en PEYEA los datos de la fila 4.1 de D5 y deje PEYEA activa.<br>2. Cierre la pestaña (y si puede, todo el navegador).<br>3. Vuelva a abrir `index.html` con doble clic. | Se abre en PEYEA (no en BCG) con los mismos valores y los resultados de CP-08: "Vector: (2.00, 1.00)." y "Cuadrante: agresivo." | ☐ Sí<br>☐ No |
| CP-39 | Persistencia: estado guardado dañado | Datos guardados de cualquier caso. | 1. Pulse F12, abra la Consola y escriba a mano `localStorage.setItem('mtx.estado', '{no es json')` y Enter.<br>2. Pulse F5.<br>3. Escriba un valor en BCG y pulse F5 otra vez. | Paso 2: la página abre en BCG con formularios vacíos, sin recuadro rojo ni mensaje de error en la página y sin errores en la consola. Paso 3: el valor escrito se conserva (la aplicación vuelve a guardar normalmente). | ☐ Sí<br>☐ No |
| CP-40 | Exportación EFI | EFI con D2 cargado y EFI activo. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo descargado con Excel, LibreOffice u otro programa de hojas de cálculo. | Se descarga `Mtx-EFI.xlsx` y se abre sin errores. Tiene dos hojas: "Datos" y "Resultados". Datos: encabezado Tipo, Factor, Peso, Clasificación y seis filas (Fortaleza F1 0.2 4, Fortaleza F2 0.15 4, Debilidad D1 0.25 1, Debilidad D2 0.2 2, Fortaleza F3 0.1 3, Debilidad D3 0.1 1, en el orden en que se crearon las filas). Resultados: encabezado Tipo, Factor, Peso, Clasificación, Ponderado; seis filas con ponderados 0.8, 0.6, 0.25, 0.4, 0.3, 0.1; una fila vacía; "Suma de pesos" 1; "Total EFI" 2.45; "Diagnóstico" Posición interna débil. Los números son celdas numéricas. | ☐ Sí<br>☐ No |
| CP-41 | Exportación BCG | BCG con D1 cargado y BCG activo. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo. | `Mtx-BCG.xlsx` con hojas "Datos" y "Resultados". Datos: encabezado División, Ingresos, Utilidades, Participación relativa, Crecimiento del mercado (%) y las filas A, B, C, D con los números de D1. Resultados: encabezado División, Cuadrante, % de ingresos, % de utilidades; A Estrella 50 50; B Interrogante 30 15; C Vaca lechera 15 30; D Perro 5 5; una fila vacía; "Total ingresos" 1000; "Total utilidades" 200. | ☐ Sí<br>☐ No |
| CP-42 | Exportación EFE | EFE con D3 cargado y EFE activo. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo. | `Mtx-EFE.xlsx` con hojas "Datos" y "Resultados", con la misma estructura que CP-40. En Resultados: ponderados 1, 0.6, 0.75, 0.3, 0.2, 0.05; "Suma de pesos" 1; "Total EFE" 2.9; "Diagnóstico" La organización aprovecha oportunidades y evita amenazas por encima del promedio. | ☐ Sí<br>☐ No |
| CP-43 | Exportación MPC | MPC con D4 cargado y MPC activo. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo. | `Mtx-MPC.xlsx`. Datos: encabezado Factor crítico, Peso, Mi empresa, Competidor A, Competidor B y cuatro filas con pesos y clasificaciones de D4. Resultados: encabezado con "(ponderado)" en cada empresa, cuatro filas de ponderados, fila "Total" con 1, 3, 3.1 y 2.9, una fila vacía y el ranking (Empresa, Total, Posición): Competidor A 3.1 1; Mi empresa 3 2; Competidor B 2.9 3. | ☐ Sí<br>☐ No |
| CP-44 | Exportación PEYEA | PEYEA con la fila 4.1 de D5 cargada y PEYEA activo. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo. | `Mtx-PEYEA.xlsx`. Datos: encabezado Eje, Factor, Valor y 16 filas (4 de Fuerza financiera (FF), 3 de Fuerza de la industria (FI), 4 de Ventaja competitiva (VC) y 5 de Estabilidad del entorno (EE)) con los valores de D5. Resultados: promedios 4, 4, -2, -3; una fila vacía; "X (VC + FI)" 2; "Y (EE + FF)" 1; "Cuadrante" Agresivo. | ☐ Sí<br>☐ No |
| CP-45 | Exportación MIE | EFI con D2 y EFE con D3 cargados; MIE activa. | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo. | `Mtx-MIE.xlsx`. Datos: "Total EFI" 2.45 y "Total EFE" 2.9. Resultados: "Nivel EFI" Promedio, "Nivel EFE" Promedio, "Celda" V, "Zona" Retener y mantener. | ☐ Sí<br>☐ No |
| CP-46 | Exportación sin datos válidos | RD. | 1. Con BCG vacío, pulse "Exportar a Excel".<br>2. Vaya a EFI, escriba los datos de D2 (3 y 3 filas) y ponga 0.05 en el peso de D3; pulse "Exportar a Excel".<br>3. Devuelva el peso a 0.10 y pulse "Exportar a Excel". | Paso 1: no se descarga ningún archivo y aparece el recuadro rojo "Complete los datos antes de exportar." Paso 2: no se descarga ningún archivo y aparece "La suma de los pesos es 0.95 y debe ser 1.00." Paso 3: se descarga `Mtx-EFI.xlsx` y el recuadro rojo desaparece. | ☐ Sí<br>☐ No |

### 6. Bloque 5: despliegue de la Historia 7 (CP-47 a CP-55)

Lista de comprobación de despliegue (LD), que se aplica en todos los casos de este bloque:

- **LD-1:** al abrir, no aparece ninguna ventana de advertencia de macros, de seguridad de Windows (por ejemplo SmartScreen o "Windows protegió su PC") ni solicitud de permisos del navegador (ubicación, notificaciones, cámara, micrófono ni almacenamiento).
- **LD-2:** se ve la página "Mtx" con los siete botones y la sección BCG.
- **LD-3:** con F12 y la pestaña "Consola", después de cargar y después de la prueba de humo no hay mensajes de error (en rojo). Los avisos amarillos o mensajes de extensiones del navegador no cuentan, pero anótelos.
- **LD-4:** con F12 y la pestaña "Red" (Network), tras recargar con F5 solo aparece el propio archivo `index.html`, sin ninguna solicitud a direcciones `http://` o `https://` ni a ningún dominio. Si el navegador pide un ícono (`favicon.ico`), anótelo.
- **LD-5:** la prueba de humo PH da los resultados esperados en cada paso.

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-47 | Historia 7: doble clic en Windows con Chrome | Windows con Chrome. Copia local de `index.html` (no descargada de internet). | 1. Haga doble clic en `index.html` con Chrome como navegador predeterminado (o clic derecho, "Abrir con", Chrome).<br>2. Aplique LD-1 a LD-5. | Se cumplen LD-1 a LD-5 en Chrome. | ☐ Sí<br>☐ No |
| CP-48 | Historia 7: Edge | Windows con Edge. Misma copia local. | 1. Abra `index.html` con Edge.<br>2. Aplique LD-1 a LD-5.<br>3. Compruebe que los datos escritos en Chrome no aparecen (cada navegador guarda por separado). | Se cumplen LD-1 a LD-5 en Edge. En Edge se empieza con un estado limpio, lo cual es esperado. | ☐ Sí<br>☐ No |
| CP-49 | Historia 7: Firefox | Windows con Firefox. Misma copia local. | 1. Abra `index.html` con Firefox.<br>2. Aplique LD-1 a LD-5 (en Firefox la pestaña de la consola y la de red se llaman "Consola" y "Red").<br>3. Compruebe que los datos de otro navegador no aparecen. | Se cumplen LD-1 a LD-5 en Firefox. El estado empieza limpio, lo cual es esperado. | ☐ Sí<br>☐ No |
| CP-50 | Historia 7: archivo descargado de internet, Chrome | Copia de `index.html` descargada de internet (por ejemplo, desde el repositorio privado en GitHub con "Download raw file", o recibida por correo o WhatsApp y descargada). | 1. Clic derecho sobre el archivo, "Propiedades": compruebe que al pie aparece el aviso de seguridad "Este archivo proviene de otro equipo..." con la casilla "Desbloquear" (esa es la marca de la web). No la marque.<br>2. Ábralo con Chrome.<br>3. Aplique LD-1 a LD-5. | La marca de la web está presente y se cumplen LD-1 a LD-5 en Chrome sin desbloquear el archivo. | ☐ Sí<br>☐ No |
| CP-51 | Historia 7: archivo descargado de internet, Edge | La misma copia descargada, con la casilla "Desbloquear" sin marcar. | 1. Ábrala con Edge.<br>2. Aplique LD-1 a LD-5. | Se cumplen LD-1 a LD-5 en Edge. | ☐ Sí<br>☐ No |
| CP-52 | Historia 7: archivo descargado de internet, Firefox | La misma copia descargada, con la casilla "Desbloquear" sin marcar. | 1. Ábrala con Firefox.<br>2. Aplique LD-1 a LD-5. | Se cumplen LD-1 a LD-5 en Firefox. | ☐ Sí<br>☐ No |
| CP-53 | Historia 7: sin conexión a internet | Copia local. Desconecte la red (modo avión, o desactive Wi-Fi y Ethernet). | 1. Con la red desconectada, abra `index.html`.<br>2. Ejecute la prueba de humo PH completa, incluida la exportación.<br>3. Aplique LD-3 y LD-4. | Todo funciona igual que con conexión, incluido el paso 6 de PH (se descarga el .xlsx). Ningún error de consola por recursos no encontrados. | ☐ Sí<br>☐ No |
| CP-54 | Historia 7: equipo sin Office ni Access | Equipo o laboratorio sin Microsoft Office ni Access instalados (si no hay uno disponible, anótelo). | 1. Abra `index.html` con doble clic.<br>2. Ejecute PH.<br>3. Abra el `.xlsx` descargado con otro programa (por ejemplo LibreOffice, Google Sheets o Excel en línea). | La aplicación funciona sin pedir instalar nada. Los pasos 1 a 5 de PH dan lo esperado, y el archivo `Mtx-EFI.xlsx` se abre en el programa alternativo con "Total EFI" 2.45. | ☐ Sí<br>☐ No |
| CP-55 | Historia 7 e Inception: carpeta movida o en otra ubicación | Copia de la carpeta del proyecto. | 1. Copie la carpeta a cuatro lugares: el Escritorio, una memoria USB, una carpeta sincronizada con OneDrive o SharePoint y una ruta con espacios y acentos (por ejemplo `C:\Prueba de Matrices\Gestión`).<br>2. En cada una, abra `index.html` con doble clic y ejecute PH pasos 1 a 5.<br>3. Anote si los datos guardados en una copia aparecen o no en otra. | En las cuatro ubicaciones la aplicación abre y funciona sin ninguna configuración ni cambio de rutas. Lo que ocurre con los datos guardados entre copias depende del navegador: anótelo. | ☐ Sí<br>☐ No |

### 7. Bloque 6: validación y robustez (CP-56 a CP-60), agregado

No estaba en la lista pedida. Son funciones ya implementadas que los bloques anteriores no ejercitan de forma directa.

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-56 | PEYEA rechaza el 0 y valores fuera de rango | RD. | 1. En PEYEA escriba en la primera fila: FF 0, FI 7, VC 0, EE -7.<br>2. En la segunda fila de cada eje escriba 3, 3, -3, -3. | Recuadro rojo con exactamente cuatro mensajes: "Fuerza financiera (FF), factor 1: el valor debe estar entre 1 y 6 (el 0 no es válido).", "Fuerza de la industria (FI), factor 1: el valor debe estar entre 1 y 6 (el 0 no es válido).", "Ventaja competitiva (VC), factor 1: el valor debe estar entre -6 y -1 (el 0 no es válido).", "Estabilidad del entorno (EE), factor 1: el valor debe estar entre -6 y -1 (el 0 no es válido)." No hay resultados ni gráfico. | ☐ Sí<br>☐ No |
| CP-57 | EFI rechaza clasificaciones y pesos inválidos | RD. | 1. En EFI escriba el nombre F1 en la primera fila y estos peso y clasificación: fila 1 0.5 y 5; fila 2 1.2 y 2.5; fila 3 abc y 0; fila 4 0.1 y 2. | Recuadro rojo con cinco mensajes: "Factor "F1": la clasificación debe ser un número entero de 1 a 4.", "Factor 2: la clasificación debe ser un número entero de 1 a 4.", "Factor 3: la clasificación debe ser un número entero de 1 a 4.", "Factor 2: el peso debe ser un número entre 0 y 1." y "Factor 3: el peso debe ser un número entre 0 y 1." No aparece el mensaje de suma de pesos ni resultados. | ☐ Sí<br>☐ No |
| CP-58 | Los nombres se muestran como texto, no como código | RD. | 1. En EFI escriba como nombre de la primera fila `<b>Marca</b>` y los datos de CP-30 en peso y clasificación. | En la tabla de resultados la fila aparece con el texto literal `<b>Marca</b>` (con los símbolos visibles y sin negrita). No se rompe la página ni aparecen errores en la consola. | ☐ Sí<br>☐ No |
| CP-59 | Agregar y quitar filas mantiene las clasificaciones alineadas (MPC) | RD. | 1. En MPC pulse "+ Agregar factor" una vez (3 factores).<br>2. Escriba pesos 0.5, 0.3 y 0.2; Mi empresa 4, 3, 2; Competidor 1 con 2, 2, 2.<br>3. Pulse ✕ junto al segundo factor.<br>4. Cambie el peso del último factor a 0.5. | Paso 2: Totales 3.30 (Mi empresa) y 2.00 (Competidor 1); "Mi empresa queda en la posición 1 de 2." Paso 3: recuadro rojo "La suma de los pesos es 0.70 y debe ser 1.00." y sin resultados; las clasificaciones que quedan son las del primer y tercer factor (Mi empresa 4 y 2; Competidor 1 con 2 y 2). Paso 4: totales 3.00 y 2.00; posición 1 de 2. | ☐ Sí<br>☐ No |
| CP-60 | MIE indica qué matriz completar o corregir | RD. | 1. Pulse "MIE" con EFI y EFE vacíos.<br>2. Escriba los datos D2 en EFI (3 y 3 filas), con EFE aún vacío; pulse "MIE".<br>3. Escriba los datos D3 en EFE (3 y 3 filas) y ponga 0.05 en el peso de D3 en EFI; pulse "MIE".<br>4. Devuelva el peso a 0.10 y pulse "MIE". | Paso 1: dos mensajes rojos "Complete la matriz EFI para ubicar la empresa en la MIE." y "Complete la matriz EFE para ubicar la empresa en la MIE." Paso 2: un solo mensaje "Complete la matriz EFE para ubicar la empresa en la MIE." Paso 3: "Corrija la matriz EFI para ubicar la empresa en la MIE: La suma de los pesos es 0.95 y debe ser 1.00." En los pasos 1 a 3 no hay gráfico. Paso 4: celda V, "Retener y mantener". | ☐ Sí<br>☐ No |

### 8. Resumen de cobertura

| Bloque | Casos | Rango |
|---|---|---|
| 1. Pruebas de Elaboration I implementadas | 15 | CP-01 a CP-15 |
| 2. Criterios provisionales de Construction II | 15 | CP-16 a CP-30 |
| 3. Gran Estrategia (GE) | 1 | CP-31 |
| 4. Navegación (4), persistencia (4) y exportación (7) | 15 | CP-32 a CP-46 |
| 5. Despliegue de la Historia 7 | 9 | CP-47 a CP-55 |
| 6. Validación y robustez (agregado) | 5 | CP-56 a CP-60 |
| Total | 60 | CP-01 a CP-60 |

Elaboration I tiene 15 pruebas con lógica implementada (1.1, 2.1 a 2.4, 3.1, 3.2, 4.1 a 4.4 y 5.1 a 5.4), y las 15 están en el bloque 1. La prueba 6.1 (GE) y la 7.1 (uso sin instalación) se cubren en los bloques 3 y 5. No se escribió ningún caso para funcionalidad que no existe: no hay caso de GE con datos, de exportación de GE, de entrada manual de totales en la MIE ni de gráficos de EFI, EFE o MPC.

### Observaciones y puntos nuevos marcados [VERIFICAR] en esta fase

Observaciones al preparar el plan (no se cambió código):

1. La aplicación no tiene un botón para borrar datos (`Persistencia.limpiar` existe, pero no está conectado a la interfaz), así que reiniciar un caso exige usar la consola del navegador. [VERIFICAR] si conviene agregar ese botón.
2. Cuando un total cae entre 1.995 y 2.00 (CP-24), la pantalla muestra 2.00 pero el sistema lo clasifica como débil, porque redondea solo al mostrar. Es coherente con el criterio documentado, pero puede confundir a un estudiante. [VERIFICAR] si los totales deben mostrarse con más decimales cuando están cerca de un umbral.
3. Con una utilidad negativa en una división, el porcentaje de utilidades puede pasar de 100 % o ser negativo (CP-20). Es la consecuencia aritmética del criterio provisional del BCG, no un error de cálculo. [VERIFICAR] junto con el punto 2 de Construction II.
4. Como la MIE no tiene campos propios, las pruebas 5.2, 5.3 y 5.4 necesitan las recetas de la sección D6 (pesos con ceros) para lograr totales exactos. Esto es consecuencia de que Elaboration II dejó abierta la entrada manual de totales.

Puntos [VERIFICAR] nuevos:

1. Alcance del criterio de red (LD-4): si el navegador pide `favicon.ico` para `file://`, hay que decidir si eso cuenta como solicitud de red. Se propone que no, siempre que no sea a un dominio externo.
2. Datos guardados entre copias del archivo en distintas carpetas (CP-55): depende del navegador y no está definido cuál es el comportamiento deseado.
3. Programa de hojas de cálculo con el que se valida el `.xlsx` (CP-40 a CP-45, CP-54): el plan pide "Excel u otro". Falta decidir si el curso exige compatibilidad con Excel específicamente.
4. Versiones mínimas de navegadores: el plan usa las versiones actuales de Chrome, Edge y Firefox. La lista de navegadores soportados de la Historia 7 sigue abierta.
5. Quién y con qué frecuencia ejecuta el plan, y cómo se registran los resultados y los fallos: hoy solo hay la columna "¿Pasó?" y el registro de ejecución.

## Transition

Objetivo de la fase: publicar la aplicación, verificarla en su dirección real, dejar material de apoyo para explicar el código, reunir en una sola lista lo que sigue abierto para llevarlo al profesor y dejar las preguntas de reflexión que pide VUP. Esta fase no cambió `index.html` ni `tests/`.

### 1. Configuración de despliegue

| Dato | Valor |
|---|---|
| URL de la aplicación | https://gerson-chumpitaz.github.io/mtx/ |
| Repositorio | https://github.com/gerson-chumpitaz/mtx, ahora **público** (antes era privado). |
| Origen de la publicación | Rama `master`, carpeta raíz (`/`). El archivo que se sirve en la raíz es `index.html`. |
| Tipo de publicación | Por rama (`build_type` = `legacy`): cada `git push` a `master` vuelve a publicar solo. |
| HTTPS | Forzado por GitHub (`https_enforced` = `true`). |

Pasos exactos que se siguieron, con la cuenta `gerson-chumpitaz` ya autenticada en `gh`:

1. Se revisó la sintaxis vigente con `gh repo edit --help`: cambiar la visibilidad exige `--visibility` junto con `--accept-visibility-change-consequences`.
2. Antes de cambiarla se comprobó qué se hace público: los cuatro archivos rastreados (`.gitignore`, `VUP.md`, `index.html` y `tests/elaboration1.test.js`) y los 12 commits. Ninguno contiene claves ni tokens. Los commits llevan como autor el correo personal `gersonechumpitazd@gmail.com`, que queda visible en el historial público.
3. Se hizo público el repositorio:

```bash
gh repo edit gerson-chumpitaz/mtx --visibility public --accept-visibility-change-consequences
```

4. Se comprobó en la documentación oficial (docs.github.com, REST API de Pages, "Create a GitHub Pages site") el endpoint vigente: `POST /repos/{owner}/{repo}/pages`, con el cuerpo `build_type` (`legacy` o `workflow`) y `source.branch` (obligatorio) y `source.path` (`/` o `/docs`, por defecto `/`). Una consulta previa al mismo endpoint respondió 404 porque el sitio aún no existía.
5. Se creó el sitio:

```bash
gh api --method POST repos/gerson-chumpitaz/mtx/pages -H "Accept: application/vnd.github+json" -f "build_type=legacy" -f "source[branch]=master" -f "source[path]=/"
```

   La respuesta trajo `html_url` = `https://gerson-chumpitaz.github.io/mtx/`.
6. Se esperó la publicación consultando cada 15 segundos hasta que el estado pasó de `building` a `built` (algo menos de un minuto):

```bash
gh api repos/gerson-chumpitaz/mtx/pages/builds/latest
```

7. Se comprobó que el archivo servido es idéntico al del repositorio: el SHA-256 de la URL y el de `index.html` es `19502baf90bfb6a934290fc35613463907c3c19b22aeae79b2a04ee94f2c4dc9` (1 020 715 bytes).

Para repetir o deshacer: un nuevo `git push origin master` publica de nuevo sin más pasos. Para desactivar el sitio: `gh api --method DELETE repos/gerson-chumpitaz/mtx/pages`. GitHub Pages sirve todo lo que hay en la raíz, así que `VUP.md` y `tests/elaboration1.test.js` también quedan accesibles públicamente en la misma dirección.

### 2. Resultado de la verificación en la URL real

Se abrió https://gerson-chumpitaz.github.io/mtx/ en el navegador integrado de la aplicación (Chromium) y se repitió una verificación básica.

| Comprobación | Resultado |
|---|---|
| La página carga | Sí: HTTP 200, HTTPS y contenido idéntico al de `index.html` del repositorio. |
| Las siete matrices son accesibles | Sí. Recorrido BCG, EFI, EFE, MPC, PEYEA, MIE, GE y otra vez BCG: en cada paso solo se ve la sección elegida y solo su botón queda resaltado. El botón de exportar se deshabilita únicamente en GE. |
| Prueba 1.1 del BCG | Coincide con Elaboration I: A Estrella 50.00 % y 50.00 %, B Interrogante 30.00 % y 15.00 %, C Vaca lechera 15.00 % y 30.00 %, D Perro 5.00 % y 5.00 %; "Total de ingresos: 1000.00. Total de utilidades: 200.00."; cuatro burbujas dibujadas. |
| Exportación a Excel | Sí: se generó `Mtx-BCG.xlsx` de 18 350 bytes con las hojas "Datos" y "Resultados", con los mismos valores de la prueba 1.1 (totales 1000 y 200). Se comprobó leyendo el archivo generado con SheetJS, no abriéndolo en Excel. |
| Persistencia | Tras recargar la URL, los datos de las cuatro divisiones (por ejemplo, Ingresos 500 en la fila 1) y los resultados siguen ahí. |
| Errores de consola | Ninguno. |
| Solicitudes de red | Solo la del propio documento (`GET https://gerson-chumpitaz.github.io/mtx/`, HTTP 200, una por carga). Ningún recurso externo ni CDN, y no hay enlaces ni scripts externos en la página. |
| Diferencias con abrir el archivo con `file://` | Ninguna de comportamiento, porque es el mismo archivo byte a byte. |

Qué no se verificó: la URL solo se probó en el navegador integrado (no en Chrome, Edge ni Firefox por separado ni en un teléfono), los datos se escribieron con eventos simulados y no con el teclado, y la exportación se validó leyendo el archivo generado, no abriéndolo en Excel. Nada apareció roto y no se tocó `index.html`.

Diferencias de entorno a tener presentes (no son fallos):

- Los datos guardados en la URL pública viven en el origen `https://gerson-chumpitaz.github.io`, separado de los que se guardan al abrir el archivo con `file://` y de los de otros navegadores.
- Ese origen lo comparten otros sitios de Pages de la misma cuenta. La clave que usa Mtx (`mtx.estado`) es específica, pero otro sitio del mismo origen podría leerla o pisarla.

### 3. Resumen del código generado

Este resumen es material de apoyo para poder explicar el código; no reemplaza que se lea. Los números de línea son de `index.html` en el commit `45d963a` (el último que lo modificó, con Radar Estratégico ya incluido; antes eran los del commit `f3e4464`, que ya incluía Análisis Estructural) y cambiarán si el archivo se edita.

**Patrones de diseño**

- **Objetos de responsabilidad única.** Cada componente es un objeto literal (`const Validador = { ... }`) con métodos sobre un solo tema, tal como lo definía el diagrama de clases de Elaboration II. No hay clases con herencia ni instancias múltiples: cada uno existe una sola vez.
- **Orquestador único.** Solo la Vista llama a los otros cinco componentes; ninguno conoce a otro. Esto se parece al patrón mediador o controlador: reduce las dependencias cruzadas y hace que el orden de las llamadas se lea en un solo lugar.
- **Funciones puras.** `Validador` y `MotorCalculo` reciben datos y devuelven resultados nuevos sin leer ni escribir en la pantalla. Por eso `tests/elaboration1.test.js` los prueba en Node sin navegador.
- **Una sola fuente de verdad.** Lo que el usuario escribió vive en la variable `estado`. Los resultados no se guardan: se recalculan cada vez que hacen falta, así que no pueden quedar desactualizados respecto de los datos.
- **Delegación de eventos y configuración por atributos.** Hay solo tres oyentes en `document` (`click`, `input`, `change`). Cada elemento de la página dice qué hace con atributos `data-accion` (por ejemplo `agregar`, `quitar`, `navegar`, `exportar`) y `data-campo` (la ruta del dato, por ejemplo `divisiones.0.ingresos`, `matriz.0.1` en Análisis Estructural o `calificaciones.0` en Radar Estratégico), de modo que no hay un oyente por campo.
- **Avisos como banderas en el resultado.** Los casos límite pendientes de confirmar salen del `MotorCalculo` como una propiedad (`enLimite`, `enEje`) y la Vista decide mostrar el aviso. El cálculo no sabe nada de la pantalla.
- **Persistencia defensiva.** Cada acceso a `localStorage` va dentro de `try/catch` y el estado guardado se revisa antes de usarse.
- **Estado por componente en lugar de estado por matriz.** En Radar Estratégico cada uno de los 14 componentes tiene su propio estado (vacío, incompleto, completo o inválido), que calcula `calcularRadar`. El módulo nunca es "inválido" como un todo: un error en un componente no frena el cálculo ni el gráfico de los otros trece.
- **Colocación de los números de puntaje en la primera posición libre.** `dibujarRadar` prueba hasta ocho posiciones alrededor de cada punto y toma la primera que no pise otro texto ni otro punto, empezando por los puntos más alejados del centro. Los anchos de texto son estimados (siete unidades por carácter) y no medidos, así que el resultado no depende de la tipografía. Si ninguna posición queda libre, el número no se dibuja.
- **Disposición en dos columnas con CSS.** Solo para RADAR: una cuadrícula (CSS Grid) con el formulario a la izquierda y un panel con errores y gráfico a la derecha, fijo con `position: sticky`, dentro de una consulta de medios de 75rem. Por debajo de ese ancho todo va en una sola columna. No interviene ningún JavaScript.

No es una arquitectura MVC formal, pero se le parece: el modelo es el `estado` más los motores, la vista son las funciones que dibujan, y el controlador es `Vista.despacharEvento`.

**Cómo está organizado el archivo** (un solo `index.html`):

| Líneas aproximadas | Contenido |
|---|---|
| 7 a 87 | Estilos (`<style>`), mínimos y funcionales, con un bloque propio para la disposición de RADAR en dos columnas. |
| 90 a 104 | Encabezado con el `<nav id="navegacion">` (nueve botones: siete matrices, AE y RADAR) y el botón "Exportar a Excel". |
| 106 a 180 | Nueve secciones `<section class="matriz">` (las siete del Módulo 1, la de Análisis Estructural y la de Radar Estratégico). Ocho llevan cuatro contenedores sueltos: `formulario`, `errores`, `resultados` y `grafico`. La de RADAR lleva `formulario`, un `panel-radar` que agrupa `errores` y `grafico`, y `resultados`. |
| 182 a 212 | `<script id="sheetjs">`: la librería SheetJS 0.20.3 completa, sin modificar. |
| 214 a 2042 | `<script id="app">`: el código de la aplicación. |

Dentro del script de la aplicación, en este orden: constantes y funciones auxiliares (líneas 217 a 390); Validador (392 a 655); MotorCalculo (657 a 886); MotorGraficos (888 a 1223); Persistencia (1225 a 1341); Exportador (1343 a 1445); Vista con sus funciones auxiliares privadas (1447 a 2026); y la rutina de arranque `arrancar()` (2028 a 2041). Análisis Estructural y Radar Estratégico no tienen un bloque propio: sus piezas están repartidas en el lugar de cada componente. Las de AE llevan el sufijo `AE` (`erroresAE` y sus dos auxiliares en el Validador, `clasificarAE` y `calcularAE` en el MotorCalculo, `dibujarAE` en el MotorGraficos, `aeVacio` y `aeValido` en Persistencia, la rama `"AE"` en `hojasExportacion`, y `formularioAE`, `resultadosAE` y `hayCalificacionesAE` en la Vista). Las de RADAR llevan el sufijo `Radar` o el nombre `RADAR`: las constantes `ETAPAS_RADAR`, `TEXTOS_RADAR` (el título y las afirmaciones del profesor), `ESTRUCTURA_RADAR` (los 14 componentes con su rango de índices planos), `TOTAL_CARACTERISTICAS_RADAR` y `ESCALA_RADAR` entre las constantes; `calificacionRadarValida` y `erroresRadar` en el Validador; `calcularRadar` en el MotorCalculo; `dibujarRadar` en el MotorGraficos; `radarVacio` y `radarValido` en Persistencia; la rama `"RADAR"` en `hojasExportacion`; y `formularioRadar`, `resultadosRadar` y `hayCalificacionesRadar` en la Vista. Existe una función auxiliar, `obtenerPorRuta`, que está definida y no se usa en ninguna parte.

**Los seis componentes**

| Componente | Responsabilidad | Métodos |
|---|---|---|
| Vista | Dibuja formularios y resultados, escribe los errores y despacha los eventos del usuario. Es el único que llama a los demás. En AE no suma ningún método: `renderFormulario('AE')` usa `formularioAE` y `renderResultados('AE', resultado)` usa `resultadosAE`, que dibuja la tabla de ranking con el selector SÍ o NO y la hoja Validadas. En RADAR tampoco suma ningún método: `renderFormulario('RADAR')` usa `formularioRadar`, que dibuja las 56 calificaciones agrupadas por etapa y por componente, con el título y la afirmación del profesor, y `renderResultados('RADAR', resultado)` usa `resultadosRadar`, la tabla de los 14 componentes. | `renderFormulario`, `renderResultados`, `mostrarErrores`, `despacharEvento` |
| Validador | Comprueba los datos antes de calcular: campos vacíos, rangos numéricos y pesos que suman 1. Devuelve `{ valido, errores }` y nunca lanza excepciones. En AE no suma ningún método: `validar('AE', datos)` usa `erroresAE` y revisa solo las celdas fuera de la diagonal. En RADAR tampoco suma ningún método público: `validar('RADAR', datos)` usa `erroresRadar`, no llama a `validarCamposVacios` (una característica sin calificar nunca es un error), revisa solo que cada calificación con valor sea un entero de 0 a 5 y devuelve además `indicesInvalidos`. | `validar`, `validarPesos`, `validarRango`, `validarCamposVacios` |
| MotorCalculo | Aplica las fórmulas y clasificaciones de las siete matrices del Módulo 1, de Análisis Estructural (motricidad, dependencia, cortes, cuadrante y proyección) y de Radar Estratégico (el estado y el promedio de cada uno de sus 14 componentes). GE no está implementada. | `calcularBCG`, `calcularEFI`, `calcularEFE`, `calcularMPC`, `calcularPEYEA`, `ubicarMIE`, `ubicarGE`, `calcularAE`, `calcularRadar` |
| MotorGraficos | Dibuja los gráficos en SVG dentro del contenedor `grafico` de la matriz. `dibujarAE` traza el plano de AE en un recuadro cuadrado, con los cuatro cuadrantes, las líneas de corte, la diagonal de igualdad y la línea de proyección de cada punto. `dibujarRadar` traza el radar de 14 puntas, con el eje fijo de 0 a 5, cinco anillos, un rótulo por punta, un punto por componente completo y su número de puntaje, que se coloca en la primera de ocho posiciones libres. | `dibujarBCG`, `dibujarPEYEA`, `dibujarMIE`, `dibujarGE`, `dibujarAE`, `dibujarRadar` |
| Persistencia | Guarda y recupera el estado en `localStorage` con la clave `mtx.estado`. Si un estado guardado antes de AE no trae `datos.AE` (o lo trae dañado), `cargar` lo completa con AE vacío sin cambiar la versión del formato. Igual con RADAR: si no trae `datos.RADAR` (o lo trae dañado, se comprueba con `radarValido`), lo completa con `radarVacio`, 56 `null`. | `guardar`, `cargar`, `limpiar` |
| Exportador | Arma un libro de Excel con SheetJS (hojas "Datos" y "Resultados") y dispara la descarga. Tiene una rama para AE y otra para RADAR. | `exportarXLSX` |

**Cómo se comunican.** Por llamadas directas a métodos, pasando objetos simples. La Vista toma un dato del `estado`, se lo entrega al Validador, con lo que este devuelve decide si llama al MotorCalculo, y con el resultado llama a `renderResultados` y al MotorGraficos. Después de cada cambio le entrega el `estado` a Persistencia. El Exportador recibe `(matriz, datos, resultado)` que le arma la Vista. Entre los componentes no hay eventos propios ni variables compartidas; solo la Vista y sus funciones auxiliares usan las variables `estado` y `contextoMatriz` (esta última le dice a `mostrarErrores` sobre qué matriz escribir).

**Cómo se maneja el error de validación.** Es un resultado normal, no una excepción. `Validador.validar` devuelve `{ valido: false, errores: [...] }` con mensajes en texto. La función `evaluarMatriz` traduce eso a uno de cuatro estados: `ok`, `vacia` (todavía no hay ningún dato, y no se muestran errores; en AE significa que ninguna celda fuera de la diagonal tiene una calificación), `invalida` o `pendiente` (GE). Si es `invalida`, la Vista llama a `mostrarErrores` para escribir la lista en el recuadro rojo de esa matriz, deja el texto "Complete o corrija los datos para ver el resultado.", vacía el gráfico y **no llama** al MotorCalculo. Los datos escritos se guardan igual, sean válidos o no. Los demás errores siguen caminos parecidos: `exportarActiva` captura las excepciones y las muestra en el mismo recuadro, y los fallos de `localStorage` se tragan (`guardar` devuelve `false`, `cargar` devuelve un estado vacío). `ubicarGE` y `dibujarGE` lanzan un error explícito de módulo pendiente, pero la interfaz nunca los llama.

Radar Estratégico es la excepción al camino anterior. Como cada uno de sus 14 componentes tiene su propio estado (vacío, incompleto, completo o inválido), `evaluarMatriz('RADAR')` nunca devuelve `invalida`: una calificación fuera de rango marca como "inválido" solo a su componente, y el resultado sale con estado `ok` junto con los mensajes de error, que `mostrarErrores` escribe en el recuadro rojo mientras se calculan y se grafican los otros componentes. El estado `vacia` sí existe en RADAR: ninguna de las 56 calificaciones tiene valor.

**Cómo seguir el flujo de la historia del BCG**, desde que el usuario escribe hasta que ve el resultado (los nombres son funciones o métodos que se pueden buscar en el archivo):

1. Al abrir la página, `arrancar()` (línea 2033) pide el estado a `Persistencia.cargar()` (1319), llama a `Vista.renderFormulario('BCG')` (1931), que arma la tabla de divisiones con `formularioBCG`, y registra los tres oyentes en `document`.
2. El usuario escribe `500` en Ingresos de la primera fila. El navegador dispara el evento `input`, que llega a `Vista.despacharEvento` (1985).
3. `despacharEvento` lee `data-campo` (`divisiones.0.ingresos`), deduce la matriz (`BCG`) de la sección donde está el campo y escribe el valor en `estado.datos.BCG` con `establecerPorRuta`.
4. `Persistencia.guardar(estado)` (1311) lo guarda en `localStorage`. Luego se llama a `actualizarMatriz('BCG')` (línea 1842).
5. `actualizarMatriz` llama a `evaluarMatriz('BCG')` (1804). Si no hay ningún dato escrito devuelve `vacia`. Si lo hay, llama a `Validador.validar('BCG', datos)` (594), que a su vez usa `validarCamposVacios` (644) y `erroresBCG` (396): campos vacíos, ingresos mayores o iguales a 0, participación relativa mayor que 0, crecimiento de -100 o más y suma de ingresos y de utilidades mayores que 0.
6. Si hay errores, el flujo termina en `mostrarErrores` (1976) como se explicó arriba. Si no, `evaluarMatriz` llama a `MotorCalculo.calcularBCG(divisiones)` (720): convierte los textos a números con `aNumero` (que acepta coma o punto), suma los totales, clasifica cada división con `clasificarBCG` (661) usando los umbrales 1.0 y 10 %, y calcula los porcentajes de ingresos y de utilidades. La Vista agrega al resultado la medida elegida para el tamaño de burbuja.
7. Con el estado `ok`, la Vista limpia los errores y llama a `renderResultados('BCG', resultado)` (1951), que arma la tabla y los totales con `resultadosBCG`, y a `MotorGraficos.dibujarBCG(resultado)` (914), que construye el SVG: cuadrantes, escala logarítmica en X y lineal en Y, y una burbuja por división con radio proporcional a la raíz de su porcentaje.
8. El usuario ve la tabla y el gráfico. Cada tecla repite los pasos 2 a 7, y al pulsar "Exportar a Excel" se llama a `exportarActiva` (1910), que repite `evaluarMatriz` y entrega el resultado a `Exportador.exportarXLSX` (1433).

Para verlo en acción con las herramientas del navegador (F12): en la pestaña de código fuente, ponga puntos de interrupción en `despacharEvento`, `evaluarMatriz`, `calcularBCG` y `dibujarBCG`, y escriba un valor en el BCG. Para ver el mismo cálculo sin navegador, la prueba 1.1 está en `tests/elaboration1.test.js`.

**Cómo seguir el flujo de la historia de Análisis Estructural (AE)**, con el mismo nivel de detalle y los mismos puntos de partida que el del BCG. Es una sola herramienta con cuatro etapas (variables, matriz, plano y selección) y comparte el recorrido de eventos del BCG, así que solo se detalla lo que cambia:

1. Al abrir la página, `arrancar()` (2033) pide el estado a `Persistencia.cargar()` (1319). Si el estado guardado es de antes de AE y no trae `datos.AE`, o lo trae dañado (se comprueba con `aeValido`, 1242), `cargar` lo reemplaza por `aeVacio` (1238): `{ variables: [], matriz: [], marcas: [] }`. Luego `renderFormulario('AE')` (1931) llama a `formularioAE` (1583), que muestra la lista de variables y, si hay alguna, la matriz de influencias.
2. El usuario pulsa "+ Agregar variable". El clic llega a `despacharEvento` (1985) con `data-accion="agregar"` y llama a `agregarFila` (1877). En la rama de AE suma un nombre vacío, una columna vacía a cada fila, una fila nueva con la diagonal en `null` y una marca `null`. "− Quitar la última variable" (`quitarFila`, 1895) hace lo contrario y siempre borra la última. Después se guarda el estado, se vuelve a dibujar el formulario (porque cambió la forma de la matriz) y se llama a `actualizarMatriz`.
3. El usuario escribe una calificación, por ejemplo un `4` en la celda de V1 sobre V2. El evento `input` llega a `despacharEvento`, que lee `data-campo` (`matriz.0.1`) y escribe el valor en `estado.datos.AE` con `establecerPorRuta`. Las celdas de la diagonal no son campos, así que no pueden escribirse. Los selectores SÍ o NO de la tabla de ranking usan `marcas.i`; el valor "—" se guarda como `null`. Cada cambio se guarda con `Persistencia.guardar` (1311) y llama a `actualizarMatriz('AE')` (1842).
4. `actualizarMatriz` llama a `evaluarMatriz('AE')` (1804). Si ninguna celda fuera de la diagonal tiene valor (`hayCalificacionesAE`, 1774, sobre `celdasFueraDeDiagonalAE`, 527), devuelve `vacia`: es el caso de una sola variable o de una matriz sin calificar. Si no, llama a `Validador.validar('AE', datos)` (594), que le pasa a `validarCamposVacios` (644) solo las celdas fuera de la diagonal y a `erroresAE` (545): al menos dos variables, matriz cuadrada y cada calificación entera de 0 a 4 (`calificacionAEValida`, 536), con un máximo de 20 mensajes por rango.
5. Si hay errores, el flujo termina en `mostrarErrores` (1976): recuadro rojo, sin tabla ni gráfico. Si no, `evaluarMatriz` arma las variables como `{ nombre, marca }` y llama a `MotorCalculo.calcularAE(variables, matriz)` (834). Este calcula la motricidad de cada variable (suma de su fila, sin la diagonal) y su dependencia (suma de su columna), los cortes (la mitad del máximo de cada una), el cuadrante con `clasificarAE` (708), donde un valor exactamente en el corte cuenta como alto, y la proyección en x, la proyección en y y el punto sobre la diagonal. Devuelve las variables en el orden en que se cargaron.
6. Con el estado `ok`, la Vista llama a `renderResultados('AE', resultado)` (1951), que arma con `resultadosAE` (1709) los cortes, la tabla de ranking con su selector, el recuadro ámbar sobre el orden pendiente y la hoja Validadas (la lista de solo lectura de las marcadas con SÍ, también en orden de carga). Después llama a `MotorGraficos.dibujarAE(resultado)` (1067), que construye el SVG: un recuadro cuadrado de 435 por 435, los cuatro cuadrantes rotulados, las dos líneas de corte, la diagonal de igualdad, una línea de proyección por variable y un punto con su rótulo V1, V2, V3... y un texto emergente con el nombre completo.
7. El usuario marca SÍ o NO. El evento `change` sigue el mismo camino de los pasos 3 a 6; como las marcas no cambian ningún valor calculado, solo cambian la columna del selector y la hoja Validadas. Al pulsar "Exportar a Excel" se llama a `exportarActiva` (1910), que repite `evaluarMatriz` y entrega el resultado a `Exportador.exportarXLSX` (1433), cuya rama `"AE"` está en `hojasExportacion` (1347).

Para verlo en acción con las herramientas del navegador (F12): ponga puntos de interrupción en `evaluarMatriz`, `erroresAE`, `calcularAE` y `dibujarAE`, y escriba una calificación en la matriz. Para ver el mismo cálculo sin navegador, las pruebas AE.1 a AE.10 están en `tests/elaboration1.test.js` (desde la línea 207).

**Cómo seguir el flujo de la historia de Radar Estratégico**, con el mismo nivel de detalle y los mismos puntos de partida que el del BCG y el de AE. Es una sola herramienta con 56 calificaciones fijas (no hay variables ni filas que agregar o quitar) y comparte el recorrido de eventos del BCG, así que solo se detalla lo que cambia. La diferencia de fondo con los otros módulos es que cada uno de sus 14 componentes tiene su propio estado y ningún error corta el flujo:

1. Al abrir la página, `arrancar()` (2033) pide el estado a `Persistencia.cargar()` (1319). Si el estado guardado es de antes de RADAR y no trae `datos.RADAR`, o lo trae dañado (se comprueba con `radarValido`, 1256: un arreglo `calificaciones` de exactamente 56 posiciones, cada una `null`, número o texto), `cargar` lo reemplaza por `radarVacio` (1252): `{ calificaciones: [null, ...] }` con 56 `null`. Luego `renderFormulario('RADAR')` (1931) llama a `formularioRadar` (1619), que recorre `ETAPAS_RADAR` y `ESTRUCTURA_RADAR` (333, los 14 componentes con su rango de índices planos) y dibuja, por etapa y por componente, una tabla con el título y las afirmaciones del profesor (`TEXTOS_RADAR`) y un campo por característica.
2. El usuario escribe una calificación, por ejemplo un `3` en la afirmación 2 de Traducción 1. El evento `input` llega a `despacharEvento` (1985), que lee `data-campo` (`calificaciones.13`, el índice plano de esa afirmación) y escribe el texto en `estado.datos.RADAR` con `establecerPorRuta`. No hay botones "+" ni "−": la estructura es fija. No se vuelve a dibujar el formulario, así que el cursor sigue en su campo. `Persistencia.guardar` (1311) guarda el estado y se llama a `actualizarMatriz('RADAR')` (1842).
3. `actualizarMatriz` llama a `evaluarMatriz('RADAR')` (1804). Si ninguna de las 56 calificaciones tiene valor (`hayCalificacionesRadar`, 1779) devuelve `vacia`: la pista, sin errores y sin gráfico. Si no, llama a `Validador.validar('RADAR', datos)` (594) y aquí está la excepción: en los demás módulos un error de validación devuelve `invalida` y corta el flujo, pero para RADAR `evaluarMatriz` sigue siempre al cálculo, llama a `MotorCalculo.calcularRadar` con las calificaciones y los índices inválidos y devuelve el estado `ok` con los mensajes en `errores` junto al resultado. En `actualizarMatriz`, con el estado `ok` se llama a `mostrarErrores` con esos mensajes (puede ser una lista vacía), a `renderResultados` y al gráfico.
4. `Validador.validar('RADAR', datos)` no llama a `validarCamposVacios` (644): una característica sin calificar es un estado normal, no un error. Llama a `erroresRadar` (580), que recorre los 14 componentes y, por cada calificación con valor que no pase `calificacionRadarValida` (572, un entero de 0 a 5 con `validarRango`), agrega un mensaje como "Traducción 1, característica 3: debe ser un número entero entre 0 y 5." y guarda su índice plano. `validar` devuelve `{ valido, errores, indicesInvalidos }`.
5. `MotorCalculo.calcularRadar(calificaciones, errores)` (872) recibe las 56 calificaciones y los índices inválidos (el segundo parámetro conserva el nombre `errores`, pero no son mensajes) y devuelve `ResultadoRadar`: un arreglo de 14 posiciones, en el orden de las puntas, cada una `{ estado, puntaje }`. El estado es "inválido" si alguna de sus características está entre los índices inválidos, "vacío" si ninguna tiene valor, "incompleto" si algunas sí y otras no, y "completo" si todas tienen valor. Solo en este último hay puntaje: el promedio sin redondear de sus calificaciones, con la cantidad de características del componente como divisor.
6. `renderResultados('RADAR', resultado)` (1951) llama a `resultadosRadar` (1736), que arma la tabla de 14 filas (etapa, componente, estado y puntaje a dos decimales solo en los completos) y una nota. `mostrarErrores` (1976) escribe el recuadro rojo, que en pantallas anchas queda a la derecha junto al gráfico. `MotorGraficos.dibujarRadar(resultado)` (1133) construye el SVG de 640 por 650: 14 puntas desde arriba y en sentido horario con su rótulo "<etapa> <posición>", cinco anillos, el eje fijo de 0 en el centro a 5 en el borde, los números de la escala sobre el eje que sale entre las puntas 11 y 12, un punto por componente completo con su texto emergente y un pie de seis líneas. El número de puntaje de cada punto se coloca en la primera de ocho posiciones alrededor de su punto que no pise ningún otro texto ni punto, empezando por los puntos más alejados del centro y con anchos de texto estimados (siete unidades por carácter). Si ninguna posición queda libre, ese número no se dibuja y el puntaje sigue en el texto emergente y en la tabla.
7. Al pulsar "Exportar a Excel" se llama a `exportarActiva` (1910), que repite `evaluarMatriz`. Si el módulo está vacío muestra "Complete los datos antes de exportar." y no exporta. En cualquier otro caso, incluidos componentes incompletos, vacíos o inválidos, entrega el resultado a `Exportador.exportarXLSX` (1433), cuya rama `"RADAR"` está en `hojasExportacion` (1347): la hoja "Datos" con las 56 calificaciones y el texto de su afirmación (un valor que no es un número se exporta como el texto escrito) y la hoja "Resultados" con el estado real de cada componente.

Para verlo en acción con las herramientas del navegador (F12): ponga puntos de interrupción en `evaluarMatriz`, `erroresRadar`, `calcularRadar` y `dibujarRadar`, y escriba una calificación en un campo. Para ver el mismo cálculo sin navegador, las pruebas RE.1 a RE.13 están en `tests/elaboration1.test.js` (desde la línea 368).

### 4. Lista consolidada de puntos [VERIFICAR] para llevar al profesor

Es un listado, no una fase de decisiones: ningún punto se resolvió aquí. Solo incluye lo que sigue abierto al 2026-10-05. Se excluyen los puntos que el juez ya cerró: las escalas de PEYEA (1 a 6 y −1 a −6, sin 0), el uso de un CDN (todo va embebido), la librería de Excel (SheetJS), la prueba del Exportador (se cubrió en Construction III), el nombre del archivo, SVG en vez de canvas, la rutina de arranque, el tipo del parámetro `matriz`, la navegación y la ubicación de la librería. Los puntos de Análisis Estructural se agregaron en la sección F con fecha 2026-10-02, y excluyen igualmente lo que el juez ya cerró en ese módulo: la regla del valor exactamente en el corte, la forma de `ResultadoAE`, cómo viajan las marcas hacia `calcularAE` y la geometría del gráfico (diagonal a 45° y proyecciones a 90°). Los puntos de Radar Estratégico se agregaron en la sección G con fecha 2026-10-05, y excluyen lo que el juez ya cerró en ese módulo: la exportación de un valor no numérico (se exporta como el texto escrito), la disposición en dos columnas y la colocación de los números de puntaje.

**Inception:** no marcó ningún punto como [VERIFICAR], pero sus riesgos piden confirmar con el profesor: (a) la terminología del modelo PE-BSC (análisis FLOR, ADN de misión y visión) y (b) la fidelidad a la terminología exacta de D'Alessio, que probablemente se usa para calificar. Están fuera del alcance de v1.

**A. Matrices: datos y reglas de cálculo**

| Matriz | Punto abierto | Fase de origen | Criterio provisional implementado |
|---|---|---|---|
| BCG | Si la herramienta debe calcular la participación relativa a partir de ventas de competidores o si la escribe el estudiante | Elaboration I | La escribe el estudiante |
| BCG | Umbrales (participación relativa 1.0 y crecimiento 10 %): fijos, configurables o al estilo D'Alessio (crecimiento centrado en 0 %) | Elaboration I | 1.0 y 10 %, fijos |
| BCG | Escala del eje X (logarítmica o lineal) y su orientación | Elaboration I | Logarítmica, alta participación a la izquierda |
| BCG | Valor exactamente en un umbral | Elaboration I | Cuenta como "alto"; aviso en pantalla |
| BCG | Utilidad negativa o cero: tamaño de la burbuja, porcentaje de utilidades mayor que 100 % o negativo, y rechazo de un conjunto con suma de utilidades cero o negativa (bloquea el caso de pérdidas totales) | Elaboration I, Construction II y III | Burbuja mínima; se rechaza si la suma es 0 o menos |
| BCG | Rango válido del crecimiento del mercado | Construction II | Mayor o igual a −100 % |
| EFI y EFE | Diagnóstico cuando el total es exactamente 2.5 | Elaboration I | Fuerte o "aprovecha"; aviso en pantalla |
| EFI y EFE | Restricción de clasificación por tipo de factor (D'Alessio usa 3 o 4 para fortalezas y 1 o 2 para debilidades) | Elaboration I | 1 a 4 para todos |
| EFI y EFE | Rechazo de clasificación fuera de 1 a 4 y de peso fuera de 0 a 1 | Elaboration I | Se rechaza |
| EFI y EFE | Mostrar los totales con más decimales cuando están cerca de un umbral (1.995 se ve como 2.00 pero cuenta como débil) | Construction III | Solo dos decimales |
| MPC | Cantidad mínima y máxima de competidores | Elaboration I | Mi empresa más al menos un competidor, sin máximo |
| MPC | Si la clasificación 1 a 4 también se valida | Elaboration I | Se valida |
| MPC | Manejo de empates | Elaboration I | Comparten posición |
| MPC, EFI y EFE | Si llevan gráfico además de la tabla (por ejemplo barras en el MPC) | Elaboration II | Solo tabla |
| PEYEA | Definición del vector: segmento del origen a (X, Y) o magnitud y ángulo | Elaboration I | Solo (X, Y) |
| PEYEA | Cuadrante cuando X = 0 o Y = 0 | Elaboration I | El 0 cuenta como positivo; aviso en pantalla |
| PEYEA | Cantidad de factores por eje | Elaboration I | Sin número fijo |
| PEYEA | Si se aceptan decimales (D'Alessio califica con enteros) | Construction II | Se aceptan |
| MIE | Numeración de celdas I a IX y asignación de zonas | Elaboration I | Según David y D'Alessio |
| MIE | Valor entre 1.99 y 2.00 (por ejemplo 1.995) | Elaboration I | Cortes en menos de 2.0 y menos de 3.0 |
| MIE | Si se permiten ingresar los totales a mano cuando EFI y EFE no están llenos | Elaboration II | No; se calcula solo desde EFI y EFE |
| GE | Origen de los dos ejes, su escala y el punto de corte | Elaboration I | No implementada |
| GE | Listas de estrategias por cuadrante (varían entre autores) | Elaboration I | Solo en la especificación, sin código |
| GE | Firmas de `ubicarGE` y `dibujarGE` | Elaboration II | Provisionales; lanzan error de módulo pendiente |
| General | Si el nombre de una división, factor o empresa es obligatorio | Construction II | Opcional |

**B. Exportación a Excel**

- Estructura del archivo (hojas, nombre, valores o fórmulas vivas, si incluye una imagen del gráfico). Elaboration II y Construction II. Provisional: dos hojas, solo valores, sin imagen, `Mtx-<SIGLA>.xlsx`.
- Si se puede exportar con datos incompletos o inválidos. Elaboration II. Provisional: se bloquea en las otras matrices. En Radar Estratégico no se bloquea, por decisión del juez en Construction II: cada componente se valida de forma independiente, y el archivo trae el estado real de cada uno.
- Con qué programa debe validarse el archivo (Excel específicamente u otro). Construction III.

**C. Persistencia y sesión**

- Mensajes de error para una matriz incompleta tras recargar, y aviso al descartar un estado corrupto. Elaboration II. Provisional: una matriz sin datos no muestra errores hasta que se escribe algo, y un estado corrupto se descarta sin aviso.
- Clave, estructura y versión del formato guardado, y si el estado es por matriz o global. Elaboration II. Provisional: una clave (`mtx.estado`), versión 1, estado global.
- Si conviene un botón para borrar los datos (hoy solo se puede desde la consola del navegador). Construction III.
- Qué debe pasar con los datos guardados entre copias del archivo en distintas carpetas. Construction III.

**D. Despliegue y compatibilidad**

- Navegadores y versiones mínimas que el curso debe soportar. Elaboration I, Construction II y III.
- Si un pedido de `favicon.ico` en `file://` cuenta como solicitud de red. Construction III. Se propone que no, si no es hacia un dominio externo.
- Soporte del selector `:has()` y de `position: sticky` en las versiones de Edge y Firefox del equipo donde se use la aplicación, que la disposición de RADAR en dos columnas necesita. Construction II del Módulo 3 (punto 7) y Construction III (casos CP-128 y CP-129). Si el navegador no los soporta, la sección de RADAR conserva 64rem de ancho y la columna del formulario queda angosta pero usable.
- Nuevo en esta fase: si es aceptable el origen compartido `gerson-chumpitaz.github.io` para los datos guardados, y que el historial público muestre el correo personal del autor de los commits.

**E. Proceso y repositorio**

- Si la carpeta `tests/` se conserva en el repositorio. Construction II.
- Diseño visual definitivo (hoy son estilos mínimos). Construction II.
- Quién ejecuta el plan de pruebas manual, con qué frecuencia y cómo se registran los resultados y los fallos. Construction III.
- El plan de pruebas manual de Análisis Estructural (CP-61 a CP-77, Construction III del Módulo 2) también queda sin ejecutar por una persona, igual que el del Módulo 1. Falta, sobre todo, Edge y Firefox, abrir el `.xlsx` de AE en un programa de hojas de cálculo y la sensación de velocidad con la matriz grande.
- El plan de pruebas manual de Radar Estratégico (CP-78 a CP-130, Construction III del Módulo 3) también queda sin ejecutar por una persona, igual que los anteriores. Falta, sobre todo, Edge y Firefox, abrir el `.xlsx` de RADAR en un programa de hojas de cálculo y la usabilidad de escribir las 56 calificaciones.

**F. Análisis Estructural (Módulo 2)**

| Punto abierto | Fase de origen | Criterio provisional implementado |
|---|---|---|
| Criterio de orden de la tabla de ranking y de la hoja Validadas. Nunca se pudo verificar: lo calcula una consulta guardada (`c03_Motricidad_Dependencia`) de la base Access del profesor, protegida con contraseña, que no se pudo abrir | Inception y Elaboration I | Orden de carga (V1, V2, V3...); aviso ámbar en pantalla |
| Si deben aceptarse calificaciones con decimales en la matriz de influencias | Elaboration I y Construction II | Se rechazan: solo enteros de 0 a 4 |
| Estructura de las hojas exportadas de AE. No pudo compararse con el Excel original por la misma contraseña | Elaboration II y Construction II | Dos hojas: "Datos" (matriz completa con los nombres de variable en la fila y la columna de encabezado, diagonal vacía) y "Resultados" (una fila por variable en orden de carga, más las filas Corte Y y Corte X); archivo `Mtx-AE.xlsx` |
| Máximo de variables que el curso necesita | Inception | Sin tope en el código; se comprobó con 200, el máximo de la guía del profesor |
| Rendimiento con una matriz grande: cuánto tiempo es aceptable al escribir y al recalcular | Construction II y III | Sin criterio. Cada tecla recalcula y redibuja todo el módulo; el plan manual (CP-77) pide anotar la experiencia sin fijar un umbral |

**G. Radar Estratégico (Módulo 3)**

| Punto abierto | Fase de origen | Criterio provisional implementado |
|---|---|---|
| Título repetido de los dos componentes de Alineamiento ("LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO", filas 69 y 74 del Excel). Falta que el profesor confirme el título correcto del segundo, cuyas características hablan de unidades de soporte | Inception | La aplicación los muestra tal cual, numerados Alineamiento 1 y Alineamiento 2, y los trata como dos componentes independientes |
| Característica 4 de Gestión 1, idéntica a la característica 4 de Motivación 3 ("Existe un mecanismo para premiar las inciativas y las sugerencias de los colaboradores") y sin relación con un componente sobre el presupuesto | Construction II | Transcrita tal cual |
| Erratas del texto del profesor ("inciativas", "seguiimiento", "establecidda", "l os EE-UN", "del la Estrategia", "EL superior" y el espacio antes de la coma en "estrategicas , actividades"), tildes que faltan y títulos con tildes inconsistentes | Construction II | Conservadas sin corregir, en el formulario y en el archivo exportado |
| Si el profesor prefiere otro control para calificar de 0 a 5 (un selector con los seis niveles y su texto) en lugar de un campo de texto | Construction III | Campo de texto con enteros de 0 a 5 |
| Si se aceptan calificaciones con decimales | Elaboration I y Construction II | Se rechazan: solo enteros de 0 a 5 |
| Criterios de usabilidad nunca acordados: cuánto tarda escribir las 56 calificaciones y cuánta página es aceptable (mide unos 5300 a 6300 px de alto según la ventana) | Construction III | Sin criterio. El plan manual (CP-116 y CP-117) pide anotar la experiencia sin fijar un umbral; en pantallas anchas el panel con errores y gráfico queda fijo a la vista |
| Si el curso necesita usar el módulo en un celular | Construction III | No se atiende, porque no hay requisito: a 390 px de ancho la página mide unos 10 000 px de alto y los rótulos del gráfico quedan en unos 6 px |
| Política de los números de puntaje del gráfico: se omiten los que no caben sin pisar otro texto o un punto (con puntajes muy bajos, como en D-RA4, se dibujan 4 de 14), y si el profesor prefiere un gráfico sin números | Construction III | Se dibujan los que caben; el puntaje siempre está en el texto emergente de cada punto y en la tabla de resultados |
| Aspecto del radar vacío, cómo se distingue en pantalla un componente "incompleto" de uno "vacío" y la redacción de los avisos y errores | Elaboration I y Construction II | Con el módulo vacío solo se ve la pista genérica, sin el marco del radar; incompleto y vacío solo se distinguen por la columna "Estado" de la tabla; la redacción de los avisos es libre |

### 5. Reflexión

Preguntas de VUP para responder a mano. No las respondió la IA.

**1. ¿Qué fue lo más importante de la especificación?**

Respuesta:

**2. ¿Qué harías distinto?**

Respuesta:

**3. ¿Qué te sorprendió de cómo la IA implementó los requisitos?**

Respuesta:

**4. ¿Cómo ayudó tener un plan de pruebas claro?**

Respuesta:

**5. ¿Qué agregarías si siguieras desarrollando el proyecto?**

Respuesta:

## Inception — Módulo 2: Análisis Estructural

Objetivo de la fase: abrir un segundo ciclo VUP dentro de este mismo documento, para el módulo de Análisis Estructural, con la misma estructura de la Inception original. Esta fase no escribe código, no define firmas de componentes ni pruebas Given-When-Then (eso es Elaboration I de este módulo) y no modifica `index.html` ni `tests/`. El módulo reutiliza la visión, la justificación y los requisitos no funcionales generales de la Inception original; aquí solo se agrega lo que es propio de Análisis Estructural.

### Nombre del módulo

Análisis Estructural (técnica MICMAC de Godet, adaptada por el profesor del curso).

### Alcance v1

Este módulo es una sola herramienta con cuatro etapas internas, no un conjunto de matrices independientes entre sí. Todo lo que sigue se apoya en la revisión directa que hizo el juez de los archivos originales del profesor (ver "Fuentes consultadas").

Etapas de la herramienta:

1. **Carga de variables.** El estudiante define las variables de su tema de análisis. La guía del profesor admite hasta 200; el uso real de un curso probablemente es mucho menor [VERIFICAR, punto 2 de la lista de abajo].
2. **Matriz de influencias directas.** Una matriz de N filas por N columnas, donde el estudiante califica de 0 a 4 cuánto influye cada variable (fila) sobre cada otra (columna): 0 no influye, 1 débil, 2 moderada, 3 fuerte, 4 muy fuerte. La diagonal queda bloqueada: una variable no se califica a sí misma.
3. **Motricidad, dependencia y plano estratégico.** El sistema calcula, para cada variable, la motricidad (suma de su fila: cuánto influye sobre las demás) y la dependencia (suma de su columna: cuánto la influyen las demás), y la ubica como un punto en un plano con la motricidad en el eje Y y la dependencia en el eje X.
4. **Selección para la siguiente etapa.** Junto al gráfico hay una tabla de ranking donde el estudiante marca con SÍ o NO cuáles variables pasan a la siguiente etapa de su trabajo. Una hoja final de solo lectura muestra la lista de las variables marcadas con SÍ.

Lo que hace el gráfico en v1:

- **Cuadrantes.** Las dos líneas que dividen el plano NO van en el promedio de los valores cargados. Van exactamente en la mitad del máximo de motricidad observado entre todas las variables (eje Y) y en la mitad del máximo de dependencia observado (eje X). Esto se verificó leyendo el código VBA del Excel original; no es una suposición de la teoría MICMAC de manual.
- **Rótulos oficiales de los cuatro cuadrantes.** Los que rotula el gráfico real del Excel, el que ve el estudiante: INDEPENDIENTES (motricidad alta, dependencia baja), AMBIGUAS (motricidad alta, dependencia alta), AUTONOMAS (motricidad baja, dependencia baja) y DEPENDIENTES (motricidad baja, dependencia alta). Son la terminología estándar de MICMAC y son los rótulos del sistema.
- **Descripción complementaria.** El PDF "Qué carga usted y qué sale solo" usa lenguaje llano para los mismos cuatro cuadrantes, en el mismo orden: "Las que hay que mover", "Inestables", "Se dejan para después" y "Resultados". Se usan solo como texto de apoyo si hace falta, no como rótulos del gráfico.
- **Proyección sobre la diagonal.** Para cada variable, el gráfico dibuja una línea desde su punto (dependencia, motricidad) hasta su proyección sobre la diagonal de igualdad, el punto donde motricidad y dependencia serían iguales: ((dependencia + motricidad) / 2, (dependencia + motricidad) / 2). Ningún PDF lo menciona, pero el código del Excel lo construye siempre y lo muestra, así que forma parte del alcance v1.

Lo que v1 no incluye de este módulo: la "siguiente etapa" del trabajo del estudiante a la que pasan las variables marcadas con SÍ. El módulo solo entrega la lista de las marcadas.

### Historia de usuario

Como estudiante, quiero cargar las variables de mi tema de análisis y calificar en una matriz cuánto influye cada una sobre las demás, para que el sistema calcule automáticamente su motricidad y dependencia, las ubique en el plano estratégico según esos valores, y me permita marcar cuáles paso a la siguiente etapa de mi trabajo.

Es una sola historia porque el módulo es una sola herramienta con cuatro etapas encadenadas: cada etapa toma como entrada lo que produjo la anterior, y ninguna se usa por separado.

### Requisitos no funcionales específicos del módulo

Ya cubiertos por los requisitos no funcionales generales de la Inception original y por lo tanto no se repiten: cálculo íntegro en el navegador, apertura con doble clic sin instalación, interfaz usable por un estudiante sin conocimientos técnicos, y guardado automático del progreso en localStorage. Los que no están cubiertos:

- **Fidelidad a la regla del profesor, no a la bibliografía.** El requisito general pide coincidir con las fórmulas clásicas verificables contra un caso de bibliografía conocido. En este módulo la referencia autoritativa es el código VBA del Excel del profesor, porque adapta la técnica: las líneas de los cuadrantes van en la mitad del máximo observado, no en el promedio. Donde la teoría de manual y el Excel del profesor difieran, el sistema sigue al Excel.
- **Tamaño de la matriz.** La matriz y el gráfico deben seguir siendo utilizables, dentro de una sola página sin backend, con el máximo que admite la guía del profesor: 200 variables, es decir una matriz de 200 por 200 con la diagonal bloqueada. No se fija todavía un tiempo de respuesta objetivo [VERIFICAR, punto 2 de la lista de abajo].
- **Diagonal no editable.** La diagonal debe quedar bloqueada de forma que el estudiante no pueda calificarla, tal como el Excel original.

### Riesgos de desarrollo y puntos [VERIFICAR] de este módulo

Riesgos:

- La base de datos original del profesor (`data-ae.accdb`) está protegida con contraseña y no se pudo abrir. No hay un caso ya resuelto contra el cual validar los cálculos de este módulo. Igual que en el módulo 1, hay que construir casos de prueba propios a partir de las reglas ya verificadas en el código VBA (motricidad como suma de fila, dependencia como suma de columna, líneas en la mitad del máximo, proyección sobre la diagonal).
- La lógica de la tabla de ranking vive en una consulta guardada dentro de esa misma base protegida, así que no pudo inspeccionarse (ver punto 1 de la lista siguiente).

Puntos [VERIFICAR] para el profesor:

1. **Criterio de orden de la tabla de "ranking estratégico".** No se pudo determinar qué ordena esa tabla: si es simplemente el orden de carga (V1, V2, V3...) o algún criterio de importancia. El cálculo real vive en la consulta guardada `c03_Motricidad_Dependencia` de la base Access protegida con contraseña, que no se pudo abrir.
2. **Cantidad máxima de variables para el curso.** La guía del profesor dice hasta 200, pero el uso real de un curso probablemente es mucho menor. Falta confirmar el máximo que debe soportar v1, porque de eso depende cuánto pesa el requisito de tamaño de la matriz.

### Fuentes consultadas

Archivos originales del profesor, en `C:\Obsidian\Mi_Segundo_Cerebro\07_gestion_estrategica_software\Modelos Estrategicos\`, revisados directamente por el juez (la sesión de Cowork del Project "Prompt Engineering") y usados aquí como hallazgos verificados:

- `Soft Analisis Estructural.xlsm`: el Excel con macros. De su código VBA salen las reglas de los cuadrantes (mitad del máximo observado), los rótulos del gráfico y la línea de proyección sobre la diagonal.
- `Que carga usted y que sale solo.pdf`: qué carga el estudiante y qué sale solo; fuente de los nombres en lenguaje llano de los cuadrantes.
- `Guia de uso - Cargar los Modelos Estrategicos.pdf`: fuente del máximo de 200 variables.
- `Manual tecnico - Modelos Estrategicos.pdf`.
- `data-ae.accdb`: base Access protegida con contraseña, que no se pudo abrir. Contiene la consulta `c03_Motricidad_Dependencia`.

Este Inception se escribió a partir de esos hallazgos, tal como los entregó el juez. El constructor no volvió a abrir los archivos originales en esta fase.

## Elaboration I — Módulo 2: Análisis Estructural

Objetivo de la fase: definir, para la historia de usuario del módulo, pruebas de aceptación concretas en formato Given-When-Then (Dado / Cuando / Entonces), con valores numéricos y resultado esperado exacto. No hay código en esta fase. Los resultados esperados de la prueba AE.1 los verificó el juez aritméticamente y se transcriben tal cual; los de las demás pruebas se calcularon a mano a partir de las fórmulas de abajo y se revisaron fila por fila y columna por columna.

Convenciones de esta sección:

- La numeración es "Prueba AE.n" (AE por Análisis Estructural) y no "Prueba X.Y", porque el módulo tiene una sola historia y las pruebas 1.1 a 7.1 ya pertenecen al Módulo 1.
- Notación de las matrices: la fila influye sobre la columna, la calificación va de 0 a 4 y la diagonal (marcada con "-") está bloqueada. Las variables se llaman V1, V2, V3... en el orden en que se cargan.
- Los valores numéricos se comparan con dos decimales. Las proyecciones se citan con cuatro decimales solo para que se puedan verificar a mano.
- Las marcas [VERIFICAR] señalan una convención que debe confirmarse con el profesor o con el juez. Se consolidan al final de esta sección.

Fórmulas verificadas contra el código VBA original en Inception, que estas pruebas no rederivan:

- motricidad(Vi) = suma de la fila i de la matriz.
- dependencia(Vi) = suma de la columna i de la matriz.
- corteY = (máximo de motricidad entre todas las variables) / 2.
- corteX = (máximo de dependencia entre todas las variables) / 2.
- Cuadrante: motricidad >= corteY y dependencia < corteX es INDEPENDIENTES. Motricidad >= corteY y dependencia >= corteX es AMBIGUAS. Motricidad < corteY y dependencia < corteX es AUTONOMAS. Motricidad < corteY y dependencia >= corteX es DEPENDIENTES. Un valor exactamente en el corte cuenta como alta en ese eje.
- Proyección (x) = (motricidad − dependencia) / 2.
- Proyección (y) = |dependencia − motricidad| / √2.

### Historia: Análisis Estructural

**Prueba AE.1: cuatro variables que cubren los cuatro cuadrantes**

- Dado que el estudiante carga cuatro variables y califica la matriz de influencias así:

  | | V1 | V2 | V3 | V4 |
  |---|---|---|---|---|
  | V1 | - | 4 | 0 | 4 |
  | V2 | 1 | - | 2 | 4 |
  | V3 | 0 | 1 | - | 0 |
  | V4 | 0 | 2 | 0 | - |

- Cuando el sistema calcula el análisis
- Entonces se cumple todo lo siguiente:
  - Los cortes son corteY = 4 y corteX = 4 (el máximo de motricidad es 8, la de V1, y el máximo de dependencia es 8, la de V4).
  - Resultado por variable:

    | Variable | Motricidad | Dependencia | Cuadrante | Proyección (x, y) |
    |---|---|---|---|---|
    | V1 | 8 | 1 | INDEPENDIENTES | (3.5, 4.9497) |
    | V2 | 7 | 7 | AMBIGUAS | (0, 0) |
    | V3 | 1 | 2 | AUTONOMAS | (−0.5, 0.7071) |
    | V4 | 2 | 8 | DEPENDIENTES | (−3, 4.2426) |

  - Cada cuadrante tiene exactamente una variable, así que esta prueba ejercita los cuatro rótulos a la vez.
  - V2 tiene motricidad igual a dependencia, por lo que cae sobre la diagonal de igualdad y su proyección es (0, 0).

**Prueba AE.2: motricidad exactamente en el corte [VERIFICAR]**

- Dado que el estudiante califica tres variables así:

  | | V1 | V2 | V3 |
  |---|---|---|---|
  | V1 | - | 2 | 2 |
  | V2 | 1 | - | 1 |
  | V3 | 0 | 0 | - |

- Cuando el sistema calcula el análisis
- Entonces las motricidades son V1 = 4, V2 = 2 y V3 = 0, y las dependencias son V1 = 1, V2 = 2 y V3 = 3. Por tanto corteY = 4 / 2 = 2 y corteX = 3 / 2 = 1.5. V2 tiene motricidad 2, exactamente igual a corteY, y cuenta como alta. El resultado es:
  - V1: INDEPENDIENTES.
  - V2: AMBIGUAS.
  - V3: DEPENDIENTES.

El juez fijó que un valor exactamente en el corte cuenta como "alta" en ese eje, la misma lógica que el umbral del BCG en el Módulo 1: un valor justo en el corte cuenta como la clasificación más notable. Queda pendiente de confirmar con el profesor, igual que el umbral del BCG.

**Prueba AE.3: dependencia exactamente en el corte [VERIFICAR]**

- Dado que el estudiante califica tres variables así (es la matriz de AE.2 transpuesta):

  | | V1 | V2 | V3 |
  |---|---|---|---|
  | V1 | - | 1 | 0 |
  | V2 | 2 | - | 0 |
  | V3 | 2 | 1 | - |

- Cuando el sistema calcula el análisis
- Entonces las motricidades son V1 = 1, V2 = 2 y V3 = 3, y las dependencias son V1 = 4, V2 = 2 y V3 = 0. Por tanto corteY = 3 / 2 = 1.5 y corteX = 4 / 2 = 2. V2 tiene dependencia 2, exactamente igual a corteX, y cuenta como alta. El resultado es:
  - V1: DEPENDIENTES.
  - V2: AMBIGUAS.
  - V3: INDEPENDIENTES.

Aplica la misma convención de valor en el corte que AE.2.

**Prueba AE.4: los cortes están en la mitad del máximo, no en el promedio**

- Dado que el estudiante califica cuatro variables así:

  | | V1 | V2 | V3 | V4 |
  |---|---|---|---|---|
  | V1 | - | 4 | 4 | 4 |
  | V2 | 3 | - | 2 | 2 |
  | V3 | 3 | 2 | - | 2 |
  | V4 | 3 | 2 | 2 | - |

- Cuando el sistema calcula el análisis
- Entonces las motricidades son V1 = 12, V2 = 7, V3 = 7 y V4 = 7, y las dependencias son V1 = 9, V2 = 8, V3 = 8 y V4 = 8. Los cortes son corteY = 12 / 2 = 6 y corteX = 9 / 2 = 4.5, y las cuatro variables quedan en AMBIGUAS. (Si los cortes se pusieran en el promedio, que es 33 / 4 = 8.25 en los dos ejes, V2, V3 y V4 caerían en AUTONOMAS y solo V1 en AMBIGUAS. Esta prueba detecta ese error.)

**Prueba AE.5: la diagonal está bloqueada**

- Dado que el estudiante tiene cargadas tres variables y abre la matriz de influencias
- Cuando intenta escribir una calificación en una celda de la diagonal (V1 con V1, V2 con V2 o V3 con V3)
- Entonces el sistema no le permite escribir en esas celdas. Además, las celdas de la diagonal no cuentan como celdas sin calificar: si el estudiante completa las seis celdas fuera de la diagonal, la matriz se considera completa y el sistema calcula.

**Prueba AE.6: celdas sin calificar y valores fuera de rango**

Misma regla que el resto del sistema (la función `evaluarMatriz` ya existente), aplicada a tres variables con las seis celdas fuera de la diagonal.

- Dado que el estudiante tiene cargadas tres variables
- Cuando se presentan estos tres casos por separado
- Entonces el sistema responde así en cada uno:
  - Ninguna de las seis celdas tiene valor: estado "vacía". No se muestra ningún error, no se calcula nada y no hay gráfico.
  - Cinco celdas tienen valor y una está en blanco: estado "inválida". Se muestra el error de campos vacíos del sistema (la redacción exacta se define en el diseño), no se calcula nada y no hay gráfico.
  - Las seis celdas tienen valor, pero una es 5 (o −1): estado "inválida". Se muestra un error que indica que la calificación debe estar entre 0 y 4, no se calcula nada y no hay gráfico.
- Los datos escritos se conservan en los tres casos: el estudiante no pierde lo que ya escribió.

**Prueba AE.7: mínimo de dos variables**

- Dado que el estudiante carga una sola variable, V1
- Cuando abre la matriz de influencias
- Entonces no hay ninguna celda que calificar (la única celda es la diagonal, bloqueada), y el sistema trata la situación igual que una matriz vacía: estado "vacía", sin error, sin cálculo y sin gráfico.

- Dado que el estudiante carga dos variables y califica V1 sobre V2 = 3 y V2 sobre V1 = 1
- Cuando el sistema calcula el análisis
- Entonces las motricidades son V1 = 3 y V2 = 1, las dependencias son V1 = 1 y V2 = 3, y corteY = corteX = 3 / 2 = 1.5. V1 queda en INDEPENDIENTES con proyección (1, 1.4142) y V2 en DEPENDIENTES con proyección (−1, 1.4142).

**Prueba AE.8: tabla de ranking y hoja Validadas en orden de carga [VERIFICAR]**

- Dado el análisis de la prueba AE.1 ya calculado
- Cuando el estudiante marca V4 con SÍ, V2 con NO, V3 con NO y, por último, V1 con SÍ
- Entonces se cumple todo lo siguiente:
  - La tabla de ranking lista las cuatro variables en orden de carga: V1, V2, V3, V4.
  - La hoja Validadas (la hoja final de solo lectura) lista exactamente V1 y V4, en ese orden, aunque V4 se haya marcado antes que V1.
  - La hoja Validadas no permite editar ninguna marca.

El orden de carga es un criterio provisional: el criterio real de la tabla de ranking no se pudo verificar (punto 1 de los [VERIFICAR] de la Inception de este módulo). Si el profesor confirma otro criterio, esta prueba debe reescribirse.

**Prueba AE.9: ninguna variable marcada con SÍ**

- Dado el análisis de la prueba AE.1 ya calculado
- Cuando el estudiante no marca ninguna variable con SÍ (todas en NO o sin marcar)
- Entonces la hoja Validadas queda sin ninguna variable en su lista, sin error y sin bloquear el resto del módulo.

**Prueba AE.10: los datos siguen tras recargar**

- Dado el análisis de la prueba AE.1 con las marcas de la prueba AE.8
- Cuando el estudiante recarga la página o vuelve a abrir el archivo
- Entonces las cuatro variables, las calificaciones de la matriz y las marcas SÍ y NO siguen ahí, el sistema recalcula los mismos resultados de la prueba AE.1 y la hoja Validadas vuelve a listar V1 y V4. Esto aplica al módulo el requisito general de guardado automático de la Inception original.

### Resumen de la fase

| Prueba | Qué verifica | Estado |
|---|---|---|
| AE.1 | Motricidad, dependencia, cortes, los cuatro cuadrantes y las proyecciones | Con prueba concreta |
| AE.2 | Motricidad exactamente en corteY | Con prueba concreta, convención del juez, confirmación del profesor pendiente [VERIFICAR] |
| AE.3 | Dependencia exactamente en corteX | Con prueba concreta, convención del juez, confirmación del profesor pendiente [VERIFICAR] |
| AE.4 | Cortes en la mitad del máximo, no en el promedio | Con prueba concreta |
| AE.5 | Diagonal bloqueada | Con prueba concreta |
| AE.6 | Celdas sin calificar (vacía e inválida) y valores fuera de rango | Con prueba concreta |
| AE.7 | Mínimo de dos variables | Con prueba concreta |
| AE.8 | Tabla de ranking y hoja Validadas en orden de carga | Con prueba concreta, orden provisional [VERIFICAR] |
| AE.9 | Ninguna variable marcada con SÍ | Con prueba concreta |
| AE.10 | Persistencia al recargar | Con prueba concreta |

Total: 10 pruebas Given-When-Then.

Puntos [VERIFICAR] de esta fase:

1. **Valor exactamente en el corte, resuelto por el juez.** Cuenta como alta en ese eje (fórmulas corregidas arriba: motricidad >= corteY, dependencia >= corteX). Es la misma lógica que el umbral del BCG en el Módulo 1. Queda pendiente de confirmar con el profesor, igual que ese umbral.
2. **Significado de la proyección, aclarado por el juez.** No hay contradicción: son dos cosas distintas del mismo código VBA. El punto de proyección sobre la diagonal, ((dependencia + motricidad) / 2, mismo valor en los dos ejes), es el punto al que el gráfico dibuja una línea desde cada variable; para V1 de AE.1 es (4.5, 4.5), y no se guarda como dato, solo se dibuja. Las fórmulas verificadas de proyección (x) e (y) son dos valores distintos, derivados de ese mismo punto, que sí se guardan en la tabla de ranking del Excel original: (x) es el desplazamiento con signo entre la variable y el punto de proyección; (y) es la distancia real, siempre positiva, entre la variable y ese mismo punto. Las pruebas de esta fase usan correctamente (x) e (y); no hace falta cambiar ninguna.
3. **Criterio de orden de la tabla de ranking y de la hoja Validadas.** Provisional: orden de carga (V1, V2, V3...). El criterio real vive en una consulta de la base Access protegida y no se pudo verificar (punto 1 de la Inception de este módulo).
4. **Calificaciones con decimales.** La escala es de 0 a 4 con cinco niveles con nombre. Ninguna prueba fija si un valor como 2.5 se rechaza. Se asume que sí, igual que las clasificaciones de EFI, EFE y MPC en el Módulo 1, pero no se fija con una prueba numérica.
5. **Redacción y estado inicial de las marcas.** Las pruebas AE.6 y AE.9 no fijan la redacción exacta de los mensajes de error ni de una lista vacía, que se definen en el diseño. Tampoco se fija si una variable recién cargada aparece como NO o sin marcar; las pruebas AE.8 y AE.9 no dependen de eso.

## Elaboration II — Módulo 2: Análisis Estructural

Objetivo de la fase: definir qué suma Análisis Estructural a la arquitectura de seis componentes ya fijada en Elaboration II del Módulo 1, con escenarios y diagramas, sin escribir código. Las fórmulas, clasificaciones y reglas de validación ya fijadas en Elaboration I de este módulo (pruebas AE.1 a AE.10) no se reinterpretan aquí; los componentes solo las aplican. Lo que esta fase no resuelve se marca [VERIFICAR] y se consolida al final.

Principios de la arquitectura: se mantienen los del Módulo 1 sin cambios. La Vista es el único orquestador, el Validador y el Motor de Cálculo no dependen del navegador, los datos ingresados se guardan en cada cambio sean válidos o no, y los resultados solo se calculan y muestran cuando la validación pasa. Lo que Análisis Estructural suma a la fuente de verdad: lo que el usuario ingresó son las variables, las calificaciones de la matriz y las marcas SÍ y NO. La motricidad, la dependencia, los cortes, los cuadrantes y las proyecciones se recalculan cuando hacen falta y no se guardan. Las marcas SÍ y NO sí se guardan, porque son datos del estudiante y no se pueden derivar de nada.

### 1. Componentes del framework

No se repite la tabla de los seis componentes del Módulo 1. Solo se agrega lo que Análisis Estructural suma a cada uno.

| Componente | Qué suma Análisis Estructural |
|---|---|
| Vista | Ningún método nuevo. Reutiliza `renderFormulario`, `renderResultados` y `despacharEvento` con un octavo valor de matriz, `"AE"`, para la matriz NxN de influencias, la tabla de ranking con sus marcas SÍ y NO y la hoja Validadas (solo lectura). Pregunta abierta para Construction I: si la hoja Validadas necesita su propia función de renderizado o le alcanza `renderResultados` [VERIFICAR]. |
| Validador | Probablemente reutiliza `validarCamposVacios` y `validarRango` aplicados a las celdas fuera de la diagonal (rango de 0 a 4, y mínimo de dos variables), sin método nuevo. Pregunta abierta para Construction I, no se decide aquí [VERIFICAR]. |
| Motor de Cálculo | Método nuevo `calcularAE(variables, matriz)`. Aplica exactamente las fórmulas ya fijadas y probadas en Elaboration I: motricidad (suma de la fila), dependencia (suma de la columna), corteY y corteX (mitad del máximo observado), cuadrante (un valor exactamente en el corte cuenta como alta) y proyección (x) e (y). Devuelve un `ResultadoAE` con esos valores por variable y los dos cortes. |
| Motor de Gráficos | Método nuevo `dibujarAE(resultado)`. Dibuja el plano con la dependencia en el eje X y la motricidad en el eje Y, las dos líneas de corte, los cuatro cuadrantes rotulados (INDEPENDIENTES, AMBIGUAS, AUTONOMAS y DEPENDIENTES), el punto de cada variable y su línea hasta el punto de proyección sobre la diagonal de igualdad. |
| Persistencia | Sin cambios: se reutiliza tal cual. Guarda las variables, la matriz y las marcas como parte del estado de la sesión. |
| Exportador | Sin cambios: se reutiliza tal cual. Qué contiene el archivo de este módulo queda abierto en el escenario AE-3. |

Consecuencia que debe recogerse en Construction I de este módulo: la decisión de Construction I del Módulo 1 dice que `matriz` es siempre uno de siete strings (`"BCG"`, `"EFI"`, `"EFE"`, `"MPC"`, `"PEYEA"`, `"MIE"`, `"GE"`). Con `"AE"` pasan a ser ocho, por lo que esa frase quedará desactualizada y pedirá un commit de sincronización en esa fase, igual que en fases anteriores. Tampoco se decide aquí cómo se acomodan los datos de este módulo en el estado guardado ni si el formato guardado cambia de versión [VERIFICAR].

Relación con las pruebas de Elaboration I de este módulo. Se usa una tabla aparte de la del Módulo 1, porque aquí las pruebas se cruzan contra los dos métodos nuevos y no contra componentes completos:

| Prueba | `calcularAE` | `dibujarAE` | Otros componentes que la ejercitan |
|---|---|---|---|
| AE.1 | Motricidad, dependencia, cortes, cuadrantes y proyecciones de las cuatro variables | Los cuatro cuadrantes rotulados, cuatro puntos y cuatro líneas de proyección | Vista |
| AE.2 | Cuadrante de V2 con motricidad exactamente en corteY | Posición del punto de V2 en AMBIGUAS | |
| AE.3 | Cuadrante de V2 con dependencia exactamente en corteX | Posición del punto de V2 en AMBIGUAS | |
| AE.4 | Cortes en la mitad del máximo y no en el promedio | Posición de las líneas de corte en el plano | |
| AE.5 | No se invoca | No se invoca | Vista (las celdas de la diagonal no se pueden escribir) y Validador (la diagonal no cuenta como celda sin calificar) |
| AE.6 | No se invoca | No se invoca | Validador (matriz incompleta y valor fuera de rango) y Vista (muestra el error o no muestra nada) |
| AE.7 | Cálculo con dos variables | Gráfico con dos variables | Validador y Vista (una sola variable se trata como vacía) |
| AE.8 | No interviene | No interviene | Vista (tabla de ranking y hoja Validadas en orden de carga) y Persistencia (guarda las marcas) |
| AE.9 | No interviene | No interviene | Vista (hoja Validadas sin variables) |
| AE.10 | Recalcula los mismos resultados de AE.1 | Vuelve a dibujar el mismo gráfico | Persistencia, Validador y Vista |

### 2. Escenarios (plays)

**Escenario AE-1: flujo principal, de la carga a la selección**

El estudiante abre el módulo de Análisis Estructural y carga las cuatro variables V1 a V4. Con cada cambio, la Vista pide a la Persistencia que guarde el estado. Cuando completa las doce celdas fuera de la diagonal con las calificaciones de la prueba AE.1, la Vista pasa los datos al Validador, que confirma que todas las celdas tienen un valor de 0 a 4 y que hay al menos dos variables. La Vista pasa las variables y la matriz al Motor de Cálculo, que devuelve para cada variable su motricidad y dependencia (8 y 1, 7 y 7, 1 y 2, 2 y 8), los cortes (4 y 4), el cuadrante (INDEPENDIENTES, AMBIGUAS, AUTONOMAS, DEPENDIENTES) y la proyección. La Vista entrega ese resultado al Motor de Gráficos, que dibuja el plano con los cuatro cuadrantes, los cuatro puntos y sus líneas de proyección, y muestra la tabla de ranking en orden de carga. Después el estudiante marca V1 y V4 con SÍ y V2 y V3 con NO, como en la prueba AE.8. Cada marca la guarda la Persistencia y la Vista actualiza la hoja Validadas, que lista V1 y V4. Las marcas no cambian ningún resultado calculado, así que en ese paso no se invocan el Validador ni los motores.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.

**Escenario AE-2: rechazo por matriz incompleta**

El estudiante tiene tres variables cargadas y completa solo cinco de las seis celdas fuera de la diagonal, como en la prueba AE.6. La Vista guarda el estado en la Persistencia y pasa los datos al Validador, que detecta una celda sin calificar y devuelve el error. La Vista muestra el error y no muestra resultados ni gráfico. El Motor de Cálculo y el Motor de Gráficos no se invocan. Con ninguna celda escrita, o con una sola variable, la Vista sigue el mismo camino pero no muestra ningún error, porque la matriz está en estado "vacía". Con un valor fuera de 0 a 4 (por ejemplo 5), el Validador devuelve un error de rango y el flujo se corta igual.
Componentes: Vista, Persistencia, Validador.
Este escenario cubre las pruebas AE.6 y la parte de AE.7 con una sola variable. Tiene el mismo patrón que el escenario 3 del Módulo 1: el flujo se corta en la validación.

**Escenario AE-3: exportación a Excel**

Con el análisis de AE.1 ya calculado y válido, el usuario pulsa "Exportar". La Vista pasa los datos al Validador, que confirma que sean válidos. La Vista pide el resultado al Motor de Cálculo y entrega al Exportador el módulo, los datos ingresados y el resultado. El Exportador genera el archivo .xlsx con la librería embebida y la Vista dispara la descarga, todo sin conexión a internet. Si los datos no son válidos, la Vista muestra los errores y no llama al Exportador, como en el escenario 4 del Módulo 1.
Componentes: Vista, Validador, Motor de Cálculo, Exportador.
[VERIFICAR] qué contiene la hoja de datos de una matriz NxN: la matriz de influencias cruda (N filas por N columnas, con la diagonal vacía), los resultados por variable (motricidad, dependencia, cuadrante y proyección), o las dos cosas en hojas separadas. Falta decidir también si el archivo incluye las marcas SÍ y NO y la lista de Validadas. Como referencia, el Módulo 1 exporta dos hojas, "Datos" y "Resultados", con solo valores y sin imagen del gráfico, de forma provisional.

**Escenario AE-4: recuperación tras recargar**

El estudiante recarga la página o vuelve a abrir el archivo, con el análisis de AE.1 y las marcas de AE.8 ya guardados. La Vista pide a la Persistencia el estado guardado. Si existe, la Vista vuelve a llenar las variables, la matriz y las marcas, y el Validador revisa los datos. Como la matriz está completa y es válida, el Motor de Cálculo recalcula los mismos resultados de AE.1 y el Motor de Gráficos vuelve a dibujar el plano; la hoja Validadas vuelve a listar V1 y V4. Si la matriz estuviera incompleta, se muestran sus datos sin resultado. Si el almacenamiento está vacío, corrupto o inaccesible, la Persistencia devuelve un estado vacío y la Vista muestra el módulo en blanco.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
Este escenario cubre la prueba AE.10. Aplica a este módulo los mismos puntos abiertos del escenario 5 del Módulo 1 (mensajes de error de una matriz incompleta tras la recarga y aviso al descartar un estado corrupto), sin agregar otros.

Cobertura de componentes por escenario:

| Componente | AE-1 | AE-2 | AE-3 | AE-4 |
|---|---|---|---|---|
| Vista | sí | sí | sí | sí |
| Validador | sí | sí | sí | sí |
| Motor de Cálculo | sí | no | sí | sí |
| Motor de Gráficos | sí | no | no | sí |
| Persistencia | sí | sí | no | sí |
| Exportador | no | no | sí | no |

### 3. Diagramas

**Diagrama de clases**

No se crea un diagrama nuevo. Se editó el `classDiagram` de la Elaboration II del Módulo 1 (arriba, en este documento) para agregar `+calcularAE(variables, matriz) ResultadoAE` a `MotorCalculo` y `+dibujarAE(resultado)` a `MotorGraficos`. Las demás clases y relaciones no cambian.

**Diagrama de secuencia del escenario AE-1: flujo principal**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: carga las variables y califica la matriz de influencias
    V->>P: guardar(estado)
    V->>VA: validar(AE, datos)
    VA-->>V: datos válidos
    V->>MC: calcularAE(variables, matriz)
    MC-->>V: motricidad, dependencia, cortes, cuadrantes y proyecciones
    V->>MG: dibujarAE(resultado)
    MG-->>V: plano con cuadrantes, puntos y líneas de proyección
    V-->>U: muestra resultados, gráfico y tabla de ranking
    U->>V: marca SÍ o NO en cada variable
    V->>P: guardar(estado)
    V-->>U: actualiza la hoja Validadas
```

**Diagrama de secuencia del escenario AE-2: rechazo por matriz incompleta**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    U->>V: deja una celda de la matriz sin calificar
    V->>P: guardar(estado)
    V->>VA: validar(AE, datos)
    VA-->>V: error, hay celdas sin calificar
    V-->>U: muestra el error sin resultados ni gráfico
```

Los escenarios AE-3 y AE-4 no tienen diagrama propio: siguen la misma secuencia que los escenarios 4 y 5 del Módulo 1, con el valor de matriz `"AE"` y las llamadas `calcularAE` y `dibujarAE` en lugar de las de EFI.

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Hoja Validadas: si necesita su propia función de renderizado en la Vista o le alcanza `renderResultados`. Se decide en Construction I.
2. Validador: si reutiliza `validarCamposVacios` y `validarRango` sobre las celdas fuera de la diagonal o necesita un método nuevo. Se decide en Construction I.
3. Estructura del Excel exportado del módulo: matriz cruda, resultados por variable o ambas en hojas separadas, y si incluye las marcas SÍ y NO y la lista de Validadas.
4. Estado guardado: cómo se acomodan variables, matriz y marcas, y si el formato guardado cambia de versión. Se decide en Construction I, junto con la sincronización de la frase de las siete siglas.
5. Los puntos abiertos de recuperación de sesión del Módulo 1 (mensajes de error de una matriz incompleta y aviso al descartar un estado corrupto) aplican también a este módulo.
6. Siguen abiertos los puntos de Elaboration I de este módulo: el orden provisional del ranking, las calificaciones con decimales y la redacción de mensajes.

## Construction I — Módulo 2: Análisis Estructural

Objetivo de la fase: resolver las tres preguntas que Elaboration II del Módulo 1 y de este módulo dejaron abiertas para Construction I, y dejar el esqueleto del proyecto extendido con los stubs de Análisis Estructural, sin comportamiento real. Los cuerpos de los métodos están vacíos o llevan un comentario de marcador de posición. Ninguna lógica de cálculo, validación, graficado ni renderizado se implementa en esta fase, y esta fase no toca `index.html` ni `tests/`: el esqueleto sigue viviendo dentro de este documento.

### 1. Decisiones de arquitectura

El stack no cambia: archivo único `index.html`, HTML, CSS y JavaScript sin framework ni backend, SVG para los gráficos, SheetJS embebido y localStorage. Se aplica tal cual a este módulo.

Las tres preguntas abiertas se resuelven con decisiones del juez:

| Pregunta abierta en Elaboration II | Decisión |
|---|---|
| ¿La hoja Validadas necesita su propia función de renderizado en la Vista? | No. Ningún método público nuevo en la Vista. `renderResultados('AE', resultado)` cubre también la hoja Validadas, con una auxiliar privada, `resultadosAE`, que construye tanto la tabla de ranking como la lista filtrada de solo lectura. |
| ¿El Validador necesita un método nuevo para esta matriz? | No. Ningún método público nuevo. `validar('AE', datos)` reutiliza `validarCamposVacios` y `validarRango` sobre las celdas fuera de la diagonal, con una auxiliar privada, `erroresAE`, equivalente a `erroresBCG` para este módulo. `erroresAE` aplica el mínimo de dos variables y excluye la diagonal del conteo de campos vacíos. |
| ¿Cómo se acomodan los datos del módulo en el estado guardado, y cambia la versión del formato? | Es una clave nueva, `datos.AE`, junto a `BCG`, `EFI` y las demás. No cambia la versión del formato guardado. `Persistencia.cargar()` debe devolver `AE` vacío si un estado guardado anterior no la tiene. Su forma está justo debajo. |

Forma de `estado.datos.AE`:

```text
{
  variables: [nombre1, nombre2, ...],
  matriz:    [[null, v12, v13, ...], [v21, null, v23, ...], ...],
  marcas:    [marca1, marca2, ...]
}
```

- `matriz[i][j]` es cuánto influye la variable i sobre la variable j. La diagonal es `null`, no 0: el 0 significa "no influye" y es una calificación hecha, mientras que `null` significa "no aplica".
- `marcas[i]` es `"SI"`, `"NO"` o `null` si no se marcó.
- Ejemplo con los datos de las pruebas AE.1 y AE.8:

```text
{
  variables: ["V1", "V2", "V3", "V4"],
  matriz: [[null, 4, 0, 4],
           [1, null, 2, 4],
           [0, 1, null, 0],
           [0, 2, 0, null]],
  marcas: ["SI", "NO", "NO", "SI"]
}
```

Reglas que acompañan a esa forma:

- **Campos del formulario.** Cada celda de la matriz usa el mismo patrón `data-campo` que las demás matrices, con una ruta como `matriz.0.1` (fila 0, columna 1). Las celdas de la diagonal no llevan campo editable, lo que cumple la prueba AE.5.
- **Agregar y quitar variables.** Quitar una variable sigue la regla ya establecida para BCG y las demás matrices: "− borra siempre el último". Se elimina la última variable, su fila y su columna de la matriz y su marca, nunca una del medio, lo que simplifica el ajuste de la matriz. Agregar una variable suma una fila y una columna al final, con la diagonal en `null`, y una marca `null`.
- **Estado guardado anterior.** Un estado guardado sin `datos.AE` no se descarta: se completa con `AE` vacío, de modo que los datos de las otras matrices sobreviven a la actualización. Además, `"AE"` pasa a ser un valor válido de matriz activa.

### 2. Esqueleto del proyecto

No hay un bloque de código nuevo. Se extendió el esqueleto de la Construction I del Módulo 1 (arriba, en este documento) agregando solo estos stubs, en el lugar de cada componente que corresponde y con la convención de nombres ya usada (`calcularBCG`, `erroresBCG`, `formularioBCG`, `resultadosBCG`). Cada cuerpo queda vacío con un marcador de posición.

| Dónde en el esqueleto | Stub nuevo | Tipo |
|---|---|---|
| `<main>`, después de la sección GE | `<section id="matriz-ae" class="matriz" data-matriz="AE" hidden>` con los cuatro contenedores `formulario`, `errores`, `resultados` y `grafico` | Contenedor HTML |
| Antes de `Validador` | `erroresAE(datos, errores)` | Auxiliar privada |
| `MotorCalculo` | `calcularAE(variables, matriz)` | Método público, ya definido en Elaboration II |
| `MotorGraficos` | `dibujarAE(resultado)` | Método público, ya definido en Elaboration II |
| `Persistencia.cargar` | Solo un comentario: devolver `AE` vacío si el estado guardado no lo trae | Comentario |
| Antes de `Vista` | `formularioAE(datos)` y `resultadosAE(resultado)` | Auxiliares privadas |

Lo que no cambia en el esqueleto: los seis objetos y sus firmas públicas, salvo las dos firmas nuevas de arriba. `Vista`, `Validador`, `Persistencia` y `Exportador` no suman ningún método público. Las auxiliares privadas son funciones sueltas del script, igual que `erroresBCG` y `formularioBCG`: no son métodos de los objetos y no modifican sus firmas.

Con estas extensiones, dos frases de "Correspondencia con Elaboration II" en la Construction I del Módulo 1 describían el esqueleto de ese módulo sin acotarlo: "El script no agrega métodos ni componentes nuevos" y "No hay métodos auxiliares". Un commit de sincronización aparte ya les agregó esa acotación, apuntando a esta sección.

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Celda fuera de la diagonal sin calificar: la decisión del juez fija `null` para la diagonal pero no para una celda por calificar. Se asume la convención de las demás matrices (texto vacío), de modo que la diagonal se distingue por su posición y no por su valor.
2. Cantidad de variables en un estado limpio: cuántas variables, con qué nombres y con cuántas marcas empieza el módulo. Las demás matrices arrancan con dos filas, y la prueba AE.7 admite una sola variable.
3. Rutas `data-campo` de los nombres de variable y de las marcas: se asume `variables.0` y `marcas.0` por analogía con `matriz.0.1`. La decisión del juez fija solo la de la matriz.

## Construction II — Módulo 2: Análisis Estructural

Objetivo de la fase: implementar en `index.html` la lógica y las pantallas de Análisis Estructural sobre el esqueleto de Construction I, con las fórmulas de Elaboration I ya corregidas. Es la primera fase de este módulo que escribe código. Lo que ya está en las fases anteriores no se repite aquí.

### 1. Qué se implementó

| Pieza | Estado |
|---|---|
| Sección `matriz-ae` y botón "AE" en la navegación | Implementada. El botón lleva el título "Análisis Estructural". El botón de exportar queda habilitado en AE. |
| `formularioAE` | Lista de variables (con "+ Agregar variable" y "− Quitar la última variable") y matriz NxN de influencias. Cada celda usa `data-campo="matriz.i.j"`; la diagonal no tiene campo y se ve bloqueada. Los rótulos de fila y columna son V1, V2, V3... |
| `erroresAE` y `validar('AE')` | Mínimo de dos variables, matriz cuadrada y calificaciones enteras de 0 a 4. La diagonal no cuenta como campo vacío: `validar` solo le pasa a `validarCamposVacios` las celdas fuera de la diagonal. |
| `calcularAE` | Aplica las fórmulas corregidas de Elaboration I (motricidad, dependencia, cortes, cuadrante con el corte como "alto", proyección (x) e (y) y punto de proyección). |
| `dibujarAE` | SVG con la misma técnica de BCG y PEYEA: plano con los cuatro cuadrantes rotulados, las dos líneas de corte, la diagonal de igualdad, un punto por variable rotulado V1, V2... (el nombre completo sale al pasar el mouse) y una línea de cada punto hasta su punto de proyección. |
| `resultadosAE` | Tabla de ranking en orden de carga con la columna SÍ o NO, y la hoja Validadas (lista de solo lectura de las marcadas con SÍ), ambas dentro de `renderResultados('AE', resultado)`. |
| Persistencia | `estadoVacio` trae `AE` vacío. `cargar()` completa `datos.AE` con vacío si un estado guardado anterior no lo trae o lo trae dañado, y conserva los demás datos. Sin cambio de versión del formato. |
| Exportador | Rama `"AE"` con las hojas "Datos" y "Resultados". Ver el punto 3. |
| Pruebas | `tests/elaboration1.test.js`: pruebas AE.1 a AE.10, más la comprobación X.6 del Exportador de AE. |

### 2. Forma final de `ResultadoAE`

La forma que devuelve `calcularAE` no cambió: `{ corteY, corteX, variables: [{ nombre, motricidad, dependencia, cuadrante, proyeccion: { x, y }, puntoProyeccion, marca }] }`, con las variables en el mismo orden de entrada. Lo que sí hubo que precisar es la entrada. `calcularAE(variables, matriz)` tiene dos parámetros y las marcas viven aparte en `estado.datos.AE.marcas`, así que el resultado no podía traer `marca` sin recibirla. Decisión de esta fase: cada elemento de `variables` puede ser un texto (el nombre) o un objeto `{ nombre, marca }`, igual que `valorDeFactor` ya acepta un número o un objeto en PEYEA. La Vista arma esos objetos a partir de `datos.variables` y `datos.marcas`. Una variable sin nombre se llama V1, V2, V3... según su posición.

### 3. [VERIFICAR] Estructura del Excel exportado

Decisión del juez: dos hojas, "Datos" y "Resultados", archivo `Mtx-AE.xlsx`. En "Datos", la matriz completa con los nombres de variable como encabezado de fila y de columna y la diagonal vacía. En "Resultados", una fila por variable en orden de carga (nombre, motricidad, dependencia, cuadrante, proyección x, proyección y y marca, que sale como SÍ, NO o vacío) y, después de una fila vacía, las filas "Corte Y" y "Corte X". No hay manera de confirmar si esto coincide con lo que produciría el Excel original: la base de datos del profesor está protegida con contraseña y ningún método disponible en este entorno logra abrirla. Queda pendiente de confirmar con el profesor.

### 4. Criterios aplicados donde las fases anteriores no fijaban un valor

- **Estado limpio:** `datos.AE` empieza con listas vacías, sin ninguna variable (cierra el punto 2 de los [VERIFICAR] de Construction I). El estudiante agrega las suyas con "+".
- **Celda sin calificar:** texto vacío, como en las demás matrices. La diagonal se distingue por su posición y vale `null` (punto 1 de Construction I).
- **Rutas de campos:** `variables.0` y `marcas.0` (punto 3 de Construction I). Una marca sin definir se guarda como `null` (el selector la manda como texto vacío y `despacharEvento` la convierte).
- **Matriz "vacía":** sin ninguna celda calificada fuera de la diagonal. Los nombres y las marcas no cuentan. Con una variable no hay celdas que calificar, así que cae en el mismo caso (prueba AE.7).
- **Calificaciones con decimales:** se rechazan, como en EFI, EFE y MPC. Sigue abierto el punto 4 de [VERIFICAR] de Elaboration I.
- **Muchos errores de rango:** el Validador lista los primeros 20 y resume el resto en una línea, para no escribir miles de mensajes en una matriz grande.
- **Aviso en pantalla:** debajo de la tabla de ranking hay un recuadro ámbar que dice que el orden es el de carga y que el criterio del curso está pendiente, igual que los demás criterios provisionales.
- **Escala del plano:** la misma en los dos ejes, para que la diagonal de igualdad quede a 45°, con un 10 % de margen sobre el máximo.
- **Todas las calificaciones en 0:** los cortes quedan en 0 y, como un valor en el corte cuenta como alto, todas las variables caen en AMBIGUAS. Es consecuencia de las fórmulas, no un criterio nuevo; no se agregó ningún aviso.

### 5. Verificación

Comando: `node tests/elaboration1.test.js`. Resultado de la última corrida: 32 de 32 comprobaciones correctas (las 21 anteriores, que incluyen la X.1 contra el diagrama de clases con `calcularAE` y `dibujarAE`, más AE.1 a AE.10 y X.6). Con cuatro fallos introducidos a propósito en una copia fuera del repositorio (corte en el promedio, corte con `>` en lugar de `>=`, variables ordenadas por motricidad y sin completar `AE` en estados anteriores), 8 comprobaciones fallaron.

Para poder probar el estado "vacía", "inválida" u "ok" con el `evaluarMatriz` real de la aplicación, `cargarApp` del arnés ahora devuelve además `evaluarMatriz` y un `fijarEstado`. No cambia ningún componente.

Además se recorrió el módulo en Chrome sin interfaz con eventos reales del DOM, sobre una copia temporal fuera del repositorio: agregar cuatro variables, completar la matriz de AE.1, ver los cuatro cuadrantes, las cuatro proyecciones y la lista Validadas con V1 y V4, quitar la última variable y la validación de un valor 5. Sin errores de consola, y BCG y GE siguen como estaban.

### 6. Qué queda para Construction III

- El resultado visual real de `dibujarAE` (posición de puntos, líneas de corte, rótulos y líneas de proyección), igual que `dibujarGE` solo se prueba por su error de pendiente.
- El bloqueo de la diagonal en pantalla (prueba AE.5) y el orden de renderizado de la tabla de ranking y de la hoja Validadas (AE.8 y AE.9).
- Rendimiento con una matriz grande (hasta 200 variables, 40 000 celdas): no se midió. Cada tecla recalcula y redibuja todo el módulo.
- Navegadores distintos de Chrome, el archivo descargado de internet y la exportación abierta en Excel: nada de esto se verificó en esta fase.

## Construction III — Módulo 2: Análisis Estructural

Objetivo de la fase: dejar un plan de pruebas manual para Análisis Estructural, pensado para que una persona lo siga con el navegador real, el mouse y el teclado. Esta fase no escribe ni modifica código de la aplicación. Cubre lo que Construction II dejó pendiente: el resultado visual de `dibujarAE`, el bloqueo de la diagonal en pantalla, el orden de la tabla de ranking y de la hoja Validadas, la experiencia con una matriz grande, Edge y Firefox, y la apertura del archivo exportado en un programa de hojas de cálculo. Lo que es propiedad del archivo completo (abre sin instalar, no pide red, funciona sin conexión) ya está en el bloque 5 de la Construction III del Módulo 1 y no se repite.

Estado del plan: ningún caso ha sido ejecutado por una persona, por eso todas las casillas "¿Pasó?" están en blanco. Los textos y números de "Resultado esperado" no se calcularon a mano: se obtuvieron recorriendo los mismos pasos sobre `index.html` en Chrome sin interfaz, con eventos reales del DOM (clics, escritura en los campos, cambio de los selectores, botón de exportar, recarga de la página y un script de consola), sobre una copia temporal fuera del repositorio. Los textos salen del DOM, la geometría del gráfico sale de los elementos SVG (posición de cada punto, línea y rótulo) y el contenido del archivo exportado sale del libro que genera SheetJS, leído antes de que el navegador lo descargue. Eso no reemplaza la ejecución real. Lo que la simulación no pudo reproducir queda marcado en cada caso: el aspecto del mouse y el retraso del texto emergente, la tecla Tab real, Edge y Firefox, la apertura en un programa de hojas de cálculo y la sensación de velocidad con la matriz grande.

### 0. Convenciones propias de Análisis Estructural

Valen las convenciones de la sección 0 de la Construction III del Módulo 1 (abrir la aplicación, cómo leer los resultados, el aviso ámbar y el recuadro rojo de errores, y la casilla "¿Pasó?"). Se agregan o se precisan estas:

- **Reinicio de datos (RD):** el mismo de siempre: F12, pestaña "Consola", escribir `localStorage.removeItem('mtx.estado'); location.reload()` y pulsar Enter. Después de un RD se abre en BCG.
- **Estado limpio de AE:** cero variables, no dos filas como las demás matrices. La sección AE muestra "Todavía no hay variables. Agregue al menos dos." y los botones "+ Agregar variable" y "− Quitar la última variable". Este último está deshabilitado mientras no haya variables.
- **Botones:** "+ Agregar variable" suma una variable al final, con su fila y su columna en la matriz. "− Quitar la última variable" borra siempre la última, nunca una del medio.
- **Rótulos:** las filas y columnas de la matriz y los puntos del gráfico se llaman V1, V2, V3... según el orden de carga. El nombre que escribe la persona se ve en la tabla de variables, en la tabla de ranking (por ejemplo "V1 (Clima)") y en el texto emergente de cada punto del gráfico.
- **Diagonal:** la celda de la diagonal (una variable sobre sí misma) se ve atenuada, con un guion, y no acepta foco ni escritura. Las demás celdas son campos de texto de ancho corto.
- **Selector de marca:** cada fila de la tabla de ranking tiene un selector con tres opciones: "—" (sin definir), "SÍ" y "NO".
- **Aviso ámbar de AE:** siempre que hay resultado, debajo de la tabla de ranking aparece el recuadro ámbar "La tabla de ranking se muestra en el orden en que se cargaron las variables. El criterio de orden del curso está pendiente de confirmar." Es el comportamiento esperado, no un fallo.
- **Cómo cargar los datos:** escribir en la matriz con la tecla Tab para saltar de una celda a la siguiente, fila por fila. Tab salta la diagonal. Los números se escriben como texto, sin decimales.

### 1. Datos de prueba

**D-AE1: las cuatro variables de la prueba AE.1**, con estos nombres y esta matriz (la fila influye sobre la columna):

| Variable | Nombre | V1 | V2 | V3 | V4 |
|---|---|---|---|---|---|
| V1 | Clima | — | 4 | 0 | 4 |
| V2 | Precio | 1 | — | 2 | 4 |
| V3 | Costos | 0 | 1 | — | 0 |
| V4 | Demanda | 0 | 2 | 0 | — |

**Cargar D-AE1 (procedimiento CA), partiendo de un RD:**

1. Pulse "AE" en la navegación.
2. Pulse "+ Agregar variable" cuatro veces.
3. En la columna "Nombre" de la tabla de variables escriba Clima, Precio, Costos y Demanda, en ese orden.
4. Haga clic en la primera celda de la matriz (V1 sobre V2) y escriba, con Tab entre celdas: 4, 0, 4 (fila V1); 1, 2, 4 (fila V2); 0, 1, 0 (fila V3); 0, 2, 0 (fila V4).

**D-AE2: secuencia de marcado de CP-70.** Se marcan las variables en un orden distinto al de carga: V4 con SÍ, V3 con SÍ, V2 con NO y, por último, V1 con SÍ.

**D-AE3: script de consola para la matriz grande (CP-77).** Genera 200 variables llamadas "Variable 1" a "Variable 200" con calificaciones enteras al azar de 0 a 4 (diagonal sin valor), las guarda con `Persistencia.guardar` y recarga la página. Pegar en la pestaña "Consola" (si el navegador pide escribir `allow pasting`, hágalo) y pulsar Enter:

```javascript
(() => {
  const n = 200;
  const variables = [], matriz = [], marcas = [];
  for (let i = 0; i < n; i++) {
    variables.push('Variable ' + (i + 1));
    marcas.push(null);
    matriz.push(Array.from({ length: n }, (_, j) => (i === j ? null : String(Math.floor(Math.random() * 5)))));
  }
  const guardado = Persistencia.cargar();
  guardado.matrizActiva = 'AE';
  guardado.datos.AE = { variables, matriz, marcas };
  console.log('Guardado:', Persistencia.guardar(guardado));
  location.reload();
})();
```

### 2. Registro de ejecución propio

Se completa antes de empezar. Es adicional al registro de la sección 0 del Módulo 1, que sigue valiendo para los datos del equipo. Aquí se anota el navegador de cada caso, porque CP-75 y CP-76 se hacen en navegadores distintos del resto.

| Dato | Valor |
|---|---|
| Persona que ejecuta | |
| Fecha | |
| Equipo y versión de Windows | |
| Ruta del archivo `index.html` probado | |
| Programa de hojas de cálculo y versión (CP-74) | |
| Resultado global | |

| Casos | Navegador y versión | Observaciones |
|---|---|---|
| CP-61 a CP-74 y CP-77 | | |
| CP-75 (Edge) | | |
| CP-76 (Firefox) | | |

### 3. Casos de prueba (CP-61 a CP-77)

Los casos continúan la numeración del Módulo 1 (CP-01 a CP-60). Cada caso parte de un RD, salvo que su precondición diga otra cosa. Las filas de la tabla de ranking se escriben "variable, motricidad, dependencia, cuadrante, proyección x, proyección y"; la columna del selector se indica aparte.

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-61 | Prueba AE.7 y estado limpio | RD. | 1. Observe la navegación al abrir.<br>2. Pulse "AE". | Hay ocho botones: BCG, EFI, EFE, MPC, PEYEA, MIE, GE y AE, y al abrir está resaltado BCG. Al pulsar AE se ve la sección "AE" con el texto de ayuda que empieza con "Análisis Estructural (MICMAC). Cargue las variables de su tema y califique de 0 a 4…", el título "Variables", el texto "Todavía no hay variables. Agregue al menos dos.", el botón "+ Agregar variable" y el botón "− Quitar la última variable" deshabilitado (atenuado). No hay tabla de matriz, ni recuadro rojo, ni gráfico. En resultados dice "Complete o corrija los datos para ver el resultado." "Exportar a Excel" está habilitado. | ☐ Sí<br>☐ No |
| CP-62 | Prueba AE.7 (una y cuatro variables) | RD. AE visible. | 1. Pulse "+ Agregar variable" una vez y observe.<br>2. Pulse "+ Agregar variable" tres veces más (cuatro variables). | Con una variable: la tabla de variables muestra la fila V1 con su campo "Nombre"; aparece el título "Matriz de influencias directas" con una sola celda, la de la diagonal, atenuada y con un guion; no hay ninguna celda editable; sin recuadro rojo; resultados "Complete o corrija los datos para ver el resultado."; el botón "− Quitar la última variable" ya está habilitado. Con cuatro variables: la matriz tiene los encabezados "Influye ↓ / sobre →", V1, V2, V3 y V4; 12 celdas editables y 4 celdas de diagonal atenuadas; sin recuadro rojo ni gráfico. | ☐ Sí<br>☐ No |
| CP-63 | Prueba AE.1 (cálculo) | RD. | 1. Cargue D-AE1 con el procedimiento CA.<br>2. Observe los resultados. | Sin recuadro rojo. Texto: "Corte de motricidad (eje Y): 4.00. Corte de dependencia (eje X): 4.00." Tabla de ranking con los encabezados Variable, Motricidad, Dependencia, Cuadrante, Proyección x, Proyección y y "¿Pasa a la siguiente etapa?", y cuatro filas en este orden: V1 (Clima), 8.00, 1.00, INDEPENDIENTES, 3.50, 4.95; V2 (Precio), 7.00, 7.00, AMBIGUAS, 0.00, 0.00; V3 (Costos), 1.00, 2.00, AUTONOMAS, -0.50, 0.71; V4 (Demanda), 2.00, 8.00, DEPENDIENTES, -3.00, 4.24. El selector de cada fila muestra "—". Debajo, el recuadro ámbar sobre el orden pendiente de confirmar. Después, el título "Validadas" y el texto "Ninguna variable marcada con SÍ." Se dibuja un gráfico. | ☐ Sí<br>☐ No |
| CP-64 | Prueba AE.1 (gráfico: puntos y cuadrantes) | D-AE1 cargado (CP-63). | 1. Observe el gráfico: los cuatro rótulos de cuadrante y los cuatro puntos. | El plano tiene el eje horizontal rotulado "Dependencia (eje X)" y el vertical "Motricidad (eje Y)", con los valores 0, 4 y 8.8 en cada eje. Los cuadrantes son: arriba a la izquierda INDEPENDIENTES (verde claro), arriba a la derecha AMBIGUAS (amarillo claro), abajo a la izquierda AUTONOMAS (celeste) y abajo a la derecha DEPENDIENTES (rosado). Hay cuatro puntos con los rótulos V1, V2, V3 y V4. Posición aproximada, como porcentaje del ancho desde la izquierda y del alto desde arriba del recuadro: V1 en 11 % y 9 % (cuadrante INDEPENDIENTES, cerca del borde superior); V2 en 80 % y 20 % (AMBIGUAS); V3 en 23 % y 89 % (AUTONOMAS, cerca del borde inferior); V4 en 91 % y 77 % (DEPENDIENTES, cerca del borde derecho). | ☐ Sí<br>☐ No |
| CP-65 | Prueba AE.1 y AE.4 (gráfico: cortes, diagonal y proyecciones) | D-AE1 cargado (CP-63). | 1. Observe las líneas punteadas del gráfico.<br>2. Fíjese en la línea diagonal y en la línea que sale de cada punto. | Hay dos líneas de corte discontinuas, una vertical y una horizontal, en la mitad del máximo observado (4 en cada eje): la vertical queda en el 45.5 % del ancho desde la izquierda y la horizontal en el 54.5 % del alto desde arriba, es decir, no están en el centro del recuadro. El recuadro del plano es cuadrado. Una diagonal de puntos va de la esquina inferior izquierda a la esquina superior derecha: es la diagonal de igualdad, donde motricidad y dependencia valen lo mismo, y queda a 45°. V2 (7 y 7) está exactamente sobre esa diagonal y su línea de proyección tiene longitud cero, por lo que no se ve. De V1, V3 y V4 sale una línea fina discontinua cuyo extremo toca la diagonal y que le queda perpendicular (90°). Esto quedó resuelto en la corrección de la geometría de `dibujarAE`: los dos ejes tienen los mismos píxeles por unidad. | ☐ Sí<br>☐ No |
| CP-66 | Prueba AE.1 (nombre completo en el gráfico) | D-AE1 cargado (CP-63). | 1. Deje el cursor quieto sobre el punto V1 durante un par de segundos.<br>2. Repita con V2, V3 y V4. | Aparece un texto emergente con el nombre completo, el cuadrante y los valores. V1: "Clima: INDEPENDIENTES, motricidad 8, dependencia 1". V2: "Precio: AMBIGUAS, motricidad 7, dependencia 7". V3: "Costos: AUTONOMAS, motricidad 1, dependencia 2". V4: "Demanda: DEPENDIENTES, motricidad 2, dependencia 8". Los textos están en el gráfico; que el navegador los muestre, y cuánto tarda, es comportamiento del navegador y no se pudo simular. | ☐ Sí<br>☐ No |
| CP-67 | Prueba AE.5 (diagonal bloqueada) | D-AE1 cargado (CP-63). | 1. Haga clic en la celda atenuada de V1 sobre V1.<br>2. Intente escribir un 3.<br>3. Repita con las celdas de V2, V3 y V4.<br>4. Haga clic en la celda V1 sobre V4, pulse Tab y observe dónde queda el cursor.<br>5. Haga clic en V2 sobre V1 y pulse Tab una vez. | Ninguna celda de la diagonal acepta foco ni escritura: no aparece cursor, no cambia ningún valor y los resultados no cambian (siguen los de CP-63). Tab no se detiene en la diagonal: desde V1 sobre V4 pasa a V2 sobre V1, y desde V2 sobre V1 pasa a V2 sobre V3 (se salta V2 sobre V2). En la simulación se comprobó que las celdas de la diagonal no son campos y no son enfocables; la tecla Tab real y el puntero no se pudieron simular. | ☐ Sí<br>☐ No |
| CP-68 | Prueba AE.6 (celdas sin calificar y valores fuera de rango) | RD. AE con cuatro variables y los nombres de D-AE1 (pasos 1 a 3 de CA). | 1. Escriba solo un 4 en V1 sobre V2.<br>2. Complete el resto de D-AE1 menos V4 sobre V3, que queda en blanco.<br>3. Escriba 5 en V4 sobre V3.<br>4. Cámbielo por -1.<br>5. Cámbielo por 2,5.<br>6. Cámbielo por 0. | Pasos 1 y 2: recuadro rojo con "Hay campos vacíos: complete todos los campos antes de calcular.", resultados "Complete o corrija los datos para ver el resultado." y sin gráfico. Pasos 3, 4 y 5: recuadro rojo con "Calificación de Demanda sobre Costos: debe ser un número entero entre 0 y 4.", sin tabla ni gráfico. Paso 6: el error desaparece y vuelve el resultado de CP-63 con su gráfico. Lo escrito se conserva en todos los pasos. | ☐ Sí<br>☐ No |
| CP-69 | Escenario AE-1 (agregar y quitar variables) | D-AE1 cargado (CP-63). | 1. Pulse "− Quitar la última variable" una vez.<br>2. Pulse "+ Agregar variable" una vez. | Paso 1: quedan tres variables (Clima, Precio y Costos) y la matriz de tres por tres con los mismos valores (4, 0 / 1, 2 / 0, 1); 6 celdas editables; sin recuadro rojo. "Corte de motricidad (eje Y): 2.00. Corte de dependencia (eje X): 2.50." Tabla de ranking de tres filas: V1 (Clima), 4.00, 1.00, INDEPENDIENTES, 1.50, 2.12; V2 (Precio), 3.00, 5.00, AMBIGUAS, -1.00, 1.41; V3 (Costos), 1.00, 2.00, AUTONOMAS, -0.50, 0.71. El gráfico tiene tres puntos. Paso 2: vuelve a haber cuatro variables, la cuarta sin nombre (se rotula V4) y con su fila y columna en blanco; 12 celdas editables; recuadro rojo con "Hay campos vacíos: complete todos los campos antes de calcular."; resultados "Complete o corrija los datos para ver el resultado." y sin gráfico. Los valores de las tres primeras variables no cambian. | ☐ Sí<br>☐ No |
| CP-70 | Prueba AE.8 (orden de carga) | D-AE1 cargado (CP-63), sin marcas. | 1. Marque con el selector, uno por uno y en este orden (D-AE2): V4 con SÍ, V3 con SÍ, V2 con NO, V1 con SÍ.<br>2. Después de cada marca, anote el orden de la lista Validadas.<br>3. Al final observe la tabla de ranking y el recuadro ámbar. | Lista Validadas tras cada paso: con V4 SÍ, "V4 (Demanda)"; con V3 SÍ, "V3 (Costos)" y "V4 (Demanda)" (V3 va antes aunque se marcó después); con V2 NO, igual que antes; con V1 SÍ, "V1 (Clima)", "V3 (Costos)" y "V4 (Demanda)". Es el orden de carga: no es el orden en que se marcaron (V4, V3, V1), ni el de motricidad de mayor a menor (V1, V4, V3), ni el de dependencia de mayor a menor (V4, V3, V1, que además coincide con el orden de marcado). La tabla de ranking sigue en el orden V1, V2, V3, V4 y los selectores muestran SÍ, NO, SÍ y SÍ. Sigue visible el recuadro ámbar sobre el orden pendiente de confirmar. Ninguna de las dos vistas se reordena. | ☐ Sí<br>☐ No |
| CP-71 | Prueba AE.8 (marca sin definir) | Marcas de CP-70 puestas. | 1. Cambie el selector de V3 a "—". | La lista Validadas queda con "V1 (Clima)" y "V4 (Demanda)". La fila V3 (Costos) de la tabla de ranking muestra "—" en el selector y conserva sus valores (1.00, 2.00, AUTONOMAS, -0.50, 0.71). | ☐ Sí<br>☐ No |
| CP-72 | Prueba AE.9 (ninguna con SÍ) | D-AE1 cargado y marcas de CP-70 puestas. | 1. Cambie a NO los selectores de V1, V3 y V4 (V2 ya está en NO). | La hoja Validadas muestra "Ninguna variable marcada con SÍ.", sin recuadro rojo. La tabla de ranking, los cortes y el gráfico siguen igual que en CP-63, con los cuatro selectores en NO. El módulo no se bloquea. | ☐ Sí<br>☐ No |
| CP-73 | Prueba AE.10 (recarga) | D-AE1 cargado, con las marcas V1 SÍ, V2 NO, V3 SÍ y V4 SÍ. | 1. Pulse F5.<br>2. Revise la sección, los campos, los selectores y el gráfico. | Después de F5 sigue abierta la sección AE, con el botón AE resaltado. Los nombres son Clima, Precio, Costos y Demanda, la matriz conserva los valores de D-AE1 y los selectores muestran SÍ, NO, SÍ y SÍ. Se ven los mismos resultados de CP-63 (cortes 4.00 y 4.00 y las cuatro filas) y el gráfico con cuatro puntos. La lista Validadas muestra "V1 (Clima)", "V3 (Costos)" y "V4 (Demanda)". | ☐ Sí<br>☐ No |
| CP-74 | Escenario AE-3 (exportación) y estructura [VERIFICAR] | D-AE1 cargado, con las marcas V1 SÍ, V2 NO, V3 SÍ y V4 SÍ. AE activa. Un programa de hojas de cálculo (Excel, LibreOffice u otro). | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo descargado con el programa de hojas de cálculo.<br>3. Revise las dos hojas.<br>4. En una celda libre escriba una fórmula que compruebe si es número, sobre algunas celdas de números (por ejemplo `=ESNUMERO(C2)` en la hoja Datos y `=ESNUMERO(B2)` en la hoja Resultados, en Excel en español; en inglés `=ISNUMBER(...)`). | Se descarga `Mtx-AE.xlsx` y se abre sin errores, sin recuadro rojo en la página. Tiene dos hojas, "Datos" y "Resultados". Datos (A1 a E5): fila 1 con Variable, Clima, Precio, Costos y Demanda; fila 2 Clima, vacío, 4, 0, 4; fila 3 Precio, 1, vacío, 2, 4; fila 4 Costos, 0, 1, vacío, 0; fila 5 Demanda, 0, 2, 0, vacío. Las celdas de la diagonal están vacías. Resultados (A1 a G8): fila 1 con Variable, Motricidad, Dependencia, Cuadrante, Proyección x, Proyección y y Marca; fila 2 Clima, 8, 1, INDEPENDIENTES, 3.5, 4.949747, SÍ; fila 3 Precio, 7, 7, AMBIGUAS, 0, 0, NO; fila 4 Costos, 1, 2, AUTONOMAS, -0.5, 0.707107, SÍ; fila 5 Demanda, 2, 8, DEPENDIENTES, -3, 4.242641, SÍ; fila 6 vacía; fila 7 Corte Y, 4; fila 8 Corte X, 4. Los valores de la matriz, la motricidad, la dependencia, las proyecciones y los cortes son celdas numéricas (la fórmula da VERDADERO); los nombres, el cuadrante y la marca son texto. La estructura es provisional y no se pudo comparar con la del Excel original [VERIFICAR]. | ☐ Sí<br>☐ No |
| CP-75 | Historia 7 y módulo AE: Edge | Windows con Edge. Misma copia local de `index.html`. RD en Edge. | 1. Abra `index.html` con Edge.<br>2. Cargue D-AE1 (procedimiento CA).<br>3. Repita lo de CP-64 y CP-65 (puntos, cuadrantes, cortes, diagonal y líneas de proyección).<br>4. Repita lo de CP-67 (clic, escritura y Tab en la diagonal). | Se cumple lo esperado en CP-63, CP-64, CP-65 y CP-67: mismos resultados, el gráfico se ve igual que en Chrome (cuatro cuadrantes rotulados, cuatro puntos con sus rótulos, las dos líneas de corte, la diagonal y las líneas de proyección) y la diagonal no acepta foco ni escritura. Sin mensajes de error en la consola. Anote el navegador y la versión en el registro. | ☐ Sí<br>☐ No |
| CP-76 | Historia 7 y módulo AE: Firefox | Windows con Firefox. Misma copia local de `index.html`. RD en Firefox. | 1. Abra `index.html` con Firefox.<br>2. Cargue D-AE1 (procedimiento CA).<br>3. Repita lo de CP-64 y CP-65.<br>4. Repita lo de CP-67. | Se cumple lo esperado en CP-63, CP-64, CP-65 y CP-67, igual que en CP-75. Anote el navegador y la versión en el registro. | ☐ Sí<br>☐ No |
| CP-77 | Requisito de tamaño de Inception (hasta 200 variables) | RD. Chrome (u otro navegador ya probado). Equipo en condiciones normales. | 1. Pegue el script D-AE3 en la consola y pulse Enter. La consola escribe "Guardado: true" y la página se recarga.<br>2. Observe cuánto tarda en aparecer la sección AE con sus datos, y anote la experiencia.<br>3. Desplácese por la matriz y por la tabla de ranking.<br>4. Escriba un 0 y luego un 4 en una celda cualquiera y observe qué tan fluido responde el recálculo.<br>5. Pase el cursor sobre un par de puntos del gráfico.<br>6. En la consola escriba `document.querySelectorAll('#matriz-ae input.celda').length`, y después lo mismo con `#matriz-ae .resultados tbody tr` y con `#matriz-ae .grafico circle`.<br>7. Haga un RD. | Tras la recarga se abre la sección AE con 200 variables, sin recuadro rojo. Hay el texto "Corte de motricidad (eje Y): …" con dos valores (varían, porque las calificaciones son al azar), una tabla de ranking de 200 filas desde "V1 (Variable 1)" hasta "V200 (Variable 200)" en orden de carga, el recuadro ámbar sobre el orden, la hoja Validadas con "Ninguna variable marcada con SÍ." y un gráfico con 200 puntos y 200 líneas de proyección. Los tres comandos del paso 6 devuelven 39800 (200 por 199 celdas editables), 200 y 200; la matriz tiene además 200 celdas de diagonal atenuadas. No se fija un tiempo como criterio de aprobación, porque ese umbral nunca se acordó con el profesor. Se marca "Sí" si el módulo carga, calcula y dibuja sin errores ni bloqueos; anote en observaciones cuánto tardó cada paso y si la escritura se sintió lenta, para llevarlo al profesor. | ☐ Sí<br>☐ No |

### 4. Resumen de cobertura

| Bloque | Casos | Rango |
|---|---|---|
| A. Estado limpio, carga y resultados | 3 | CP-61 a CP-63 |
| B. Gráfico de `dibujarAE` | 3 | CP-64 a CP-66 |
| C. Diagonal, validación y variables | 3 | CP-67 a CP-69 |
| D. Marcas, orden de ranking y hoja Validadas | 3 | CP-70 a CP-72 |
| E. Recarga y exportación | 2 | CP-73 y CP-74 |
| F. Edge y Firefox | 2 | CP-75 y CP-76 |
| G. Matriz grande | 1 | CP-77 |
| Total | 17 | CP-61 a CP-77 |

Lo que Construction II dejó pendiente y dónde se cubre:

| Pendiente de Construction II | Casos |
|---|---|
| Resultado visual real de `dibujarAE` | CP-64, CP-65, CP-66 |
| Bloqueo de la diagonal en pantalla (prueba AE.5) | CP-62, CP-67 |
| Orden de la tabla de ranking y de la hoja Validadas (pruebas AE.8 y AE.9) | CP-70, CP-71, CP-72 |
| Rendimiento con una matriz grande | CP-77 |
| Edge y Firefox | CP-75, CP-76 |
| Exportación abierta en un programa de hojas de cálculo | CP-74 |

Las pruebas AE.6, AE.7 y AE.10 de Elaboration I, que `tests/elaboration1.test.js` ya verifica sin pantalla, se cubren aquí en lo que tienen de visible: CP-68 y CP-61 a CP-62, y CP-73. No se escribió ningún caso para funcionalidad que no existe: no hay caso de criterio de orden del ranking distinto al orden de carga ni de otra estructura de exportación.

### Observaciones y puntos nuevos marcados [VERIFICAR] en esta fase

Observaciones al preparar el plan (no se cambió código):

1. El botón de la navegación y el título de la sección dicen "AE"; el nombre "Análisis Estructural" solo aparece como texto emergente del botón y en el párrafo de ayuda.
2. El recuadro ámbar sobre el orden del ranking se ve siempre que hay resultado, no solo en un caso límite. Es deliberado: el orden de carga es provisional (punto 3 de Elaboration I de este módulo).
3. El estado limpio de AE no tiene variables, así que la persona debe pulsar "+ Agregar variable" al menos dos veces antes de poder calificar nada.
4. Con la matriz de 200 variables la aplicación recalcula y redibuja todo el módulo en cada tecla. En esta fase no se fija ningún tiempo como criterio de aprobación y no se midió desde el punto de vista de una persona.

Puntos [VERIFICAR] nuevos:

1. Criterio de rendimiento para la matriz grande: cuánto tiempo es aceptable al escribir y al recalcular con hasta 200 variables, y si el curso necesita llegar a ese tamaño. Sigue abierto el punto 2 de la Inception de este módulo.
2. Quién ejecuta este plan, con qué frecuencia y cómo se registran los fallos: igual que el punto 5 de la Construction III del Módulo 1.

## Transition — Módulo 2: Análisis Estructural

Objetivo de la fase: publicar Análisis Estructural, verificarlo en su dirección real, y cerrar el ciclo VUP de este módulo. El resumen del código y la lista consolidada de [VERIFICAR] para el profesor no se repiten aquí: se actualizaron dentro del "## Transition" del Módulo 1, partes 3 y 4, porque describen el archivo completo. Esta fase no cambió `index.html` ni `tests/`, y no resolvió ningún punto [VERIFICAR]: solo publica y enumera.

### 1. Configuración de despliegue

No hay nada nuevo que configurar. El sitio de GitHub Pages ya existe (ver la parte 1 del Transition del Módulo 1: URL https://gerson-chumpitaz.github.io/mtx/, rama `master`, carpeta raíz, publicación por rama, HTTPS forzado) y se vuelve a publicar solo con cada `git push` a `master`.

Lo que se hizo en este módulo:

1. Se subieron los commits del módulo a `master`. El último push fue `36d5e83..f3e4464`; `f3e4464` es el último commit que modificó `index.html`.
2. Se esperó la publicación consultando cada 15 segundos, con el mismo método del Módulo 1, hasta que el estado pasó de `building` a `built` (algo menos de un minuto):

```bash
gh api repos/gerson-chumpitaz/mtx/pages/builds/latest
```

   La respuesta final trajo `status` = `built` y `commit` = `f3e4464`.
3. Se comprobó que el archivo servido es idéntico al del repositorio: el SHA-256 de la URL y el de `index.html` es `f648e0736d7f92c7f355cca3af0ae8dec369935ad7690bfdab0d439d390b5d06` (1 039 759 bytes). `VUP.md` también queda servido en la misma dirección, como ya advertía el Módulo 1.

Los commits de esta fase solo cambian `VUP.md`: al subirlos Pages vuelve a publicar, con el mismo `index.html`.

### 2. Resultado de la verificación en la URL real

Se abrió https://gerson-chumpitaz.github.io/mtx/ en el navegador integrado de la aplicación (Chromium), se partió de un estado limpio (`localStorage.removeItem('mtx.estado')`) y se recorrió AE con los datos de la prueba AE.1, con eventos del DOM. No se repitió la verificación de las otras siete matrices: sigue siendo válida la del Módulo 1.

| Comprobación | Resultado |
|---|---|
| La página carga | Sí: HTTP 200 y contenido idéntico a `index.html` del commit (ver arriba). |
| Navegación | Ocho botones (BCG, EFI, EFE, MPC, PEYEA, MIE, GE y AE). Al abrir está resaltado BCG. |
| Estado limpio de AE | Sin variables, con el texto "Todavía no hay variables. Agregue al menos dos." y "− Quitar la última variable" deshabilitado. |
| Carga de AE.1 | Cuatro variables (Clima, Precio, Costos y Demanda): 12 celdas editables y 4 celdas de diagonal sin campo. Sin recuadro de errores. |
| Resultados de AE.1 | "Corte de motricidad (eje Y): 4.00. Corte de dependencia (eje X): 4.00." Filas: V1 (Clima), 8.00, 1.00, INDEPENDIENTES, 3.50, 4.95; V2 (Precio), 7.00, 7.00, AMBIGUAS, 0.00, 0.00; V3 (Costos), 1.00, 2.00, AUTONOMAS, -0.50, 0.71; V4 (Demanda), 2.00, 8.00, DEPENDIENTES, -3.00, 4.24. Coinciden con la prueba AE.1 y con CP-63. |
| Gráfico | Cuatro puntos con los rótulos V1 a V4 y los rótulos de cuadrante INDEPENDIENTES, AMBIGUAS, AUTONOMAS y DEPENDIENTES; dos líneas de corte; los textos emergentes traen el nombre completo (por ejemplo "Clima: INDEPENDIENTES, motricidad 8, dependencia 1"). |
| Geometría corregida | El área de trazado mide 435 por 435 (`viewBox` de 525 por 520). La diagonal de igualdad mide 45° y las líneas de proyección de V1, V3 y V4 le quedan a 90°, con el extremo sobre la diagonal; V2, con motricidad igual a dependencia, tiene línea de longitud cero. Se midió desde la geometría del SVG, igual que en la corrección. |
| Diagonal bloqueada | Las celdas de la diagonal no tienen campo; un clic y una tecla sobre una de ellas no cambian el estado guardado. |
| Marcas y hoja Validadas | Con V4 SÍ, V3 SÍ, V2 NO y V1 SÍ (en ese orden), la hoja Validadas lista "V1 (Clima)", "V3 (Costos)" y "V4 (Demanda)": orden de carga. Se ve el recuadro ámbar sobre el orden pendiente. Con ninguna marcada dice "Ninguna variable marcada con SÍ." |
| Exportación | Se generó `Mtx-AE.xlsx` con las hojas "Datos" y "Resultados" y los mismos valores de CP-74 (cortes 4 y 4, proyección y 4.949747 para V1). Se comprobó leyendo el libro de SheetJS antes de que el navegador lo descargara, no abriéndolo en Excel. |
| Persistencia | Tras recargar la URL siguen la sección AE abierta, los nombres, la matriz, los selectores (SÍ, NO, SÍ y SÍ), los mismos resultados, el gráfico con cuatro puntos y la lista Validadas con V1, V3 y V4. |
| Errores de consola | Ninguno, ni en el recorrido ni tras recargar. |
| Solicitudes de red | Solo la del propio documento (`GET https://gerson-chumpitaz.github.io/mtx/`, HTTP 200, una por carga). Ningún recurso externo ni CDN. |

Al terminar se borró `mtx.estado` del origen para dejarlo limpio. Qué no se verificó: la URL solo se probó en el navegador integrado (no en Chrome, Edge ni Firefox por separado ni en un teléfono); los datos se escribieron con eventos simulados y no con el teclado; la exportación se validó leyendo el libro generado, no abriéndolo en un programa de hojas de cálculo (no se descargó ningún archivo); no se probó la matriz de 200 variables en la URL pública; y el plan de pruebas manual de este módulo (CP-61 a CP-77) no se ha ejecutado con una persona.

### 3. Reflexión

Preguntas de VUP. La plantilla original pide que las responda la persona a mano; estas cinco las redactó el juez (la IA) con autorización explícita de Gerson.

**1. ¿Qué fue lo más importante de la especificación?**

Respuesta: Haber derivado las reglas de Análisis Estructural del código VBA real del Excel del profesor, no de la teoría MICMAC de los libros. La teoría dice que los cortes van en el promedio; el Excel del profesor los pone en la mitad del máximo. Si la especificación se hubiera basado solo en la teoría, el programa habría calculado algo razonable pero distinto de lo que el curso espera.

**2. ¿Qué harías distinto?**

Respuesta: Pediría que la geometría del gráfico, específicamente si dos ejes necesitan la misma escala en píxeles para que una diagonal tenga sentido visual, se revisara en Elaboration II o en Construction I, cuando se diseña el componente, y no en Construction III, cuando ya se prueba. La fórmula estaba bien desde el principio; lo que faltó fue pensar en la geometría de la pantalla al mismo tiempo que en la fórmula.

**3. ¿Qué te sorprendió de cómo la IA implementó los requisitos?**

Respuesta: Que el error más interesante de todo el módulo no lo encontró ninguna de las 32 pruebas automatizadas, sino el propio plan de pruebas manual, al simular el gráfico y medir un ángulo con un transportador digital, por así decirlo. El código calculaba los números correctos y aun así el dibujo salía torcido. También me sorprendió que la IA se corrigiera a sí misma varias veces antes de entregarme algo, por ejemplo un error en su propio script de prueba, sin que yo tuviera que encontrarlo.

**4. ¿Cómo ayudó tener un plan de pruebas claro?**

Respuesta: Los diez casos de Elaboration I, con números calculados y verificados a mano desde el principio, se volvieron el patrón de referencia para todo lo demás: la implementación, las pruebas automatizadas y después el plan manual reutilizaron la misma matriz de ejemplo. Cualquier desviación tenía un número exacto contra el cual compararse, no una impresión de "se ve bien".

**5. ¿Qué agregarías si siguieras desarrollando el proyecto?**

Respuesta: Un botón para reiniciar los datos sin usar la consola del navegador, que quedó pendiente desde el Módulo 1, y resolver con el profesor el criterio de orden del ranking de Análisis Estructural, que nunca se pudo verificar porque la base de datos está protegida.

## Inception — Módulo 3: Radar Estratégico

Objetivo de la fase: abrir un tercer ciclo VUP dentro de este mismo documento, para el módulo de Radar Estratégico, con la misma estructura de la Inception original y de la del Módulo 2. Esta fase no escribe código, no define firmas de componentes ni pruebas Given-When-Then (eso es Elaboration I de este módulo) y no modifica `index.html` ni `tests/`. El módulo reutiliza la visión, la justificación y los requisitos no funcionales generales de la Inception original; aquí solo se agrega lo que es propio de Radar Estratégico.

### Nombre del módulo

Radar Estratégico (el profesor lo nombra "El Radar de la Posición Estratégica"; se basa en el modelo Execution Premium de Kaplan y Norton, los mismos autores del Balanced Scorecard).

### Alcance v1

Este módulo es una sola herramienta con cinco etapas internas fijas, no un conjunto de matrices independientes entre sí. Mide qué tan lejos está una organización de la gestión estratégica ideal. Todo lo que sigue se apoya en la revisión directa que hizo el juez del archivo original del profesor (ver "Fuentes consultadas").

Las cinco etapas, en orden:

1. **Movilización:** liderazgo ejecutivo para el cambio. 3 componentes.
2. **Traducción:** la estrategia en términos operacionales. 3 componentes.
3. **Alineamiento:** toda la organización en torno a la estrategia. 2 componentes.
4. **Motivación:** hacer de la estrategia el trabajo de todos. 3 componentes.
5. **Gestión:** la estrategia como proceso continuo. 3 componentes.

Estructura del contenido:

- Son 14 componentes en total (3 + 3 + 2 + 3 + 3). Cada componente se descompone en un número fijo de características a evaluar, de tres a cinco cada uno, 56 en total.
- Los nombres de las etapas, los componentes y las características, así como el texto de cada afirmación, son contenido fijo del profesor. El estudiante no los redacta ni los modifica.

Lo que hace la herramienta en v1:

- **Calificación.** El estudiante califica cada una de las 56 afirmaciones con un número de 0 a 5 que mide su nivel de concordancia con ella. La escala está invertida: el propio archivo lo advierte con la frase "a mayor intensidad de acuerdo, menor alejamiento y menor debe ser el número a utilizar". 0 significa "Estoy completamente de acuerdo" (el objetivo ideal ya se cumple, cero alejamiento) y 5 significa "Estoy en completo desacuerdo" (máximo alejamiento). No es una escala donde más puntaje es mejor. La escala completa, tal como figura en el Excel:
  - 0: Estoy completamente de acuerdo.
  - 1: Estoy bastante de acuerdo.
  - 2: Estoy algo de acuerdo.
  - 3: No estoy muy de acuerdo.
  - 4: No estoy casi nada de acuerdo.
  - 5: Estoy en completo desacuerdo.
- **Puntaje por componente.** El puntaje de cada uno de los 14 componentes es el promedio de los puntajes de sus características. En el Excel original la fórmula real de una celda es `=SUM(F23:F26)/4` para un componente de cuatro características, y otras hojas del mismo archivo que usan la misma mecánica calculan lo mismo con `=AVERAGE(...)`. En Mtx se usa el promedio (AVERAGE) en vez de replicar una suma dividida entre un número fijo escrito a mano: es matemáticamente idéntico y no depende de un número mágico que se desalinee si cambia la cantidad de características de un componente.
- **Gráfico radar.** Un radar de 14 puntas, una por componente. Se verificó directamente en el XML del gráfico del Excel, no por lectura visual, que el eje de valores va de 0 en el centro a 5 en el borde exterior. Como la escala está invertida, una figura que se estira hacia afuera en una punta señala un problema en ese componente, no una fortaleza, al revés de la lectura intuitiva habitual de un gráfico radar. Esto debe quedar explícito en la interfaz para que el estudiante no lo lea al revés.
- **Rótulos de las puntas.** El gráfico original de Excel no tiene ninguna etiqueta de categoría conectada a los datos: la serie solo referencia el rango de los 14 promedios, sin ningún rango de categorías, así que las 14 puntas aparecen sin nombre. El nombre completo de cada componente y el nombre corto de cada etapa existen como texto al lado del gráfico, pero no están enlazados a él. Decisión del juez: en Mtx sí se rotulan las puntas, al menos con el nombre corto de la etapa de cada componente, porque un radar de catorce puntas sin ninguna referencia visual no es utilizable para un estudiante. Es una mejora de interfaz, no un cambio de cálculo, en el mismo espíritu que las correcciones visuales ya aceptadas en Construction II del Módulo 1.

### Fuera de alcance (v1) de este módulo

- **Ejemplo 5S de Lean.** El mismo archivo contiene, en dos hojas adicionales ("5s situación anterior" y "situación actual con mejoras"), un ejemplo distinto armado con la misma mecánica de radar pero aplicado a la metodología 5S (Seiri, Seiton, Seiso, Seiketsu, Shitsuke), con un caso de antes y después de una empresa de servicio posventa. Es un ejemplo ilustrativo propio del profesor para otra aplicación de la misma herramienta genérica, no parte de lo que el estudiante del curso debe cargar para su Radar Estratégico. Queda fuera de v1.

### Historia de usuario

Como estudiante, quiero calificar mi nivel de acuerdo con cada una de las cincuenta y seis afirmaciones agrupadas en las cinco etapas de la gestión estratégica, para que el sistema calcule automáticamente el puntaje de cada uno de los catorce componentes y me muestre en un gráfico radar qué tan lejos estoy del objetivo ideal en cada uno.

Es una sola historia porque el módulo es una sola herramienta con cinco etapas internas fijas, no matrices independientes entre sí: el estudiante no recorre las etapas por separado ni obtiene un resultado útil de una sola, sino que califica las 56 afirmaciones y lee un único radar.

### Requisitos no funcionales específicos del módulo

Ya cubiertos por los requisitos no funcionales generales de la Inception original y por lo tanto no se repiten: cálculo íntegro en el navegador, apertura con doble clic sin instalación, interfaz usable por un estudiante sin conocimientos técnicos, y guardado automático del progreso en localStorage. Los que no están cubiertos:

- **Estructura fija, sin listas de longitud variable.** Esta es la diferencia estructural más importante frente a los módulos anteriores. El estudiante no agrega ni quita etapas, componentes ni características: solo califica los 56 ítems fijos del profesor. No aplica el patrón de botones "+" y "−" usado en Análisis Estructural, ni ninguna otra forma de alterar la cantidad de filas. La interfaz no debe ofrecer controles para agregar, quitar, renombrar ni reordenar etapas, componentes o características.
- **Fidelidad a la regla del profesor, no a la bibliografía.** El requisito general pide coincidir con las fórmulas clásicas verificables contra un caso de bibliografía conocido. En este módulo la referencia autoritativa es el Excel del profesor: puntaje de componente como promedio de sus características, escala de 0 a 5 con el sentido invertido, y eje del gráfico de 0 en el centro a 5 en el borde. Donde la presentación habitual de un gráfico radar y el Excel del profesor difieran, el sistema sigue al Excel, salvo las mejoras de interfaz aceptadas (rótulos de las puntas).
- **Lectura invertida explícita.** La interfaz debe dejar claro, junto a la calificación y junto al gráfico, que 0 es el mejor valor y 5 el peor, y que una punta que se estira hacia afuera indica un problema. Es un requisito de interfaz propio de este módulo, porque la lectura intuitiva de una escala y de un gráfico radar es la contraria.

### Riesgos de desarrollo y puntos [VERIFICAR] de este módulo

Riesgos:

- Para este modelo no existe ninguna guía del profesor. Ni la "Guía de uso, Cargar los Modelos Estratégicos" ni "Qué carga usted y qué sale solo" lo cubren, porque ambas se limitan explícitamente a los tres modelos que dependen de Access. El "Manual técnico" solo confirma, en una tabla, que este modelo no usa base de datos y calcula directamente en la hoja. Las reglas de este módulo salen por eso de la inspección directa del Excel (celdas, fórmulas y XML del gráfico), no de un documento que las explique.
- El contenido fijo del profesor trae al menos una inconsistencia propia, no de la lógica del sistema (punto 1 de la lista siguiente). Como el texto es del profesor y el estudiante no lo edita, cualquier error de redacción se hereda tal cual hasta que el profesor lo confirme.

Puntos [VERIFICAR]:

1. **Título repetido de dos componentes de Alineamiento (para el profesor).** El componente de la fila 69 y el de la fila 74 del Excel, ambos dentro de la etapa de Alineamiento, tienen exactamente el mismo título ("LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO"), aunque sus características a evaluar son distintas: las de la fila 74 hablan de reuniones con "unidades de soporte", no de unidades de negocio. Parece un error de copiar y pegar. No se corrige el texto por cuenta propia: falta que el profesor confirme el título correcto del segundo componente.
2. **Gráfico con ítems sin completar (de desarrollo, se resuelve en Elaboration I).** Falta decidir si debe exigirse completar las 56 características antes de mostrar el gráfico, o si se permite verlo parcial, con las que falten en su valor por defecto.

### Fuentes consultadas

Archivos originales del profesor, en `C:\Obsidian\Mi_Segundo_Cerebro\07_gestion_estrategica_software\Modelos Estrategicos\`, revisados directamente por el juez (la sesión de Cowork del Project "Prompt Engineering") abriendo sus celdas, sus fórmulas y el XML real de su gráfico, y usados aquí como hallazgos verificados:

- `RADAR ESTRATEGICO (1).xls`: el Excel del modelo. De él salen las cinco etapas, los 14 componentes y las 56 características, la escala invertida de 0 a 5, la fórmula del puntaje de cada componente, el eje del gráfico (0 a 5), la ausencia de etiquetas de categoría en el gráfico, el título repetido de la fila 69 y la fila 74, y las dos hojas del ejemplo 5S.
- `Manual tecnico - Modelos Estrategicos.pdf`: solo confirma, en una tabla, que este modelo no usa base de datos y calcula directamente en la hoja.
- `Guia de uso - Cargar los Modelos Estrategicos.pdf` y `Que carga usted y que sale solo.pdf`: no cubren este modelo, porque se limitan a los tres modelos que dependen de Access.

Este Inception se escribió a partir de esos hallazgos, tal como los entregó el juez. El constructor no volvió a abrir los archivos originales en esta fase.

## Elaboration I — Módulo 3: Radar Estratégico

Objetivo de la fase: definir, para la historia de usuario del módulo, pruebas de aceptación concretas en formato Given-When-Then (Dado / Cuando / Entonces), con valores numéricos y resultado esperado exacto. No hay código en esta fase. Los resultados esperados de la prueba RE.1 los verificó el juez aritméticamente y se transcriben tal cual; los de las demás pruebas se calcularon a mano a partir de las fórmulas de abajo.

Convenciones de esta sección:

- La numeración es "Prueba RE.n" (RE por Radar Estratégico) y no "Prueba X.Y", porque el módulo tiene una sola historia y las pruebas 1.1 a 7.1 y AE.1 a AE.10 ya pertenecen a los módulos anteriores.
- Los componentes se identifican por su posición dentro de la etapa ("componente 2 de Movilización"), no por su título. Esto es deliberado: dos componentes de Alineamiento comparten título en el documento original (punto 1 de los [VERIFICAR] de la Inception de este módulo) y por eso el título no sirve como identificador. El texto de las 56 afirmaciones es contenido fijo del profesor y no se transcribe aquí: las pruebas solo necesitan la calificación de cada una, que se anota en el orden en que aparecen dentro de su componente.
- Las calificaciones van de 0 a 5, con la escala invertida: 0 es el mejor valor (el objetivo ideal ya se cumple) y 5 el peor (máximo alejamiento).
- Los puntajes se comparan con dos decimales. En el radar, la distancia al centro de un punto es su puntaje dividido entre 5, como fracción del radio del gráfico.
- Las marcas [VERIFICAR] señalan una convención que debe confirmarse con el profesor o con el juez. Se consolidan al final de esta sección.

Estructura fija, transcrita de la Inception (ya verificada contra el Excel original, estas pruebas no la rederivan). Los números de "punta" son el orden de los 14 componentes alrededor del radar:

| Etapa | Componente (posición) | Características | Punta |
|---|---|---|---|
| Movilización | 1 | 4 | 1 |
| Movilización | 2 | 4 | 2 |
| Movilización | 3 | 4 | 3 |
| Traducción | 1 | 5 | 4 |
| Traducción | 2 | 4 | 5 |
| Traducción | 3 | 3 | 6 |
| Alineamiento | 1 | 4 | 7 |
| Alineamiento | 2 | 4 | 8 |
| Motivación | 1 | 4 | 9 |
| Motivación | 2 | 4 | 10 |
| Motivación | 3 | 4 | 11 |
| Gestión | 1 | 4 | 12 |
| Gestión | 2 | 4 | 13 |
| Gestión | 3 | 4 | 14 |

Total: 5 etapas, 14 componentes, 56 características.

Fórmulas verificadas, que estas pruebas no rederivan:

- Puntaje de un componente = promedio (AVERAGE) de las calificaciones de sus propias características. El divisor es la cantidad de características de ese componente (3, 4 o 5), no un número fijo.
- Un componente tiene tres estados, que se evalúan componente por componente, de forma independiente de los otros trece:
  - **Completo:** todas sus características tienen calificación válida. Se calcula su puntaje y se grafica su punto.
  - **Incompleto:** solo algunas de sus características están calificadas. No se calcula, no se grafica y no se muestra ningún error: es un estado normal de trabajo en progreso, no una entrada inválida.
  - **Vacío:** ninguna de sus características está calificada. No se calcula, no se grafica y no se muestra ningún error.
- Una calificación fuera de la escala 0 a 5 sí es una entrada inválida (ver la prueba RE.10).
- El gráfico muestra únicamente los puntos de los componentes completos.
- El gráfico es un radar de 14 puntas, una por componente, con el eje de valores de 0 en el centro a 5 en el borde exterior.

### Historia: Radar Estratégico

**Prueba RE.1: cuatro componentes completos y diez vacíos (caso principal)**

- Dado que el estudiante califica por completo solo cuatro componentes y deja los otros diez sin tocar:
  - Componente 1 de Movilización (4 características): 0, 1, 2, 3.
  - Componente 1 de Traducción (5 características): 0, 0, 5, 5, 5.
  - Componente 3 de Traducción (3 características): 0, 0, 0.
  - Componente 3 de Gestión (4 características): 5, 5, 5, 5.
- Cuando el sistema calcula el radar
- Entonces se cumple todo lo siguiente:
  - Resultado por componente:

    | Componente | Calificaciones | Divisor | Puntaje | Estado | Punta | Posición en el radar |
    |---|---|---|---|---|---|---|
    | Movilización 1 | 0, 1, 2, 3 | 4 | 1.50 | completo | 1 | al 30 % del radio |
    | Traducción 1 | 0, 0, 5, 5, 5 | 5 | 3.00 | completo | 4 | al 60 % del radio |
    | Traducción 3 | 0, 0, 0 | 3 | 0.00 | completo | 6 | en el centro |
    | Gestión 3 | 5, 5, 5, 5 | 4 | 5.00 | completo | 14 | en el borde exterior |

  - Los otros diez componentes quedan en estado "vacío": no tienen puntaje, no tienen punto en el gráfico y no muestran ningún error.
  - El gráfico muestra exactamente cuatro puntos.
  - Esta prueba ejercita a la vez el promedio con tres divisores distintos (4, 5 y 3), un promedio no entero (1.50), el extremo ideal (0.00, en el centro) y el extremo peor (5.00, en el borde). Traducción 3 con 0.00 es un componente completo en el ideal, no un componente sin calificar: el sistema debe distinguirlos, porque los dos tienen "cero" a la vista.

**Prueba RE.2: un componente incompleto no se calcula, y se calcula al completarlo**

- Dado el estado de la prueba RE.1, y que el estudiante califica solo dos de las cuatro características del componente 2 de Movilización: la primera con 1 y la segunda con 3, dejando la tercera y la cuarta sin calificar
- Cuando el sistema calcula el radar
- Entonces se cumple todo lo siguiente:
  - El componente 2 de Movilización queda en estado "incompleto": no tiene puntaje, no tiene punto en el gráfico y no se muestra ningún error (un promedio de 1 y 3 sobre dos valores, que daría 2.00, no debe aparecer).
  - El gráfico sigue mostrando exactamente los mismos cuatro puntos de la prueba RE.1, sin cambios.
  - Los otros nueve componentes siguen "vacíos".
- Cuando después el estudiante califica la tercera con 2 y la cuarta con 4
- Entonces el componente 2 de Movilización pasa a "completo" con puntaje (1 + 3 + 2 + 4) / 4 = 2.50, aparece su punto en la punta 2 al 50 % del radio, y el gráfico muestra cinco puntos. Los cuatro puntos anteriores no cambian.

**Prueba RE.3: el módulo entero vacío**

- Dado que el estudiante abre el módulo sin haber calificado ninguna de las 56 características
- Cuando el sistema calcula el radar
- Entonces los 14 componentes quedan en estado "vacío": ninguno tiene puntaje, el gráfico no tiene ningún punto y no se muestra ningún error. El sistema no bloquea la pantalla ni exige completar nada para poder abrir el módulo y empezar a calificar.

**Prueba RE.4: el eje es fijo de 0 a 5, no se ajusta a los datos**

- Dado que el estudiante califica por completo un solo componente, el componente 1 de Movilización, con 0, 1, 2, 3 (puntaje 1.50), y deja los otros trece sin tocar
- Cuando el sistema calcula el radar
- Entonces el gráfico muestra un solo punto, en la punta 1, al 30 % del radio (1.50 / 5). El eje sigue yendo de 0 en el centro a 5 en el borde exterior: el punto no se estira hasta el borde aunque sea el único ni el de mayor valor. Esta prueba detecta un gráfico que ajuste su escala al máximo observado.

**Prueba RE.5: catorce puntas rotuladas con la etapa de cada componente**

- Dado el estado de la prueba RE.1
- Cuando el estudiante mira el gráfico
- Entonces se cumple todo lo siguiente:
  - El gráfico tiene exactamente 14 puntas, una por componente, en el orden de la tabla de estructura fija, aunque solo cuatro de ellas tengan punto.
  - Cada punta está rotulada, al menos, con el nombre corto de la etapa de su componente: tres puntas "Movilización" (1 a 3), tres "Traducción" (4 a 6), dos "Alineamiento" (7 y 8), tres "Motivación" (9 a 11) y tres "Gestión" (12 a 14).
  - Los cuatro puntos de la prueba RE.1 caen en las puntas 1 (Movilización), 4 (Traducción), 6 (Traducción) y 14 (Gestión).

El Excel original no rotula las puntas: esta es la mejora de interfaz decidida por el juez en la Inception de este módulo, que no cambia ningún cálculo.

**Prueba RE.6: la lectura invertida está explícita en la interfaz**

- Dado que el estudiante abre el módulo, en cualquier estado de calificación
- Cuando mira las preguntas y el gráfico
- Entonces se cumple todo lo siguiente:
  - Junto a la calificación se muestran los seis niveles de la escala con su texto: 0 "Estoy completamente de acuerdo", 1 "Estoy bastante de acuerdo", 2 "Estoy algo de acuerdo", 3 "No estoy muy de acuerdo", 4 "No estoy casi nada de acuerdo" y 5 "Estoy en completo desacuerdo".
  - Junto a la escala o al gráfico se indica que 0 es el mejor valor (el objetivo ideal ya se cumple) y 5 el peor (máximo alejamiento).
  - Junto al gráfico se indica que un punto que se aleja del centro señala un problema en ese componente, no una fortaleza.

La redacción exacta de estos avisos se define en el diseño. La prueba solo fija que existan y que digan lo anterior.

**Prueba RE.7: la estructura es fija y no se puede alterar**

- Dado que el estudiante abre el módulo
- Cuando recorre las cinco etapas
- Entonces se cumple todo lo siguiente:
  - Hay exactamente 5 etapas, con 3, 3, 2, 3 y 3 componentes, es decir 14 componentes.
  - Los componentes tienen 4, 4, 4 (Movilización), 5, 4, 3 (Traducción), 4, 4 (Alineamiento), 4, 4, 4 (Motivación) y 4, 4, 4 (Gestión) características, es decir 56 características, cada una con su selector de calificación de 0 a 5.
  - No existe ningún control para agregar o quitar etapas, componentes o características: no hay botones "+" ni "−" como los de Análisis Estructural.
  - Los nombres de las etapas, los componentes y las características, y el texto de cada afirmación, no son editables.

**Prueba RE.8: los dos componentes de Alineamiento con el mismo título son distintos entre sí**

Ver el punto 1 de los [VERIFICAR] de la Inception de este módulo. Esta prueba no resuelve cuál es el título correcto del segundo componente: solo verifica que el sistema los trate como dos componentes independientes.

- Dado que el estudiante califica por completo solo el componente 2 de Alineamiento, con 3, 3, 3, 3, y deja el componente 1 de Alineamiento sin tocar
- Cuando el sistema calcula el radar
- Entonces el componente 2 de Alineamiento queda "completo" con puntaje 3.00, su punto cae en la punta 8, y el componente 1 de Alineamiento queda "vacío", sin punto en la punta 7, aunque los dos tengan el mismo título.
- Cuando después el estudiante califica por completo el componente 1 de Alineamiento con 1, 1, 1, 1
- Entonces el componente 1 de Alineamiento pasa a "completo" con puntaje 1.00 y su punto cae en la punta 7. El componente 2 sigue en 3.00 en la punta 8, sin cambios.

**Prueba RE.9: editar una calificación recalcula o devuelve el componente a incompleto**

- Dado el estado de la prueba RE.1
- Cuando el estudiante cambia la cuarta característica del componente 1 de Movilización de 3 a 5
- Entonces el puntaje de ese componente pasa a (0 + 1 + 2 + 5) / 4 = 2.00, su punto se mueve de 30 % a 40 % del radio, y los otros tres puntos no cambian.
- Cuando después el estudiante borra la calificación de la primera característica de ese mismo componente, dejándola sin calificar
- Entonces el componente 1 de Movilización pasa a "incompleto": su punto desaparece del gráfico, no se muestra ningún error, y el gráfico queda con los otros tres puntos. Las otras tres calificaciones del componente se conservan.

**Prueba RE.10: calificación fuera de rango**

Misma regla que el resto del sistema (como en AE.6): un valor fuera de la escala es una entrada inválida y no se calcula con él. La diferencia con un componente incompleto es que aquí sí se muestra un error.

- Dado el estado de la prueba RE.1, pero con la cuarta característica del componente 1 de Movilización en 6 (o −1) en vez de 3
- Cuando el sistema calcula el radar
- Entonces se cumple todo lo siguiente:
  - El componente 1 de Movilización queda en estado "inválido": se muestra un error que indica que la calificación debe estar entre 0 y 5, no tiene puntaje y no tiene punto en el gráfico.
  - Los otros tres componentes completos de la prueba RE.1 no se ven afectados: siguen con 3.00, 0.00 y 5.00 y con sus puntos en el gráfico.
  - El valor escrito se conserva: el estudiante no pierde lo que ya escribió.

El mecanismo por el cual un valor fuera de rango puede llegar al sistema (un campo numérico, datos guardados alterados) se define en el diseño. La prueba solo fija la respuesta del sistema si llega.

**Prueba RE.11: los datos siguen tras recargar**

- Dado el estado de la prueba RE.2 antes de completar el componente 2 de Movilización: los cuatro componentes completos de la prueba RE.1 y el componente 2 de Movilización con solo dos características calificadas (1 y 3)
- Cuando el estudiante recarga la página o vuelve a abrir el archivo
- Entonces las calificaciones de las 56 características siguen ahí, tal como estaban, incluidas las dos del componente incompleto. El sistema recalcula los mismos resultados de la prueba RE.1: cuatro puntos con 1.50, 3.00, 0.00 y 5.00, y el componente 2 de Movilización sigue "incompleto", sin punto. Esto aplica al módulo el requisito general de guardado automático de la Inception original.

### Resumen de la fase

| Prueba | Qué verifica | Estado |
|---|---|---|
| RE.1 | Promedio con distinto número de características, extremo ideal (centro), extremo peor (borde), no entero, diez componentes vacíos | Con prueba concreta, verificada por el juez |
| RE.2 | Componente incompleto sin cálculo, ni punto, ni error; paso a completo | Con prueba concreta |
| RE.3 | Módulo entero vacío | Con prueba concreta |
| RE.4 | Eje fijo de 0 a 5, sin ajuste a los datos | Con prueba concreta |
| RE.5 | Catorce puntas rotuladas con la etapa | Con prueba concreta |
| RE.6 | Lectura invertida explícita en la interfaz (escala de seis niveles, 0 mejor, 5 peor) | Con prueba concreta |
| RE.7 | Estructura fija: 5 etapas, 14 componentes, 56 características, sin controles para alterarla | Con prueba concreta |
| RE.8 | Los dos componentes de Alineamiento con el mismo título son independientes | Con prueba concreta, título correcto pendiente [VERIFICAR] |
| RE.9 | Editar o borrar una calificación recalcula o devuelve a incompleto | Con prueba concreta |
| RE.10 | Calificación fuera de rango | Con prueba concreta |
| RE.11 | Persistencia al recargar | Con prueba concreta |

Total: 11 pruebas Given-When-Then.

Nota de consistencia con el Módulo 2: aquí un componente parcialmente calificado queda "incompleto", sin error, mientras que en Análisis Estructural una matriz parcialmente calificada es "inválida". La diferencia es arquitectónica, no un olvido: la matriz de Análisis Estructural es un solo objeto interdependiente (una celda afecta la motricidad y la dependencia de dos variables a la vez), mientras que cada componente del Radar se calcula de forma completamente independiente de los otros trece.

Puntos de esta fase:

1. **Criterio de completitud, resuelto por el juez.** Un componente se calcula y se grafica si todas sus características están calificadas, por componente y no por módulo. Con esto queda resuelto el punto 2 de los [VERIFICAR] de la Inception de este módulo (si exigir las 56 características antes de mostrar el gráfico).
2. **Título repetido de los dos componentes de Alineamiento [VERIFICAR].** Sigue pendiente del profesor (punto 1 de la Inception de este módulo). La prueba RE.8 no depende de cuál sea el título correcto.
3. **Calificaciones con decimales [VERIFICAR].** La escala es de 0 a 5 con seis niveles con nombre. Ninguna prueba fija si un valor como 2.5 se rechaza. Se asume que sí, igual que en los otros dos módulos, pero no se fija con una prueba numérica.
4. **Redacción y mecanismos que se definen en el diseño.** Las pruebas RE.6 y RE.10 no fijan la redacción exacta de los avisos ni de los mensajes de error, y la prueba RE.3 no fija si el gráfico vacío muestra su marco con las 14 puntas rotuladas o un aviso. Tampoco se fija cómo se identifica visualmente un componente "incompleto" frente a uno "vacío" en la interfaz (si es que se distinguen); las pruebas solo exigen que ninguno tenga puntaje, punto ni error.

## Elaboration II — Módulo 3: Radar Estratégico

Objetivo de la fase: definir qué suma Radar Estratégico a la arquitectura de seis componentes ya fijada en Elaboration II del Módulo 1, con escenarios y diagramas, sin escribir código. Las fórmulas, los estados de cada componente y las reglas de validación ya fijados en Elaboration I de este módulo (pruebas RE.1 a RE.11) no se reinterpretan aquí; los componentes solo las aplican. Lo que esta fase no resuelve se marca [VERIFICAR] y se consolida al final.

Principios de la arquitectura: se mantienen los del Módulo 1 sin cambios. La Vista es el único orquestador, el Validador y el Motor de Cálculo no dependen del navegador, los datos ingresados se guardan en cada cambio sean válidos o no, y los resultados solo se calculan y muestran cuando la validación pasa. Lo que Radar Estratégico suma a la fuente de verdad: lo que el usuario ingresó son las calificaciones de las 56 características, y algunas pueden estar sin calificar. El puntaje de cada componente, su estado (completo, incompleto o vacío) y la posición de su punto en el radar se recalculan cuando hacen falta y no se guardan. La estructura (etapas, componentes y características) es contenido fijo del profesor y no forma parte de los datos del estudiante, así que tampoco se guarda. Una diferencia con los módulos anteriores: aquí la validación y el cálculo trabajan componente por componente, no sobre el módulo completo. Un componente que no está completo no corta el flujo de los otros trece (ver el escenario RE-2).

### 1. Componentes del framework

No se repite la tabla de los seis componentes del Módulo 1. Solo se agrega lo que Radar Estratégico suma a cada uno.

| Componente | Qué suma Radar Estratégico |
|---|---|
| Vista | Ningún método nuevo. Reutiliza `renderFormulario`, `renderResultados` y `despacharEvento` con un noveno valor de matriz, `"RADAR"`, para las 56 calificaciones agrupadas en cinco etapas y el radar de 14 puntas. Pregunta abierta para Construction I: si conviene seguir llamando "matriz" a ese discriminador cuando este módulo no tiene ninguna matriz NxN, o generalizar el nombre del campo ahora que ya no describe solo matrices [VERIFICAR]. |
| Validador | Reutiliza `validarRango` con los límites de este módulo (0 a 5) sobre cada calificación (RE.10). La clasificación de cada componente en completo, incompleto o vacío no es el patrón binario de vacía o inválida que ya existe en los otros módulos, y probablemente necesita un método o ayudante nuevo. Pregunta abierta para Construction I, no se decide aquí [VERIFICAR]. |
| Motor de Cálculo | Método nuevo `calcularRadar(calificaciones, errores)`. Aplica exactamente la regla ya fijada y probada en Elaboration I: el puntaje de cada componente es el promedio de las calificaciones de sus propias características, con el número de características de ese componente como divisor, y solo se calcula si el componente está completo y ninguna de sus características tiene un error de rango. Devuelve un `ResultadoRadar` con el estado de los 14 componentes (completo, incompleto, vacío o inválido) y el puntaje de los completos. |
| Motor de Gráficos | Método nuevo `dibujarRadar(resultado)`. Dibuja el radar de 14 puntas, cada una rotulada con el nombre corto de la etapa de su componente (RE.5), con el eje fijo de 0 en el centro a 5 en el borde, sin ajustarse al máximo observado (RE.4). Grafica solo los puntos de los componentes completos. |
| Persistencia | Sin cambios: se reutiliza tal cual. Guarda las 56 calificaciones, con sus huecos, como parte del estado de la sesión. |
| Exportador | Sin cambios en su firma. Qué contiene el archivo de este módulo queda abierto en el escenario RE-4. |

Consecuencia que debe recogerse en Construction I de este módulo: la decisión de Construction I dice que `matriz` es siempre un string con una de ocho siglas (las siete del Módulo 1 y `"AE"`). Con `"RADAR"` pasan a ser nueve, por lo que esa frase quedará desactualizada y pedirá un commit de sincronización en esa fase, igual que cuando se agregó `"AE"`. Tampoco se decide aquí cómo se acomodan las calificaciones de este módulo en el estado guardado ni si el formato guardado cambia de versión [VERIFICAR].

Relación con las pruebas de Elaboration I de este módulo. Se usa una tabla aparte de las de los módulos anteriores, porque aquí las pruebas se cruzan contra los dos métodos nuevos y no contra componentes completos:

| Prueba | `calcularRadar` | `dibujarRadar` | Otros componentes que la ejercitan |
|---|---|---|---|
| RE.1 | Cuatro promedios con divisores 4, 5, 3 y 4 (1.50, 3.00, 0.00 y 5.00) y diez componentes vacíos | Cuatro puntos: al 30 % del radio, al 60 %, en el centro y en el borde | Vista |
| RE.2 | Componente 2 de Movilización incompleto sin puntaje, y luego completo con 2.50 | El punto no aparece mientras está incompleto y aparece en la punta 2 al 50 % del radio al completarlo | Vista |
| RE.3 | Los 14 componentes vacíos | Radar sin ningún punto | Vista |
| RE.4 | Puntaje 1.50 de un solo componente | Punto al 30 % del radio, con el eje fijo de 0 a 5 | |
| RE.5 | No interviene | 14 puntas rotuladas con la etapa, y los cuatro puntos en las puntas 1, 4, 6 y 14 | |
| RE.6 | No interviene | No interviene | Vista (escala de seis niveles y avisos de lectura invertida) |
| RE.7 | No interviene | No interviene | Vista (estructura fija y ningún control para alterarla) |
| RE.8 | Dos componentes de Alineamiento calculados de forma independiente | Puntos en las puntas 7 y 8 | |
| RE.9 | Recalcula el puntaje al editar, y devuelve el componente a incompleto al borrar | El punto se mueve o desaparece | Vista y Persistencia |
| RE.10 | No calcula el componente inválido y calcula los otros tres | Sin punto para el componente inválido | Validador y Vista (muestra el error) |
| RE.11 | Recalcula los mismos resultados de RE.1 y el componente incompleto sigue incompleto | Vuelve a dibujar los mismos cuatro puntos | Persistencia, Validador y Vista |

### 2. Escenarios (plays)

**Escenario RE-1: flujo principal, del calificar al radar**

El estudiante abre el módulo de Radar Estratégico y empieza a calificar características. Con cada cambio, la Vista pide a la Persistencia que guarde el estado y pasa las calificaciones al Validador, que confirma que todos los valores escritos están entre 0 y 5. La Vista pasa las calificaciones al Motor de Cálculo, que clasifica cada uno de los 14 componentes y calcula el puntaje de los que están completos. La Vista entrega ese resultado al Motor de Gráficos, que dibuja el radar con sus 14 puntas rotuladas por etapa y el eje fijo de 0 a 5, con el punto de cada componente completo. Mientras el estudiante avanza, el radar va mostrando los puntos solo de los componentes que se completan. Con las calificaciones de la prueba RE.1 el radar muestra cuatro puntos, con 1.50, 3.00, 0.00 y 5.00 en las puntas 1, 4, 6 y 14, y las otras diez puntas siguen sin punto aunque ya estén rotuladas. Junto a la calificación y al gráfico, la Vista muestra la escala de seis niveles y los avisos de lectura invertida (RE.6).
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.

**Escenario RE-2: componente sin completar, sin rechazo**

El estudiante califica solo dos de las cuatro características del componente 2 de Movilización, como en la prueba RE.2, o no ha calificado ninguna característica de un componente, como en la prueba RE.3. La Vista guarda el estado en la Persistencia y pasa las calificaciones al Validador, que no encuentra ningún valor fuera de rango y no devuelve ningún error. El Motor de Cálculo clasifica ese componente como "incompleto" (o "vacío") y no calcula su puntaje. El Motor de Gráficos no recibe ni dibuja ningún punto para ese componente, y la Vista no muestra ningún aviso de error ni bloquea nada: los componentes completos siguen con su puntaje y su punto. Cuando el estudiante termina de calificar las características que faltan, el componente pasa a "completo" y su punto aparece. A diferencia del Escenario AE-2, aquí no hay rechazo ni mensaje de error, y el flujo no se corta: el Motor de Cálculo y el Motor de Gráficos sí se invocan.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
Este escenario cubre las pruebas RE.2 y RE.3. [VERIFICAR] si el filtro que deja fuera del gráfico a los componentes no completos lo hace el Motor de Cálculo (`ResultadoRadar` trae puntaje solo en los completos) o el Motor de Gráficos (`dibujarRadar` recibe los 14 estados y decide qué dibuja). Se decide en Construction I, junto con la forma de `ResultadoRadar`.

**Escenario RE-3: calificación fuera de rango**

El estudiante escribe, o llega al sistema, una calificación fuera de 0 a 5 en una característica (por ejemplo 6) del componente 1 de Movilización, como en la prueba RE.10. La Vista guarda el estado en la Persistencia, sin perder el valor escrito, y pasa las 56 calificaciones al Validador, que detecta el valor fuera de rango con `erroresRadar` y devuelve el error asociado a esa característica. La Vista pasa las 56 calificaciones y esos errores al Motor de Cálculo con `calcularRadar(calificaciones, errores)`, que marca ese componente como "inválido" sin calcularlo, y calcula normalmente los otros trece. La Vista muestra el error del componente inválido. El Motor de Gráficos no dibuja su punto, pero sí el de los componentes completos, que mantienen su puntaje. Es la diferencia con los módulos anteriores, donde un error de validación corta todo el flujo.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
Este escenario cubre la prueba RE.10. [VERIFICAR] cómo se excluye el componente inválido del cálculo: si el Validador devuelve los errores por componente y la Vista le pasa al Motor de Cálculo solo los componentes sin error, o si `calcularRadar` reconoce el estado "inválido" por sí mismo. Se decide en Construction I, junto con la pregunta abierta del Validador de la tabla de arriba.

**Escenario RE-4: exportación a Excel**

Con el radar de RE.1 ya calculado, el usuario pulsa "Exportar". La Vista pasa los datos al Validador, que revisa el rango de las calificaciones, y pide el resultado al Motor de Cálculo con `calcularRadar(calificaciones, errores)`. La Vista entrega al Exportador el módulo, los datos ingresados y el resultado. El Exportador genera el archivo .xlsx con la librería embebida y la Vista dispara la descarga, todo sin conexión a internet. Un componente incompleto, vacío o inválido no impide exportar: la invalidez de este módulo es por componente y no del módulo entero (a diferencia del escenario 4 del Módulo 1), así que la Vista muestra los errores de los componentes inválidos y llama igual al Exportador, que escribe el estado real de cada componente. Solo se pide completar datos si no hay ninguna característica calificada (módulo vacío).
Componentes: Vista, Validador, Motor de Cálculo, Exportador.
Resuelto en Construction II del Módulo 3: dos hojas, "Datos" con las 56 calificaciones y "Resultados" con los 14 componentes y su estado real, y se puede exportar con componentes sin completar (el puntaje se escribe solo en los completos). Como referencia, el Módulo 1 exporta dos hojas, "Datos" y "Resultados", con solo valores y sin imagen del gráfico, de forma provisional.

**Escenario RE-5: recuperación tras recargar**

El estudiante recarga la página o vuelve a abrir el archivo, con el estado de la prueba RE.11 ya guardado: cuatro componentes completos y el componente 2 de Movilización con dos características calificadas. La Vista pide a la Persistencia el estado guardado. Si existe, la Vista vuelve a llenar las 56 calificaciones, incluidas las de un componente incompleto, y el Validador revisa los datos. El Motor de Cálculo recalcula los mismos resultados de RE.1 y clasifica de nuevo el componente 2 de Movilización como incompleto, y el Motor de Gráficos vuelve a dibujar los mismos cuatro puntos. Si el almacenamiento está vacío, corrupto o inaccesible, la Persistencia devuelve un estado vacío y la Vista muestra el módulo en blanco, con los 14 componentes vacíos.
Componentes: Vista, Persistencia, Validador, Motor de Cálculo, Motor de Gráficos.
Este escenario cubre la prueba RE.11. Aplica a este módulo el mismo punto abierto del escenario 5 del Módulo 1 (aviso al descartar un estado corrupto), sin agregar otros: el mensaje de error de una matriz incompleta tras la recarga no aplica, porque aquí un componente incompleto no es un error.

Cobertura de componentes por escenario:

| Componente | RE-1 | RE-2 | RE-3 | RE-4 | RE-5 |
|---|---|---|---|---|---|
| Vista | sí | sí | sí | sí | sí |
| Validador | sí | sí | sí | sí | sí |
| Motor de Cálculo | sí | sí | sí | sí | sí |
| Motor de Gráficos | sí | sí | sí | no | sí |
| Persistencia | sí | sí | sí | no | sí |
| Exportador | no | no | no | sí | no |

### 3. Diagramas

**Diagrama de clases**

No se crea un diagrama nuevo. Se editó el `classDiagram` de la Elaboration II del Módulo 1 (arriba, en este documento) para agregar `+calcularRadar(calificaciones, errores) ResultadoRadar` a `MotorCalculo` y `+dibujarRadar(resultado)` a `MotorGraficos`. Las demás clases y relaciones no cambian.

**Diagrama de secuencia del escenario RE-1: flujo principal**

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: califica características de uno o varios componentes
    V->>P: guardar(estado)
    V->>VA: validar(RADAR, datos)
    VA-->>V: ninguna calificación fuera de rango
    V->>MC: calcularRadar(calificaciones, errores)
    MC-->>V: estado de los 14 componentes y puntaje de los completos
    V->>MG: dibujarRadar(resultado)
    MG-->>V: radar de 14 puntas rotuladas con los puntos de los completos
    V-->>U: muestra el radar, los puntajes y los avisos de lectura invertida
```

**Diagrama de secuencia del escenario RE-2: componente sin completar, sin rechazo**

A diferencia del diagrama del Escenario AE-2, el flujo no se corta después del Validador: el cálculo y el gráfico se ejecutan igual, y no hay ningún mensaje de error.

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: deja un componente con características sin calificar
    V->>P: guardar(estado)
    V->>VA: validar(RADAR, datos)
    VA-->>V: ninguna calificación fuera de rango, sin error
    V->>MC: calcularRadar(calificaciones, errores)
    MC-->>V: ese componente incompleto o vacío sin puntaje, los demás con su estado
    V->>MG: dibujarRadar(resultado)
    MG-->>V: radar sin punto para ese componente
    V-->>U: muestra el radar sin aviso de error
```

**Diagrama de secuencia del escenario RE-3: calificación fuera de rango**

Tampoco se corta todo el flujo: el error afecta solo a un componente.

```mermaid
sequenceDiagram
    autonumber
    actor U as Usuario
    participant V as Vista
    participant P as Persistencia
    participant VA as Validador
    participant MC as Motor de Cálculo
    participant MG as Motor de Gráficos
    U->>V: escribe una calificación fuera de 0 a 5
    V->>P: guardar(estado)
    V->>VA: validar(RADAR, datos)
    VA-->>V: error de rango en una característica
    V->>MC: calcularRadar(calificaciones, errores)
    MC-->>V: ese componente "inválido", estado y puntaje de los demás
    V->>MG: dibujarRadar(resultado)
    MG-->>V: radar sin punto para el componente inválido
    V-->>U: muestra el error de ese componente y el radar de los demás
```

Los escenarios RE-4 y RE-5 no tienen diagrama propio: siguen la misma secuencia que los escenarios 4 y 5 del Módulo 1, con el valor de matriz `"RADAR"` y las llamadas `calcularRadar` y `dibujarRadar` en lugar de las de EFI.

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Discriminador `matriz`: si se sigue llamando así con el noveno valor `"RADAR"` o se generaliza el nombre del campo. Se decide en Construction I, junto con la sincronización de la frase de las ocho siglas.
2. Validador: si la clasificación de cada componente en completo, incompleto o vacío necesita un método o ayudante nuevo, y cómo se excluye de los cálculos un componente con una calificación fuera de rango (escenario RE-3). Se decide en Construction I.
3. Filtro de los componentes no completos: si lo hace `calcularRadar` o `dibujarRadar`, y la forma de `ResultadoRadar` (escenario RE-2). Se decide en Construction I.
4. Estructura del Excel exportado del módulo: calificaciones crudas, promedios por componente o ambas en hojas separadas, y qué se escribe para los componentes no completos (escenario RE-4).
5. Estado guardado: cómo se acomodan las 56 calificaciones con sus huecos, y si el formato guardado cambia de versión. Se decide en Construction I.
6. El punto abierto de recuperación de sesión del Módulo 1 (aviso al descartar un estado corrupto) aplica también a este módulo.
7. Siguen abiertos los puntos de Elaboration I de este módulo: el título repetido de los dos componentes de Alineamiento, las calificaciones con decimales, y la redacción de avisos y errores junto con el aspecto del radar vacío.

## Construction I — Módulo 3: Radar Estratégico

Objetivo de la fase: resolver las preguntas que Elaboration II de este módulo dejó abiertas para Construction I, y dejar el esqueleto del proyecto extendido con los stubs de Radar Estratégico, sin comportamiento real. Los cuerpos de los métodos están vacíos o llevan un comentario de marcador de posición. Ninguna lógica de cálculo, validación, graficado ni renderizado se implementa en esta fase, y esta fase no toca `index.html` ni `tests/`: el esqueleto sigue viviendo dentro de este documento.

### 1. Decisiones de arquitectura

El stack no cambia: archivo único `index.html`, HTML, CSS y JavaScript sin framework ni backend, SVG para los gráficos, SheetJS embebido y localStorage. Se aplica tal cual a este módulo.

Las preguntas abiertas se resuelven con decisiones del juez:

| Pregunta abierta en Elaboration II | Decisión |
|---|---|
| ¿Se sigue llamando `matriz` al discriminador, con un noveno valor, o se generaliza el nombre? | Se mantiene el nombre, con un valor nuevo, `"RADAR"`. Cambiarlo tocaría los ocho usos ya existentes del campo por un beneficio solo cosmético, y el campo ya funciona como discriminador de tipo desde que incluye PEYEA (un vector) y MIE (una cuadrícula), que tampoco son matrices de influencia en sentido estricto. "Matriz" es el nombre histórico del campo, no una descripción literal de su contenido. |
| ¿El Validador necesita un método nuevo para clasificar los estados de un componente? | No. Ningún método público nuevo. `validar('RADAR', datos)` reutiliza `validarRango` (límites 0 a 5) sobre cada una de las 56 calificaciones que tengan un valor, con una auxiliar privada, `erroresRadar`, que solo revisa el rango sobre las calificaciones presentes. Una calificación sin valor (sin calificar) nunca es un error de rango: ausencia no es invalidez. `erroresRadar` devuelve los errores por característica, no por componente. La clasificación en completo, incompleto o vacío no es del Validador: la hace `calcularRadar`. |
| ¿Quién filtra los componentes no completos y cómo se excluye uno inválido? ¿Cuál es la forma de `ResultadoRadar`? | `calcularRadar(calificaciones, errores)` recibe las 56 calificaciones y el resultado de `erroresRadar`, y devuelve un arreglo de 14 posiciones (forma abajo). `dibujarRadar(resultado)` solo recorre ese arreglo y dibuja un punto donde `estado` sea `"completo"`: no vuelve a decidir nada que ya haya decidido `calcularRadar`. |
| ¿Cómo se acomodan las calificaciones en el estado guardado, y cambia la versión del formato? | Es una clave nueva, `datos.RADAR`, junto a las otras ocho. No cambia la versión del formato guardado. `Persistencia.cargar()` debe devolver `RADAR` con 56 `null` si un estado guardado anterior no la tiene. Su forma está justo debajo. |
| ¿Qué contiene el archivo exportado (escenario RE-4)? | No se resuelve en esta fase: sigue abierto para Construction II. |

**Forma de `ResultadoRadar`.** Un arreglo de 14 posiciones, una por componente, en el orden fijo de la tabla de estructura de Elaboration I de este módulo (Movilización 1, 2 y 3, Traducción 1, 2 y 3, Alineamiento 1 y 2, Motivación 1, 2 y 3, Gestión 1, 2 y 3, que son también las puntas 1 a 14 del radar). Cada posición es `{ estado, puntaje }`:

- `estado: "inválido"` si alguna de las características de ese componente tiene un error de rango, sin importar cuántas otras estén calificadas. `puntaje: null`.
- Si ninguna de sus características tiene error:
  - `estado: "completo"` con `puntaje` (el promedio de sus calificaciones) si todas tienen valor.
  - `estado: "incompleto"` con `puntaje: null` si algunas tienen valor y otras no.
  - `estado: "vacío"` con `puntaje: null` si ninguna tiene valor.

Con esto, `ResultadoRadar` agrega a los tres estados de Elaboration I el cuarto, `"inválido"`, que corresponde a la prueba RE.10. Ejemplo con el caso de la prueba RE.1, donde las posiciones 0, 3, 5 y 13 son completas y las otras diez están vacías:

```text
[
  { estado: "completo", puntaje: 1.5 },   // Movilización 1
  { estado: "vacío",    puntaje: null },  // Movilización 2
  { estado: "vacío",    puntaje: null },  // Movilización 3
  { estado: "completo", puntaje: 3.0 },   // Traducción 1
  { estado: "vacío",    puntaje: null },  // Traducción 2
  { estado: "completo", puntaje: 0.0 },   // Traducción 3
  { estado: "vacío",    puntaje: null },  // Alineamiento 1
  { estado: "vacío",    puntaje: null },  // Alineamiento 2
  { estado: "vacío",    puntaje: null },  // Motivación 1
  { estado: "vacío",    puntaje: null },  // Motivación 2
  { estado: "vacío",    puntaje: null },  // Motivación 3
  { estado: "vacío",    puntaje: null },  // Gestión 1
  { estado: "vacío",    puntaje: null },  // Gestión 2
  { estado: "completo", puntaje: 5.0 }    // Gestión 3
]
```

Esta decisión cambia una firma de Elaboration II: `calcularRadar(calificaciones)` pasa a `calcularRadar(calificaciones, errores)`. La nota del diagrama de clases, el escenario RE-3 y su diagrama de secuencia de Elaboration II de este módulo describen todavía la firma anterior y el filtrado de componentes por parte de la Vista; quedan pendientes de un commit de sincronización aparte, igual que en fases anteriores.

**Forma de `estado.datos.RADAR`:**

```text
{
  calificaciones: [c1, c2, ..., c56]
}
```

- Cada `ci` es un entero de 0 a 5, o `null` si no se calificó. En el estado limpio, las 56 posiciones son `null`.
- El índice de cada característica es plano, de 0 a 55, y sigue el orden fijo de la tabla de estructura de Elaboration I: desde la primera característica del primer componente de Movilización hasta la última de Gestión. Las características de cada componente ocupan estos índices:

  | Componente | Índices | Componente | Índices |
  |---|---|---|---|
  | Movilización 1 | 0 a 3 | Alineamiento 2 | 28 a 31 |
  | Movilización 2 | 4 a 7 | Motivación 1 | 32 a 35 |
  | Movilización 3 | 8 a 11 | Motivación 2 | 36 a 39 |
  | Traducción 1 | 12 a 16 | Motivación 3 | 40 a 43 |
  | Traducción 2 | 17 a 20 | Gestión 1 | 44 a 47 |
  | Traducción 3 | 21 a 23 | Gestión 2 | 48 a 51 |
  | Alineamiento 1 | 24 a 27 | Gestión 3 | 52 a 55 |

- Ejemplo con los datos de la prueba RE.1 (calificadas las posiciones 0 a 3, 12 a 16, 21 a 23 y 52 a 55, el resto `null`):

```text
{
  calificaciones: [
    0, 1, 2, 3,           // Movilización 1
    null, null, null, null,   // Movilización 2
    null, null, null, null,   // Movilización 3
    0, 0, 5, 5, 5,        // Traducción 1
    null, null, null, null,   // Traducción 2
    0, 0, 0,              // Traducción 3
    null, null, null, null,   // Alineamiento 1
    null, null, null, null,   // Alineamiento 2
    null, null, null, null,   // Motivación 1
    null, null, null, null,   // Motivación 2
    null, null, null, null,   // Motivación 3
    null, null, null, null,   // Gestión 1
    null, null, null, null,   // Gestión 2
    5, 5, 5, 5            // Gestión 3
  ]
}
```

Reglas que acompañan a esa forma:

- **Campos del formulario.** Cada calificación usa el mismo patrón `data-campo` que las demás matrices, con una ruta como `calificaciones.0` (índice plano, no fila y columna como en Análisis Estructural, porque aquí no hay matriz NxN).
- **Sin agregar ni quitar.** La estructura es fija (prueba RE.7): el módulo no tiene botones "+" ni "−", y la forma guardada nunca cambia de longitud, siempre son 56 posiciones.
- **Estado guardado anterior.** Un estado guardado sin `datos.RADAR` no se descarta: se completa con `RADAR` con 56 `null`, de modo que los datos de las otras matrices sobreviven a la actualización. Además, `"RADAR"` pasa a ser un valor válido de matriz activa. Con él, el discriminador `matriz` tiene nueve valores.

### 2. Esqueleto del proyecto

No hay un bloque de código nuevo. Se extendió el esqueleto de la Construction I del Módulo 1 (arriba, en este documento) agregando solo estos stubs, en el lugar de cada componente que corresponde y con la convención de nombres ya usada (`calcularAE`, `erroresAE`, `formularioAE`, `resultadosAE`). Cada cuerpo queda vacío con un marcador de posición.

| Dónde en el esqueleto | Stub nuevo | Tipo |
|---|---|---|
| `<main>`, después de la sección AE | `<section id="matriz-radar" class="matriz" data-matriz="RADAR" hidden>` con los cuatro contenedores `formulario`, `errores`, `resultados` y `grafico` | Contenedor HTML |
| Antes de `Validador` | `erroresRadar(datos, errores)` | Auxiliar privada |
| `MotorCalculo` | `calcularRadar(calificaciones, errores)` | Método público, con una firma ajustada respecto de Elaboration II |
| `MotorGraficos` | `dibujarRadar(resultado)` | Método público, ya definido en Elaboration II |
| `Persistencia.cargar` | Solo un comentario: devolver `RADAR` con 56 `null` si el estado guardado no lo trae | Comentario |
| Antes de `Vista` | `formularioRadar(datos)` y `resultadosRadar(resultado)` | Auxiliares privadas |

Lo que no cambia en el esqueleto: los seis objetos y sus firmas públicas, salvo las dos firmas nuevas de arriba. `Vista`, `Validador`, `Persistencia` y `Exportador` no suman ningún método público. Las auxiliares privadas son funciones sueltas del script, igual que `erroresAE` y `formularioAE`: no son métodos de los objetos y no modifican sus firmas.

### Puntos nuevos marcados [VERIFICAR] en esta fase

1. Forma de cada error de `erroresRadar`: la decisión del juez fija que son por característica, no por componente, pero no el formato de cada error. Se asume que cada error identifica la característica por su índice plano (0 a 55) y lleva un mensaje, a confirmar al implementar contra el formato de errores de los otros módulos.
2. Redondeo del `puntaje`: la decisión del juez fija que es el promedio, no cuántos decimales conserva. Se asume el promedio sin redondear en `ResultadoRadar`, y que los dos decimales de las pruebas de Elaboration I son solo la forma de comparar y de mostrar.
3. Control de la calificación en el formulario (selector, botones u otro) y mecanismo por el cual un valor fuera de rango llega al sistema (prueba RE.10): se definen en Construction II, junto con el diseño de `formularioRadar`.
4. Contenido del archivo exportado y qué se escribe para los componentes no completos (escenario RE-4): sigue abierto para Construction II.
5. Siguen abiertos los puntos de Elaboration I y II de este módulo: el título repetido de los dos componentes de Alineamiento, las calificaciones con decimales, la redacción de avisos y errores, el aspecto del radar vacío y el aviso al descartar un estado corrupto.

## Construction II — Módulo 3: Radar Estratégico

Objetivo de la fase: implementar en `index.html` la lógica y las pantallas de Radar Estratégico sobre el esqueleto de Construction I, con las reglas ya fijadas en Elaboration I (pruebas RE.1 a RE.11) y las decisiones de arquitectura de Construction I. Es la primera fase de este módulo que escribe código. Lo que ya está en las fases anteriores no se repite aquí.

### 1. Qué se implementó

| Pieza | Estado |
|---|---|
| Sección `matriz-radar` y botón "RADAR" en la navegación | Implementada. El botón lleva el título "Radar Estratégico". El botón de exportar queda habilitado en RADAR. `"RADAR"` es el noveno valor de `MATRICES`. |
| `ESTRUCTURA_RADAR` | Una sola constante con los 14 componentes en el orden de las puntas (etapa, posición, cantidad de características, los índices planos `desde` y `hasta` derivados de esas cantidades, y el `titulo` y las `afirmaciones` del profesor). De ella salen los rangos de índice, el nombre de cada punta ("Movilización 1", "Alineamiento 2"...), los textos del formulario, los mensajes de error y las hojas del Excel. La tabla de índices no está escrita a mano en ningún otro lugar del código. Se acompaña de `ETAPAS_RADAR` (nombre, descripción corta y características por componente de cada etapa; de ahí sale la cantidad de características de cada componente), de `TEXTOS_RADAR` (el título y las afirmaciones de los 14 componentes, transcritos del Excel del profesor) y de `ESCALA_RADAR` (los seis niveles). |
| `formularioRadar` | Las 56 entradas agrupadas por componente y, estas, por etapa, con el helper `entrada()` y `data-campo="calificaciones.i"` (índice plano). El encabezado de cada componente es "<nombre>: <título del profesor>" y cada fila de su tabla muestra "<n>. <afirmación del profesor>" junto a su calificación; ese n es el mismo que usan los mensajes de error ("Traducción 2, característica 3"). Antes de las etapas, un aviso de que la escala está invertida y la lista de los seis niveles. No hay ningún botón "+" ni "−". |
| `erroresRadar` y `validar('RADAR')` | Revisa solo el rango (enteros de 0 a 5, con `validarRango`) de las calificaciones que tienen valor. Empuja un mensaje por característica fuera de rango, con el nombre de su componente y su posición dentro de él ("Traducción 2, característica 3: debe ser un número entero entre 0 y 5."), y devuelve los índices planos. `validar` los devuelve en la propiedad extra `indicesInvalidos`, solo para RADAR, y no llama a `validarCamposVacios` para esta matriz. |
| `calcularRadar` | Devuelve el arreglo de 14 posiciones `{ estado, puntaje }` con los cuatro estados por componente, de forma independiente. El puntaje es el promedio sin redondear de las calificaciones del componente, con su propia cantidad de características como divisor. |
| `dibujarRadar` | SVG con la misma técnica de BCG, PEYEA y AE (`svg()` y `texto()`): 14 puntas desde arriba y en sentido horario, cinco anillos (1 a 5), el eje fijo de 0 en el centro a 5 en el borde, cada punta rotulada "<Etapa> <posición>", y un punto con su puntaje solo donde el estado es "completo". Los puntos no se unen con una línea. Los números de la escala (0 a 5) van sobre el eje que sale entre las puntas 11 y 12, donde no hay ninguna punta ni rótulo. Debajo, seis líneas cortas (no más de 65 caracteres) que recuerdan la escala invertida. |
| `resultadosRadar` | Tabla de 14 filas (Etapa, Componente, Estado, Puntaje), con el puntaje a dos decimales solo en los completos, y una nota. |
| Flujo genérico (decisión C) | `evaluarMatriz`, `actualizarMatriz` y `exportarActiva` ajustados para que un componente incompleto o inválido no bloquee el módulo. Ver el punto 2. |
| Persistencia | `estadoVacio` trae `RADAR` con 56 `null`. `cargar()` completa `datos.RADAR` con 56 `null` si un estado guardado anterior no lo trae o lo trae dañado, y conserva los demás datos. Sin cambio de versión del formato. |
| Exportador | Rama `"RADAR"` con las hojas "Datos" (con el texto de cada afirmación) y "Resultados" (con el título de cada componente). Ver el punto 3. |
| Pruebas | `tests/elaboration1.test.js`: pruebas RE.1 a RE.12, más la comprobación X.7 del Exportador de RADAR. RE.12 comprueba el marcado de la disposición en dos columnas (punto 7). |

### 2. Forma final de `ResultadoRadar` y ajuste del flujo genérico

La forma que devuelve `calcularRadar` no cambió respecto de Construction I: un arreglo de 14 posiciones `{ estado, puntaje }`, con `estado` en `"completo"`, `"incompleto"`, `"vacío"` o `"inválido"` y `puntaje: null` salvo en los completos. Lo que se precisó:

- El segundo parámetro, `errores`, conserva ese nombre para no romper la firma, pero contiene los índices planos de las características fuera de rango, no mensajes. Está documentado en un comentario sobre el método.
- Un componente con una característica fuera de rango es "inválido" aunque el resto de sus características esté sin calificar.
- `dibujarRadar` solo recorre el arreglo y dibuja un punto donde `estado` es "completo": no vuelve a decidir nada.

El flujo genérico de las otras ocho matrices da por hecho que la matriz es válida o inválida como un todo, y por eso no servía para RE.2, RE.3, RE.9 y RE.10. Se aplicaron los tres cambios puntuales de la decisión C sobre código ya existente, más el mismo ajuste en `exportarActiva`:

1. En `Validador.validar`, la llamada a `validarCamposVacios` exceptúa también a RADAR, igual que ya exceptuaba a AE con `celdasFueraDeDiagonalAE`.
2. En `evaluarMatriz`, `hayCalificacionesRadar` decide el estado "vacía"; el corte `if (!validacion.valido)` exceptúa a RADAR (nunca es "invalida", siempre sigue al cálculo); se agrega el caso `'RADAR'` y el `return` final usa siempre `validacion.errores` (en las otras ocho matrices es `[]` en ese punto, así que no cambia nada para ellas).
3. En `actualizarMatriz`, `Vista.mostrarErrores([])` pasa a `Vista.mostrarErrores(evaluacion.errores)` y se agrega `dibujarRadar` a la cadena de dibujo. `exportarActiva` hace el mismo cambio antes de llamar al Exportador.

Consecuencia: un módulo RADAR con calificaciones es siempre de estado "ok". Los errores de los componentes inválidos viajan en `evaluacion.errores` junto al resultado, y la pantalla los muestra sin dejar de calcular y dibujar los otros componentes.

### 3. Estructura del Excel exportado (cierra el escenario RE-4)

Decisión del juez: dos hojas, "Datos" y "Resultados", archivo `Mtx-RADAR.xlsx` (el nombre sale de la lógica genérica). En "Datos", 56 filas (Etapa, Componente, Característica, Afirmación, Calificación), con el texto exacto de la afirmación del profesor, la calificación vacía si no se calificó y el valor tal cual si está fuera de rango. En "Resultados", 14 filas (Etapa, Componente, Título del componente, Estado, Puntaje), con el estado real de cada componente, "Inválido" incluido, y el puntaje solo en los "Completo". Solo se agregó la rama `"RADAR"` a `hojasExportacion`; la firma de `Exportador.exportarXLSX` no cambia.

La exportación no se bloquea por componentes incompletos, vacíos o inválidos: se puede exportar en cualquier estado del módulo que no sea "vacía" (el mismo criterio que ya aplicaba `exportarActiva` de forma genérica). Esto se aparta de la redacción original del escenario RE-4 de Elaboration II ("si hay un valor fuera de rango, la Vista muestra el error y no llama al Exportador"), escrita antes de que Construction I resolviera que en este módulo la invalidez es por componente. Un commit de sincronización aparte actualiza esa frase.

### 4. Criterios aplicados donde las fases anteriores no fijaban un valor

- **Estado limpio:** `datos.RADAR` empieza con 56 `null`.
- **Control de calificación:** el helper `entrada()` ya existente, un campo de texto, sin crear un control nuevo (cierra el punto 3 de los [VERIFICAR] de Construction I). Ese helper marca el campo con `inputmode="decimal"`, no `numeric`; no se cambió para no tocar un helper compartido por las demás matrices.
- **Calificación borrada:** el formulario guarda lo que escribe el estudiante, así que una calificación borrada queda como texto vacío y no como `null`. Para todo el código ambos valen "sin calificar" (`esVacio`).
- **Forma de los errores:** un mensaje de texto por característica, en el arreglo compartido de errores, y los índices planos aparte, como valor de retorno (cierra el punto 1 de los [VERIFICAR] de Construction I).
- **Redondeo:** el puntaje no se redondea nunca internamente; los dos decimales son solo de presentación en la tabla, en el rótulo de cada punto y en el título que sale al pasar el mouse (cierra el punto 2 de los [VERIFICAR] de Construction I).
- **Calificaciones con decimales:** se rechazan (por ejemplo 2.5), como en AE, EFI, EFE y MPC. Sigue abierto el punto 3 de los [VERIFICAR] de Elaboration I de este módulo.
- **Texto de las afirmaciones y títulos de los componentes:** en una primera versión de esta fase las 56 afirmaciones no estaban en el repositorio y cada característica se rotulaba "Característica 1", "Característica 2"... Ahora el formulario muestra el título de cada componente y el texto de cada afirmación, que el juez extrajo de la hoja "radar" de `RADAR ESTRATEGICO (1).xls` y entregó transcrito. La política es la transcripción fiel: el texto es del profesor y cualquier cambio se decide con él. Solo se hicieron dos normalizaciones mecánicas, quitar las viñetas iniciales y colapsar los espacios repetidos. No se corrigió ninguna errata, tilde ni palabra, y se conservan los defectos visibles (ver "Puntos [VERIFICAR] para el profesor"). El texto vive en una sola constante, `TEXTOS_RADAR`, que `ESTRUCTURA_RADAR` incorpora; la cantidad de características de cada componente sigue saliendo de `ETAPAS_RADAR`, y una comprobación de las pruebas verifica que las dos coincidan en los 14 componentes. Los números de la primera columna de cada tabla (1, 2, 3...) no son del profesor: son la posición de la característica dentro de su componente, la misma que usan los mensajes de error.
- **Módulo vacío:** con ninguna característica calificada no se muestra ni la tabla ni el radar, solo la pista genérica de las demás matrices ("Complete o corrija los datos para ver el resultado."), y no hay ningún error. El aspecto del radar vacío seguía abierto desde Elaboration I.
- **Componente incompleto frente a vacío:** la interfaz solo los distingue por la columna "Estado" de la tabla. Ninguno tiene punto, puntaje ni error.
- **Radar:** la primera punta va arriba y el orden sigue el sentido horario; los puntos no se unen con una línea, porque unir solo los componentes completos dibujaría una figura engañosa cuando faltan puntas. Los rótulos de las puntas van a 20 unidades del borde (en vez de 12) para que el rótulo de un componente en 5.00 no toque su propio punto.
- **Textos de lectura invertida:** el aviso va en el formulario, en la nota de la tabla de resultados y en seis líneas cortas al pie del SVG (RE.6). La redacción es libre; solo el contenido es el de la prueba.
- **Persistencia:** `cargar()` considera válido un `RADAR` con exactamente 56 posiciones, cada una `null`, número o texto. No revisa el rango: un valor fuera de rango se conserva y lo marca el Validador.

### 5. Verificación

Comando: `node tests/elaboration1.test.js`. Resultado de la última corrida: 45 de 45 comprobaciones correctas (las 32 anteriores, que incluyen la X.1 contra el diagrama de clases con `calcularRadar` y `dibujarRadar`, más RE.1 a RE.12 y X.7; eran 44 de 44 antes de la prueba RE.12). La ronda correctiva del texto del profesor no agregó comprobaciones: agregó aserciones dentro de RE.7 (14 títulos, 56 afirmaciones, `afirmaciones.length` igual a `caracteristicas` en los 14 componentes y el texto exacto de la primera afirmación de Movilización 1, la quinta de Traducción 1, la primera de Alineamiento 2 y la cuarta de Gestión 3) y cambió las de X.7 para las nuevas columnas del Excel. La ronda de las dos columnas (punto 7) agregó la prueba RE.12, que lee el marcado estático de la sección RADAR y comprueba que hay un solo `.formulario`, `.errores`, `.resultados` y `.grafico`, y que `.errores` y `.grafico` están dentro de un `.panel-radar` que va después de `.formulario` y antes de `.resultados`. Con una mutación en una copia fuera del repositorio (quitar el `div class="grafico"` de la sección RADAR) RE.12 falló, con el mensaje "cantidad de .grafico: se esperaba 1 y salió 0" y 44 de 45 correctas; con una segunda mutación (sacar el `div class="grafico"` del panel, dejándolo después de él) RE.12 también falló. Antes de implementar los dos métodos nuevos, la X.1 era la única que fallaba, porque el diagrama de Elaboration II ya los declaraba. Con tres fallos introducidos a propósito en una copia fuera del repositorio (divisor fijo de 4 en el promedio, `validarCamposVacios` aplicado también a RADAR y eje del gráfico ajustado al máximo observado), fallaron 7, 5 y 1 comprobaciones, respectivamente.

Las pruebas que dependen del dibujo y del formulario (RE.4 a RE.7) no tienen navegador. Para que no queden sin verificar, el arnés usa un DOM simulado mínimo y comprueba la estructura de lo que generan `dibujarRadar` y `formularioRadar`: cuántos puntos hay, a qué fracción del radio, sobre qué punta, qué rótulos, cuántos campos y cuántos botones. No comprueba cómo se ve en pantalla. La tabla de índices de esas pruebas está transcrita de Elaboration I, sin derivarla de la aplicación, para comprobar la aplicación contra ella.

Además se recorrió el módulo en el navegador integrado del escritorio de Claude, con una copia servida desde un servidor temporal fuera del repositorio (ya detenido) y eventos reales del DOM (en la primera ronda): escribir las calificaciones de RE.1 (cuatro puntos), un componente incompleto, un valor 7 fuera de rango (mensaje "Alineamiento 2, característica 1", componente "Inválido" y los otros cuatro puntos intactos), 56 campos y ningún botón "+" ni "−", recarga con los valores y la matriz activa conservados, y volver a BCG. Los rótulos del radar caben dentro del recuadro del SVG. Sin errores de consola. No se pudo tomar una captura de pantalla (la ventana estaba minimizada) ni se pulsó "Exportar a Excel" para no descargar archivos; el contenido del archivo lo cubre X.7.

En la ronda correctiva se repitió el recorrido con el caso de RE.1 cargado y se midió el cuadro de cada rótulo de punta, de cada número de la escala, de cada punto y de cada línea del pie del SVG: ninguno se superpone con otro, todos quedan dentro del recuadro de 640 por 650 y las líneas del pie, aun con 9 píxeles por carácter, caben dentro del ancho. La primera medición encontró que el rótulo "Gestión 3" tocaba su punto en 5.00; por eso los rótulos se alejaron del borde. Es una medición de cuadros, no una captura: cómo se ve el radar sigue siendo trabajo de Construction III.

### 6. Qué queda para Construction III

- El resultado visual real de `dibujarRadar` (si los 14 rótulos se leen sin pisarse, el aspecto de los puntos y los anillos, el radar con muy pocos puntos) y el aspecto del formulario con sus 56 campos.
- La usabilidad de escribir 56 calificaciones: orden de tabulación, y si conviene un control distinto de un campo de texto. Hallazgo de la revisión del juez: con las 56 afirmaciones a la vista, la página mide unos 5600 píxeles de alto, y el mensaje de error y el gráfico quedan muy por debajo de los campos, lejos de lo que el estudiante acaba de escribir. Ver el punto 7.
- La apertura del archivo `Mtx-RADAR.xlsx` en un programa de hojas de cálculo, con componentes incompletos e inválidos.
- Navegadores distintos de Chrome y el archivo descargado de internet: nada de esto se verificó en esta fase.
- La comprobación de la disposición de dos columnas del punto 7 en el navegador real del equipo (Chrome, Edge y Firefox), incluido el selector `:has()`.

### 7. Ronda correctiva: disposición en dos columnas

Hallazgo del juez al renderizar el módulo en un navegador real: con las 56 afirmaciones a la vista, la página de RADAR medía unos 5900 píxeles de alto, y el recuadro de errores y el gráfico quedaban al final, lejos de los campos que la persona está escribiendo.

**Contrato.** Con una ventana de 75rem (1200 px) o más, dos columnas: a la izquierda el formulario y a la derecha un panel con el recuadro de errores (cuando hay) y el gráfico, fijo (`position: sticky`) en la parte alta de la ventana mientras se recorren los 56 campos, con un alto máximo igual al de la ventana menos un margen y desplazamiento interno solo como último recurso. La tabla de resultados y su nota van debajo de las dos columnas, a todo el ancho. Con menos de 75rem, una sola columna en el orden formulario, errores, gráfico y resultados, con el gráfico centrado y de no más de 36rem. El ancho máximo de `main` (64rem) pasa a 92rem solo cuando la sección RADAR está visible y solo con 75rem o más; las otras ocho secciones conservan 64rem. No cambia ningún JavaScript, dato, persistencia, exportación ni texto.

**Solución adoptada.** La referencia del juez, tal cual. En el marcado de `matriz-radar` solo cambia el orden de los bloques y aparece el contenedor `panel-radar`, que agrupa `.errores` y `.grafico`; las cuatro clases que usa el código se conservan. El código las busca con `querySelector` desde la sección (o con un selector descendiente), así que las encuentra dentro del panel. El CSS nuevo es un bloque justo después de la regla `.grafico svg`: una consulta de medios de 75rem con una cuadrícula de dos columnas (la derecha de 36rem) y el selector `main:has(#matriz-radar:not([hidden]))` para ampliar `main` a 92rem, más una regla que limita el `svg` del gráfico a 36rem y lo centra.

**Decisiones del juez**

- La tabla de resultados queda debajo, a todo el ancho, y no dentro del panel fijo. Medida en la captura, la tabla de 15 filas mide unos 470 px y el panel ya mide unos 680 px, así que juntos pasarían de los 744 px disponibles en una ventana de 768 de alto.
- La columna derecha reserva su espacio con el módulo vacío, para que el diseño no se reacomode cuando aparece el primer componente completo. Es deliberado.
- Los umbrales son 75rem para pasar a dos columnas, 92rem de ancho máximo de `main` y 36rem para el panel y el gráfico.
- Los celulares (390 px de ancho) no se atienden en esta fase: ahí el gráfico mide unos 324 px y sus rótulos quedan en unos 6 px. Es una observación para Construction III, no un defecto corregido.

**Mediciones** (navegador integrado del escritorio de Claude, sobre una copia servida desde un servidor temporal fuera del repositorio, ya detenido; con `getBoundingClientRect`; datos de RE.1 más Movilización 2 incompleta y un 7 en Alineamiento 2, característica 1). Los anchos de documento son el `clientWidth`, que descuenta la barra de desplazamiento de 15 px.

| Criterio | Medición |
|---|---|
| C1. `svg` y `ul` de errores dentro de la ventana, con la página al 25, 50 y 75 % | A 1366 por 768: `svg` en y de 76 a 661 (alto 585) y x de 742 a 1318; `ul` en y de 24 a 64. A 1280 por 720: `svg` en y de 76 a 661 y x de 656 a 1232; `ul` en y de 24 a 64. Los mismos valores en los tres puntos de desplazamiento; todo queda dentro de la ventana. El panel mide 657 px de alto, que con el margen superior de 12 px cabe en 720. |
| C2. Ancho del `svg` | 576 px a 1366 por 768 y a 1280 por 720 (al menos 560). |
| C3. Sin desborde horizontal (`scrollWidth` y `clientWidth`) | 1366 por 768: 1351 y 1351. 1280 por 720: 1265 y 1265. 1200 por 800: 1185 y 1185. 1190 por 800: 1175 y 1175. 390 por 844: 390 y 390. |
| C4. Ancho de `main` en las otras ocho secciones a 1366 por 768 | 1024 px en BCG, EFI, EFE, MPC, PEYEA, MIE, GE y AE, antes y después de visitar RADAR; igual a lo que daba `HEAD` antes de esta ronda (medido sobre una copia sin la ronda). En RADAR, `main` mide 1351 px, que es todo el ancho disponible (el máximo de 92rem no se alcanza). |
| C5. Una sola columna por debajo de 75rem | 1190 por 800: `display: block`; tope vertical del formulario 145, errores 4400, gráfico 4453, resultados 5046; `svg` de 576 px. 390 por 844: `display: block`; 304, 8796, 8871 y 9208; `svg` de 324 px. |
| C6. Orden de los campos | Los 56 `data-campo` del documento son `calificaciones.0` a `calificaciones.55`, en orden; ningún `tabindex`. Con la tecla Tab real: 10 pulsaciones desde el campo 0 llevan al 10, y 45 más al 55 (en y 695 dentro de la ventana, con la página desplazada 4265 px). Con el foco en el campo 55 el `svg` sigue en y de 76 a 661 y el `ul` de 24 a 64. |
| C7. Módulo vacío (1366 por 768) | Ningún `svg`; `.errores` y `.grafico` sin hijos; `.resultados` con solo "Complete o corrija los datos para ver el resultado."; el panel existe, con 576 px de ancho (la columna reservada) y 0 de alto; sin desborde. |
| C8. `git diff` de `index.html` | 14 líneas agregadas y 2 quitadas, en dos hunks: las líneas 40 a 49 (el bloque CSS nuevo, dentro de `<style>`, que cierra en la 84) y las 171 a 175 (el marcado de la sección, entre `<main>` en la 103 y `</main>` en la 177). Los bloques `<script id="sheetjs">` (línea 179) y `<script id="app">` (línea 211) no cambian. |
| C9. Pruebas | 45 de 45. En `tests/elaboration1.test.js` solo cambia el comentario del encabezado que cuenta las pruebas y se inserta RE.12; ninguna prueba anterior se cambió ni se eliminó. |
| C10. Consola | Sin errores ni advertencias. La consola del navegador estaba vacía tras cada carga, y una captura propia de `console.error`, `console.warn` y del evento `error` de la ventana, activa mientras se escribían calificaciones y se visitaban las nueve secciones, quedó con 0 entradas. |

Altura total de la página con esos datos a 1366 por 768: 6024 px antes de esta ronda (medida sobre una copia sin ella) y 5636 después. La página sigue siendo larga; lo que cambia es que el panel con los errores y el gráfico acompaña a la persona mientras recorre los campos. A 1280 por 720 mide 5940 px y a 1200 por 800, 6320 px, porque la columna del formulario es más angosta y las afirmaciones ocupan más líneas.

**[VERIFICAR]** El soporte del selector `:has()` en las versiones de Edge y Firefox del equipo donde se use la aplicación. Si el navegador no lo soporta, `main` conserva 64rem y la columna del formulario queda angosta (unos 360 px) pero usable. Construction III lo prueba en Edge y Firefox. Todo lo medido arriba se midió en un solo navegador, el integrado del escritorio de Claude.

### Puntos [VERIFICAR] para el profesor

Salen de transcribir fielmente el texto del Excel. Ninguno se corrigió en la aplicación ni en el Excel exportado.

1. **Título repetido de los dos componentes de Alineamiento.** Ya está registrado en la Inception de este módulo (punto 1 de los [VERIFICAR]); remitirse a él. La aplicación lo muestra tal cual: "Alineamiento 1: LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO" y "Alineamiento 2: LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO". Las afirmaciones del segundo hablan de unidades de soporte y de áreas o secciones, no de unidades de negocio.
2. **Característica repetida y fuera de lugar.** La característica 4 de Gestión 1 ("Existe un mecanismo para premiar las inciativas y las sugerencias de los colaboradores") es idéntica a la característica 4 de Motivación 3. Es fiel al Excel, pero no tiene relación con un componente sobre el presupuesto ("EL PRESUPUESTO ESTÁ ESTABLECIDO Y EXISTE UN MÉTODO DE SEGUIMIENTO"). Falta que el profesor confirme cuál es la afirmación correcta de Gestión 1.
3. **Erratas evidentes del texto transcrito**, que se conservaron sin corregir y que el profesor puede querer corregir en su archivo:
   - "inciativas" (característica 4 de Motivación 3 y característica 4 de Gestión 1).
   - "seguiimiento" (característica 3 de Gestión 2).
   - "establecidda" (característica 1 de Motivación 3).
   - "l os EE-UN" (característica 3 de Alineamiento 1).
   - "del la Estrategia" (característica 4 de Gestión 3).
   - "EL superior" (característica 2 de Motivación 2).
   - El espacio antes de la coma en "estrategicas , actividades" (característica 1 de Traducción 3).
   - Además, faltan tildes en varias palabras ("estrategicos", "actuacion", "organizacion", "formulacion", "tecnologia", "periodicas") y algunos títulos en mayúsculas llevan tilde y otros no ("ESTÁN" frente a "VISION, MISION").

## Construction III — Módulo 3: Radar Estratégico

Objetivo de la fase: dejar un plan de pruebas manual para Radar Estratégico, pensado para que una persona lo siga con el navegador real, el mouse y el teclado. Normalmente esta fase no escribe ni modifica código de la aplicación ni de las pruebas automáticas; la excepción son las tres correcciones descritas en el punto 5, que el juez decidió al verificar este plan contra la aplicación real. Cubre lo que Construction II dejó pendiente en el punto 6 de su sección: el resultado visual real de `dibujarRadar` y del formulario con sus 56 campos, la usabilidad de escribir las 56 calificaciones, la disposición en dos columnas de la ronda correctiva, el archivo `Mtx-RADAR.xlsx` abierto en un programa de hojas de cálculo, y Edge y Firefox. Lo que es propiedad del archivo completo (abre sin instalar, no pide red, funciona sin conexión, archivo descargado de internet) ya está en el bloque 5 de la Construction III del Módulo 1 y no se repite.

Estado del plan: ningún caso ha sido ejecutado por una persona, por eso todas las casillas "¿Pasó?" están en blanco. Los textos y números de "Resultado esperado" no se calcularon a mano, salvo los puntajes de los datos de prueba, que se calcularon a mano por separado y se compararon con lo medido (ver "Datos de prueba"). Se obtuvieron recorriendo los mismos pasos sobre `index.html` en el navegador integrado del escritorio de Claude (un navegador Chromium), con eventos reales del DOM (clics, escritura en los campos con teclas reales, la tecla Tab, la tecla F5, el borrado con Retroceso, el botón de exportar, la recarga de la página y scripts de consola), sobre una copia temporal servida desde un servidor fuera del repositorio, ya detenido y borrado. Los textos salen del DOM, la geometría del gráfico sale de los elementos SVG (posición de cada punto, círculo y rótulo) y el contenido del archivo exportado sale del libro que genera SheetJS, leído antes de que el navegador lo descargue. Los resultados esperados describen el comportamiento después de las tres correcciones de código del punto 5 (exportación de un valor no numérico, colocación de los números de puntaje y tamaño del panel). Eso no reemplaza la ejecución real. Lo que la simulación no pudo reproducir queda marcado en cada caso: el aspecto del mouse y el retraso del texto emergente, la sensación de comodidad al escribir, Edge, Firefox, la apertura en un programa de hojas de cálculo y la escala de pantalla de Windows.

### 0. Convenciones propias de Radar Estratégico

Valen las convenciones de la sección 0 de la Construction III del Módulo 1 (abrir la aplicación, cómo leer los resultados, el aviso ámbar y el recuadro rojo de errores, y la casilla "¿Pasó?"). Se agregan o se precisan estas:

- **Reinicio de datos (RD):** el mismo de siempre: F12, pestaña "Consola", escribir `localStorage.removeItem('mtx.estado'); location.reload()` y pulsar Enter. Después de un RD se abre en BCG.
- **Estado limpio de RADAR:** los 56 campos en blanco, sin etapas ni componentes que agregar o quitar. La sección muestra, en este orden, el texto de ayuda, el aviso de escala invertida con los seis niveles, las cinco etapas con sus 14 componentes y, por componente, una tabla de dos columnas ("Afirmación" y "Calificación (0 a 5)"). En resultados dice "Complete o corrija los datos para ver el resultado.", sin recuadro rojo y sin gráfico.
- **Estructura fija:** el módulo no tiene botones "+" ni "−", ni selectores. Las etapas, los componentes y las afirmaciones son texto fijo del profesor y no se editan. Lo único que escribe la persona son las 56 calificaciones.
- **Escala invertida:** 0 es el mejor valor (el objetivo ideal ya se cumple) y 5 el peor (máximo alejamiento). Cada calificación es un número entero de 0 a 5. Un punto del radar que se aleja del centro señala un problema, no una fortaleza.
- **Estados de un componente:** "Completo" (todas sus características calificadas: tiene puntaje y punto), "Incompleto" (algunas sí y otras no: sin puntaje, sin punto y sin error), "Vacío" (ninguna calificada: igual) e "Inválido" (alguna calificación fuera de 0 a 5, no entera o no numérica: sin puntaje, sin punto y con un mensaje de error). Cada componente se evalúa solo, sin depender de los otros trece. Los estados salen en la columna "Estado" de la tabla de resultados.
- **Nombres:** cada componente se llama "<etapa> <posición>" (por ejemplo "Movilización 1", "Alineamiento 2"). Ese nombre es el rótulo de su punta en el radar, el que usan los mensajes de error y el de la hoja "Resultados". Dos componentes de Alineamiento comparten título en el documento del profesor, así que lo que los distingue es el nombre. Las características se numeran de 1 al número de características de su componente.
- **Mensajes de error:** el recuadro rojo lista un mensaje por cada calificación no válida, con el formato "<componente>, característica <n>: debe ser un número entero entre 0 y 5." y en el orden del formulario.
- **Disposición según la ventana:** con una ventana de 75rem (1200 px) o más, el formulario va a la izquierda y, a la derecha, un panel con los errores (cuando hay) y el gráfico, que se queda a la vista mientras se recorre el formulario; la tabla de resultados queda debajo, a todo el ancho. Con menos de 75rem todo va en una sola columna: formulario, errores, gráfico y resultados. Los casos de disposición se escriben como relaciones y no como píxeles, porque la tipografía cambia entre equipos y con ella las alturas.
- **Gráfico:** radar de 14 puntas, una por componente, con la primera punta (Movilización 1) arriba y las demás en sentido horario. Eje fijo de 0 en el centro a 5 en el borde exterior, con cinco anillos (1 a 5). Las posiciones del SVG se dan en unidades del recuadro de 640 de ancho por 650 de alto, con el centro en (320, 285) y el radio exterior en 200, y también como porcentaje del ancho y del alto del recuadro. La distancia de un punto al centro es su puntaje dividido entre 5, como fracción del radio, sobre la dirección de su punta.
- **Cómo cargar los datos:** hay dos formas. Escribir en los campos con la tecla Tab (procedimiento CA-R, abajo), que es lo que hace una persona, o pegar un script en la consola (el patrón del D-AE3 del Módulo 2), que sirve para llegar rápido a un estado. Los casos dicen cuál usar.

### 1. Datos de prueba

Los arreglos de los cuatro conjuntos van por componente, en el orden de las puntas (Movilización 1, 2 y 3; Traducción 1, 2 y 3; Alineamiento 1 y 2; Motivación 1, 2 y 3; Gestión 1, 2 y 3), con "—" para una característica sin calificar.

**D-RA1: los datos de la prueba RE.1** (cuatro componentes completos, el resto vacío):

| Punta | Componente | Calificaciones | Estado y puntaje esperados |
|---|---|---|---|
| 1 | Movilización 1 | 0, 1, 2, 3 | Completo 1.50 |
| 4 | Traducción 1 | 0, 0, 5, 5, 5 | Completo 3.00 |
| 6 | Traducción 3 | 0, 0, 0 | Completo 0.00 |
| 14 | Gestión 3 | 5, 5, 5, 5 | Completo 5.00 |
| 2, 3, 5, 7 a 13 | los otros diez | — | Vacío |

**D-RA2: los 14 componentes completos**, con promedios variados:

| Punta | Componente | Calificaciones | Puntaje esperado |
|---|---|---|---|
| 1 | Movilización 1 | 1, 2, 2, 3 | 2.00 |
| 2 | Movilización 2 | 0, 1, 1, 1 | 0.75 |
| 3 | Movilización 3 | 4, 4, 5, 4 | 4.25 |
| 4 | Traducción 1 | 2, 3, 3, 2, 3 | 2.60 |
| 5 | Traducción 2 | 1, 1, 2, 1 | 1.25 |
| 6 | Traducción 3 | 2, 2, 3 | 2.33 (7 entre 3, sin redondear internamente) |
| 7 | Alineamiento 1 | 5, 5, 4, 5 | 4.75 |
| 8 | Alineamiento 2 | 0, 0, 0, 1 | 0.25 |
| 9 | Motivación 1 | 3, 3, 3, 3 | 3.00 |
| 10 | Motivación 2 | 2, 1, 2, 2 | 1.75 |
| 11 | Motivación 3 | 4, 4, 5, 5 | 4.50 |
| 12 | Gestión 1 | 1, 0, 1, 0 | 0.50 |
| 13 | Gestión 2 | 3, 2, 3, 3 | 2.75 |
| 14 | Gestión 3 | 2, 2, 1, 1 | 1.50 |

**D-RA3: estados mezclados**, con tres clases de valor no válido (un 6, un 2.5 y el texto "abc"):

| Punta | Componente | Calificaciones | Estado esperado |
|---|---|---|---|
| 1 | Movilización 1 | 1, 1, 1, 1 | Completo 1.00 |
| 2 | Movilización 2 | 2, 3, —, — | Incompleto |
| 3 | Movilización 3 | —, —, —, — | Vacío |
| 4 | Traducción 1 | 0, 1, 6, —, — | Inválido (el 6 es la característica 3) |
| 5 | Traducción 2 | 2.5, 1, 1, 1 | Inválido (el 2.5 es la característica 1) |
| 6 | Traducción 3 | —, —, — | Vacío |
| 7 | Alineamiento 1 | —, —, —, — | Vacío |
| 8 | Alineamiento 2 | —, —, —, — | Vacío |
| 9 | Motivación 1 | —, abc, —, — | Inválido (el "abc" es la característica 2) |
| 10 a 13 | Motivación 2 y 3, Gestión 1 y 2 | todo — | Vacío |
| 14 | Gestión 3 | 0, 0, 0, 0 | Completo 0.00 |

Con D-RA3 el recuadro rojo lista tres errores, en este orden: "Traducción 1, característica 3: debe ser un número entero entre 0 y 5.", "Traducción 2, característica 1: debe ser un número entero entre 0 y 5." y "Motivación 1, característica 2: debe ser un número entero entre 0 y 5."; y el gráfico tiene dos puntos: Movilización 1 al 20 % del radio y Gestión 3 en el centro.

**D-RA4: calificaciones bajas ("empresa sana")**, con los 14 componentes completos y puntajes muy cercanos al centro del radar (CP-130):

| Punta | Componente | Calificaciones | Puntaje esperado |
|---|---|---|---|
| 1 | Movilización 1 | 0, 0, 1, 0 | 0.25 |
| 2 | Movilización 2 | 0, 1, 0, 0 | 0.25 |
| 3 | Movilización 3 | 1, 0, 0, 0 | 0.25 |
| 4 | Traducción 1 | 0, 0, 0, 1, 0 | 0.20 |
| 5 | Traducción 2 | 1, 0, 0, 0 | 0.25 |
| 6 | Traducción 3 | 0, 0, 1 | 0.33 |
| 7 | Alineamiento 1 | 0, 1, 0, 0 | 0.25 |
| 8 | Alineamiento 2 | 0, 0, 0, 0 | 0.00 |
| 9 | Motivación 1 | 0, 0, 0, 1 | 0.25 |
| 10 | Motivación 2 | 1, 0, 0, 0 | 0.25 |
| 11 | Motivación 3 | 0, 0, 1, 0 | 0.25 |
| 12 | Gestión 1 | 0, 1, 0, 0 | 0.25 |
| 13 | Gestión 2 | 0, 0, 0, 0 | 0.00 |
| 14 | Gestión 3 | 1, 0, 0, 1 | 0.50 |

**Cargar un conjunto escribiendo (procedimiento CA-R), partiendo de un RD:**

1. Pulse "RADAR" en la navegación.
2. Haga clic en el primer campo de calificación (la afirmación "1." de Movilización 1).
3. Escriba la calificación de la primera característica y pulse Tab. Siga en el orden de la tabla del conjunto: los cuatro campos de Movilización 1, los cuatro de Movilización 2, y así hasta Gestión 3. Tab pasa siempre al campo siguiente del formulario; la sección no tiene otros controles entre los campos.
4. Para dejar una característica sin calificar, pulse Tab sin escribir nada. Para un valor como 2.5 o abc, escríbalo tal cual.
5. Son 56 campos: 56 calificaciones o espacios en blanco y 55 pulsaciones de Tab.

**Cargar un conjunto con un script de consola (SC), partiendo de un RD.** Pegar en la pestaña "Consola" (si el navegador pide escribir `allow pasting`, hágalo) y pulsar Enter. Se reemplaza el arreglo `porComponente` por el del conjunto:

```javascript
(() => {
  const porComponente = /* arreglo del conjunto */;
  const calificaciones = porComponente.flat().map((v) => (v === null ? null : String(v)));
  const guardado = Persistencia.cargar();
  guardado.matrizActiva = 'RADAR';
  guardado.datos.RADAR = { calificaciones };
  console.log('Guardado:', Persistencia.guardar(guardado));
  location.reload();
})();
```

Los arreglos, con `null` para una característica sin calificar:

- **D-RA1:** `[[0,1,2,3],[null,null,null,null],[null,null,null,null],[0,0,5,5,5],[null,null,null,null],[0,0,0],[null,null,null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[5,5,5,5]]`
- **D-RA2:** `[[1,2,2,3],[0,1,1,1],[4,4,5,4],[2,3,3,2,3],[1,1,2,1],[2,2,3],[5,5,4,5],[0,0,0,1],[3,3,3,3],[2,1,2,2],[4,4,5,5],[1,0,1,0],[3,2,3,3],[2,2,1,1]]`
- **D-RA3:** `[[1,1,1,1],[2,3,null,null],[null,null,null,null],[0,1,6,null,null],[2.5,1,1,1],[null,null,null],[null,null,null,null],[null,null,null,null],[null,"abc",null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[null,null,null,null],[0,0,0,0]]`
- **D-RA4:** `[[0,0,1,0],[0,1,0,0],[1,0,0,0],[0,0,0,1,0],[1,0,0,0],[0,0,1],[0,1,0,0],[0,0,0,0],[0,0,0,1],[1,0,0,0],[0,0,1,0],[0,1,0,0],[0,0,0,0],[1,0,0,1]]`

**Resultados esperados de los conjuntos.** Los puntajes los calculó el juez a mano y se compararon con lo que mostró la aplicación: coinciden los de los tres primeros conjuntos, sin ninguna diferencia (D-RA1: 1.50, 3.00, 0.00 y 5.00; D-RA2: 2.00, 0.75, 4.25, 2.60, 1.25, 2.33, 4.75, 0.25, 3.00, 1.75, 4.50, 0.50, 2.75 y 1.50; D-RA3: los estados y los tres mensajes de arriba). Los puntajes de D-RA4 también los calculó el juez a mano (0.25, 0.25, 0.25, 0.20, 0.25, 0.33, 0.25, 0.00, 0.25, 0.25, 0.25, 0.25, 0.00 y 0.50, en el orden de las puntas) y coinciden con la tabla de resultados que muestra la aplicación.

**Posición de los puntos de D-RA1 en el SVG** (unidades del recuadro y porcentaje del ancho y del alto desde la esquina superior izquierda): Movilización 1 en (320, 225), es decir 50.0 % y 34.6 %, a 60 de distancia del centro (0.30 del radio); Traducción 1 en (436.99, 258.30), 68.3 % y 39.7 %, a 120 (0.60); Traducción 3 en (320, 285), 50.0 % y 43.8 %, en el centro (0.00); Gestión 3 en (233.22, 104.81), 36.4 % y 16.1 %, a 200 (1.00).

**Posición de los puntos de D-RA2 en el SVG** (unidades del recuadro; entre paréntesis, el porcentaje del ancho y del alto; después, la fracción del radio):

| Punta | Componente | Puntaje | Posición | Fracción del radio |
|---|---|---|---|---|
| 1 | Movilización 1 | 2.00 | (320.00, 205.00), 50.0 % y 31.5 % | 0.40 |
| 2 | Movilización 2 | 0.75 | (333.02, 257.97), 52.0 % y 39.7 % | 0.15 |
| 3 | Movilización 3 | 4.25 | (452.91, 179.01), 70.8 % y 27.5 % | 0.85 |
| 4 | Traducción 1 | 2.60 | (421.39, 261.86), 65.8 % y 40.3 % | 0.52 |
| 5 | Traducción 2 | 1.25 | (368.75, 296.13), 57.6 % y 45.6 % | 0.25 |
| 6 | Traducción 3 | 2.33 | (392.97, 343.19), 61.4 % y 52.8 % | 0.4667 |
| 7 | Alineamiento 1 | 4.75 | (402.44, 456.18), 62.9 % y 70.2 % | 0.95 |
| 8 | Alineamiento 2 | 0.25 | (320.00, 295.00), 50.0 % y 45.4 % | 0.05 |
| 9 | Motivación 1 | 3.00 | (267.93, 393.12), 41.9 % y 60.5 % | 0.60 |
| 10 | Motivación 2 | 1.75 | (265.27, 328.64), 41.4 % y 50.6 % | 0.35 |
| 11 | Motivación 3 | 4.50 | (144.51, 325.05), 22.6 % y 50.0 % | 0.90 |
| 12 | Gestión 1 | 0.50 | (300.50, 280.55), 47.0 % y 43.2 % | 0.10 |
| 13 | Gestión 2 | 2.75 | (234.00, 216.42), 36.6 % y 33.3 % | 0.55 |
| 14 | Gestión 3 | 1.50 | (293.97, 230.94), 45.9 % y 35.5 % | 0.30 |

### 2. Registro de ejecución propio

Se completa antes de empezar. Es adicional al registro de la sección 0 del Módulo 1, que sigue valiendo para los datos del equipo. Aquí se anotan el navegador de cada caso y la pantalla, porque los casos de disposición dependen del ancho de la ventana, y CP-128 y CP-129 se hacen en navegadores distintos del resto.

| Dato | Valor |
|---|---|
| Persona que ejecuta | |
| Fecha | |
| Equipo y versión de Windows | |
| Resolución de la pantalla (por ejemplo 1920 por 1080) | |
| Escala de pantalla de Windows (100 %, 125 %, 150 %...) | |
| Programa de hojas de cálculo y versión (CP-122 a CP-127) | |
| Ruta del archivo `index.html` probado | |
| Resultado global | |

| Casos | Navegador, versión y tamaño de la ventana | Observaciones |
|---|---|---|
| CP-78 a CP-106 y CP-130 | | |
| CP-107 a CP-117 | | |
| CP-118 a CP-127 | | |
| CP-128 (Edge) | | |
| CP-129 (Firefox) | | |

El ancho útil de la ventana en píxeles es el que cuenta para los casos de disposición. Con una escala de Windows de 125 % o 150 %, una pantalla de 1920 por 1080 se comporta como una ventana más estrecha (unos 1536 o 1280 píxeles de ancho) y la altura disponible baja de la misma manera. Para cada caso del bloque D anote el ancho y el alto de la ventana (por ejemplo, abriendo F12 en la pestaña "Consola" y escribiendo `innerWidth + ' por ' + innerHeight`).

### 3. Casos de prueba (CP-78 a CP-130)

Los casos continúan la numeración del Módulo 2 (CP-61 a CP-77). Cada caso parte de un RD, salvo que su precondición diga otra cosa, y comprueba una sola cosa. Los casos de disposición suponen una ventana de 1366 por 768 salvo que digan otro tamaño.

| N.º | Origen | Precondición | Pasos | Resultado esperado | ¿Pasó? |
|---|---|---|---|---|---|
| CP-78 | Escenario RE-1 y estado limpio (navegación) | RD. | 1. Observe la navegación al abrir.<br>2. Pulse "RADAR". | Hay nueve botones: BCG, EFI, EFE, MPC, PEYEA, MIE, GE, AE y RADAR, y al abrir está resaltado BCG. Al pulsar RADAR queda resaltado ese botón, se ve la sección con el título "RADAR" y, debajo, el texto de ayuda: "Radar Estratégico (El Radar de la Posición Estratégica, de Kaplan y Norton). Califique de 0 a 5 su nivel de acuerdo con cada una de las 56 afirmaciones, agrupadas en 14 componentes y 5 etapas. El puntaje de cada componente y el radar salen solos cuando todas las características de ese componente están calificadas." El botón "Exportar a Excel" está habilitado. El texto emergente del botón RADAR ("Radar Estratégico") es comportamiento del navegador y no se pudo simular. | ☐ Sí<br>☐ No |
| CP-79 | Prueba RE.6 (aviso de escala invertida) | RD. RADAR visible. | 1. Lea el recuadro ámbar y la lista que está debajo, antes de las etapas. | El recuadro ámbar dice: "La escala está invertida: 0 es el mejor valor (el objetivo ideal ya se cumple, cero alejamiento) y 5 el peor (máximo alejamiento). Cuanto mayor es su acuerdo, menor debe ser el número." Debajo, la lista de los seis niveles: "0: Estoy completamente de acuerdo", "1: Estoy bastante de acuerdo", "2: Estoy algo de acuerdo", "3: No estoy muy de acuerdo", "4: No estoy casi nada de acuerdo" y "5: Estoy en completo desacuerdo". | ☐ Sí<br>☐ No |
| CP-80 | Prueba RE.7 (etapas y componentes) | RD. RADAR visible. | 1. Lea los títulos de las cinco etapas.<br>2. Lea el encabezado de cada uno de los 14 componentes. | Las cinco etapas, en este orden: "Movilización: liderazgo ejecutivo para el cambio", "Traducción: la estrategia en términos operacionales", "Alineamiento: toda la organización en torno a la estrategia", "Motivación: hacer de la estrategia el trabajo de todos" y "Gestión: la estrategia como proceso continuo". Los 14 encabezados, con el formato "<nombre>: <título>" (3, 3, 2, 3 y 3 componentes por etapa): "Movilización 1: LA VISION, MISION Y ESTRATEGIA ESTÁN CLARAMENTE DEFINIDAS"; "Movilización 2: LOS EJECUTIVOS LIDERAN EL CAMBIO ESTRATEGICO Y CREAN EQUIPO LIDER DEL PROYECTO"; "Movilización 3: LOS EJECUTIVOS COMUNICAN EL SENTIDO DE URGENCIA"; "Traducción 1: LA ESTRATEGIA ESTA EXPLICITADA A TRAVES DE UN MAPA ESTRATEGICO COMO PARTE DEL PROCESO DE PLANEAMIENTO: LOS OBJETIVOS ESTRATÉGICOS"; "Traducción 2: LOS INDICADORES SON UTILIZADOS PARA COMUNICAR LA ESTRATEGIA Y SON BALANCEADOS EN LAS PERSPECTIVAS"; "Traducción 3: LAS METAS SON ESTABLECIDAS PARA CADA INDICADOR Y LAS INICIATIVAS ESTRATEGICAS SON CLARAMENTE DEFINIDAS"; "Alineamiento 1: LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO"; "Alineamiento 2: LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO"; "Motivación 1: LA COMUNICACIÓN ES ABIERTA Y TRANSPARENTE, PARA QUE SEA FLUIDA"; "Motivación 2: LAS METAS INDIVIDUALES ESTÁN ESTABLECIDAS Y DETERMINADAS"; "Motivación 3: MEDIANTE LA REMUNERACIÓN VARIABLE, LA EMPRESA ASOCIA TALENTOS"; "Gestión 1: EL PRESUPUESTO ESTÁ ESTABLECIDO Y EXISTE UN MÉTODO DE SEGUIMIENTO"; "Gestión 2: LA EMPRESA TIENE SISTEMAS PARA SEGUIMIENTO DE LAS OPERACIONES"; "Gestión 3: LA EMPRESA REALIZA UN SEGUIMIENTO SISTEMÁTICO DE LA GESTION ESTRATÉGICA". | ☐ Sí<br>☐ No |
| CP-81 | Prueba RE.7 (56 campos y estructura de cada tabla) | RD. RADAR visible. | 1. Cuente los campos de calificación de cada componente, de arriba abajo.<br>2. Lea los encabezados de columna de una tabla. | Cada componente tiene una tabla con los encabezados "Afirmación" y "Calificación (0 a 5)", y tantas filas como características: Movilización 1, 2 y 3, 4 cada uno; Traducción 1, 2 y 3, 5, 4 y 3; Alineamiento 1 y 2, 4 cada uno; Motivación 1, 2 y 3, 4 cada uno; y Gestión 1, 2 y 3, 4 cada uno. En total 56 campos, todos en blanco. Cada fila empieza con su número ("1.", "2."...) seguido del texto de la afirmación. | ☐ Sí<br>☐ No |
| CP-82 | Prueba RE.7 (estructura fija) | RD. RADAR visible. | 1. Busque en toda la sección un botón "+", un botón "−" o un selector. | No hay ningún botón ni selector en la sección RADAR (la sección no agrega ni quita etapas, componentes ni características, a diferencia de AE). Los títulos y las afirmaciones no son campos: no se pueden editar. | ☐ Sí<br>☐ No |
| CP-83 | Prueba RE.3 (módulo vacío) | RD. RADAR visible. Ventana de 1366 por 768. | 1. Observe la zona de resultados, la de errores y el lado derecho de la página. | Resultados: "Complete o corrija los datos para ver el resultado." No hay tabla de resultados, ni recuadro rojo, ni gráfico. La columna derecha existe pero está vacía: el formulario ocupa la columna izquierda, todo lo que queda entre el borde izquierdo de la sección y el panel, y a la derecha queda un hueco de 576 px (36rem), que es el mismo ancho que ocupará el panel con el gráfico (el diseño no se reacomoda cuando aparece el primer componente completo; se comprueba en CP-87). | ☐ Sí<br>☐ No |
| CP-84 | Escenario RE-4 (exportar con el módulo vacío) | RD. RADAR visible, sin ninguna calificación. | 1. Pulse "Exportar a Excel".<br>2. Revise la zona de descargas del navegador. | Aparece el recuadro rojo con "Complete los datos antes de exportar." y no se descarga ningún archivo. | ☐ Sí<br>☐ No |
| CP-85 | Contenido fijo del profesor (afirmaciones) | RD. RADAR visible. | 1. Lea, en cada caso del texto entre comillas, la afirmación indicada. | Se leen exactamente estas (con su número y sus mayúsculas, tildes y errores tal como están): Movilización 1, 1: "1. La Estrategia está definida y formalizada por escrito". Movilización 1, 4: "4. Existe el convencimiento en el Empresario y en la Gerencia que la Gestión Estratégica es su misión principal". Traducción 1, 5: "5. La Empresa tiene definidos el despliegue de sus objetivos a los niveles inferiores de la organizacion". Traducción 3, 1: "1. Las iniciativas estrategicas , actividades y tareas a realizar están determinados". Alineamiento 1, 3: "3. Los miembros de l os EE-UN participan en la formulacion de la estrategia". Alineamiento 2, 1: "1. Los Gerentes programan reuniones periodicas para evaluar la información necesaria con sus unidades de soporte". Gestión 3, 4: "4. La empresa tiene una reunión anual de redefinición del la Estrategia". | ☐ Sí<br>☐ No |
| CP-86 | Contenido fijo del profesor (defectos conservados a propósito) | RD. RADAR visible. | 1. Lea la afirmación 4 de Motivación 3 y la 4 de Gestión 1.<br>2. Compare los encabezados de Alineamiento 1 y Alineamiento 2. | La afirmación 4 de Motivación 3 y la 4 de Gestión 1 son idénticas: "4. Existe un mecanismo para premiar las inciativas y las sugerencias de los colaboradores" (con "inciativas", sin corregir). Los encabezados de Alineamiento 1 y 2 dicen lo mismo después de los dos puntos ("LA ESTRATEGIA CORPORATIVA ES UTILIZADA PARA GUIAR LAS ESTRATEGIAS DE LAS UNIDADES DE NEGOCIO") y se distinguen solo por el nombre. Son conservaciones del texto del profesor y no fallos, y están registradas para confirmar con él. | ☐ Sí<br>☐ No |
| CP-87 | Prueba RE.1 y RE.2 (calcular con calificaciones escritas, un componente completo) | RD. RADAR visible. Ventana de 1366 por 768. | 1. Haga clic en el primer campo y escriba, con Tab entre campos: 0, 1, 2, 3 (los cuatro de Movilización 1).<br>2. Observe la tabla de resultados y el gráfico. | Sin recuadro rojo. La tabla de resultados tiene el encabezado Etapa, Componente, Estado y Puntaje y 14 filas; la primera es "Movilización", "Movilización 1", "Completo", "1.50", y las otras 13 dicen "Vacío" y sin puntaje. Debajo, la nota "Un componente se calcula y se grafica cuando todas sus características están calificadas. Los que están incompletos o vacíos todavía no muestran su punto. Recuerde que la escala está invertida: un puntaje cercano a 0 es lo ideal y uno cercano a 5 señala un problema." El gráfico aparece a la derecha con un solo punto, en la punta de Movilización 1 (arriba), a 0.30 del radio, con el número "1.50" al lado. El formulario no cambia de ancho al aparecer el gráfico. | ☐ Sí<br>☐ No |
| CP-88 | Prueba RE.1 (cuatro componentes completos) | RD. | 1. Cargue D-RA1 (con CA-R o con SC).<br>2. Observe la tabla de resultados y el gráfico. | Sin recuadro rojo. Estados de la tabla: Movilización 1 Completo 1.50, Traducción 1 Completo 3.00, Traducción 3 Completo 0.00, Gestión 3 Completo 5.00, y los otros diez componentes "Vacío" sin puntaje. El gráfico tiene exactamente cuatro puntos. | ☐ Sí<br>☐ No |
| CP-89 | Prueba RE.2 (componente incompleto) | RD. RADAR visible. | 1. Escriba las dos primeras calificaciones de Movilización 2 (los campos 5 y 6 del formulario), por ejemplo 2 y 3, y deje las otras dos en blanco.<br>2. Observe los resultados y el recuadro de errores. | Sin recuadro rojo y sin ningún aviso. En la tabla, Movilización 2 dice "Incompleto" y no tiene puntaje (no aparece un promedio de los dos valores escritos). Los otros 13 componentes dicen "Vacío". El gráfico aparece con su marco y sin ningún punto (ver CP-97). | ☐ Sí<br>☐ No |
| CP-90 | Prueba RE.10 (valor mayor que 5) | RD. RADAR visible. | 1. Escriba 7 en la primera característica de Movilización 1 (primer campo). | Recuadro rojo con una sola línea: "Movilización 1, característica 1: debe ser un número entero entre 0 y 5." En la tabla, Movilización 1 dice "Inválido" y no tiene puntaje; los otros 13 dicen "Vacío". Lo escrito se conserva en el campo. | ☐ Sí<br>☐ No |
| CP-91 | Prueba RE.10 (otras clases de valor no válido) | D-RA2 cargado (con SC). | 1. En la primera característica de Movilización 1 escriba, uno por vez y reemplazando lo que hay, -1; después 2.5; después abc. | Con cada uno de los tres valores, el mismo resultado: recuadro rojo con "Movilización 1, característica 1: debe ser un número entero entre 0 y 5.", Movilización 1 "Inválido" sin puntaje, los otros 13 componentes "Completo" con sus puntajes de D-RA2 (el gráfico tiene 13 puntos). Lo escrito se conserva en el campo. | ☐ Sí<br>☐ No |
| CP-92 | Prueba RE.10 (corregir un valor no válido) | Continuación de CP-91, con abc escrito. | 1. Reemplace abc por 1 (el valor original de D-RA2). | El recuadro rojo desaparece, Movilización 1 vuelve a "Completo" con 2.00 y el gráfico vuelve a 14 puntos. Los otros 13 puntajes no cambiaron. | ☐ Sí<br>☐ No |
| CP-93 | Prueba RE.10 (estados mezclados) | RD. | 1. Cargue D-RA3 (con SC o con CA-R).<br>2. Lea los estados, los mensajes y cuente los puntos del gráfico. | Estados: Movilización 1 Completo 1.00; Movilización 2 Incompleto; Movilización 3 Vacío; Traducción 1 Inválido; Traducción 2 Inválido; Traducción 3, Alineamiento 1 y 2 Vacío; Motivación 1 Inválido; Motivación 2 y 3, Gestión 1 y 2 Vacío; Gestión 3 Completo 0.00. El recuadro rojo lista tres líneas, en este orden: "Traducción 1, característica 3: debe ser un número entero entre 0 y 5.", "Traducción 2, característica 1: debe ser un número entero entre 0 y 5." y "Motivación 1, característica 2: debe ser un número entero entre 0 y 5." El gráfico tiene dos puntos: Movilización 1 a 0.20 del radio (con el número "1.00") y Gestión 3 en el centro (con "0.00"). Los campos con 6, 2.5 y abc conservan lo escrito. | ☐ Sí<br>☐ No |
| CP-94 | Prueba RE.1 (catorce componentes completos y dos decimales) | RD. | 1. Cargue D-RA2 (con SC).<br>2. Lea la columna Puntaje de la tabla de resultados. | Sin recuadro rojo. Los 14 componentes dicen "Completo" y los puntajes, en el orden de las puntas, son: 2.00, 0.75, 4.25, 2.60, 1.25, 2.33, 4.75, 0.25, 3.00, 1.75, 4.50, 0.50, 2.75 y 1.50, todos con dos decimales. Traducción 3 muestra 2.33 (es 7 entre 3 y la aplicación no lo redondea internamente). El gráfico tiene 14 puntos. | ☐ Sí<br>☐ No |
| CP-95 | Prueba RE.9 (de completo a incompleto y de vuelta) | RD. Movilización 1 con 0, 1, 2, 3 (CP-87). | 1. Haga clic en el cuarto campo de Movilización 1, seleccione su contenido y pulse Retroceso.<br>2. Observe resultados, errores y gráfico.<br>3. Escriba 3 otra vez. | Paso 2: Movilización 1 pasa a "Incompleto", sin puntaje, sin recuadro rojo; el gráfico queda sin ningún punto y las otras tres calificaciones del componente se conservan. Paso 3: vuelve a "Completo" con 1.50 y reaparece el punto. | ☐ Sí<br>☐ No |
| CP-96 | Prueba RE.3 (borrar todo vuelve al módulo vacío) | Continuación de CP-95: Movilización 1 con 0, 1, 2, 3. | 1. Borre los cuatro campos de Movilización 1, uno por uno, con Retroceso.<br>2. Observe resultados y gráfico después del último. | Con tres campos borrados, Movilización 1 es "Incompleto" y el gráfico sigue visible, sin puntos. Después de borrar el último, la zona de resultados dice "Complete o corrija los datos para ver el resultado.", sin tabla, sin recuadro rojo y sin gráfico (el módulo vuelve al estado de CP-83). | ☐ Sí<br>☐ No |
| CP-97 | Prueba RE.5 (marco del radar sin ningún punto) | RD. Movilización 1 con una sola calificación escrita, por ejemplo un 2 en el primer campo. | 1. Observe el gráfico. | El gráfico aparece con su marco completo y sin ningún punto ni número de puntaje: 14 puntas rotuladas, cinco anillos concéntricos, los números 0 a 5 sobre el eje de la izquierda y las seis líneas de texto del pie. En la tabla, Movilización 1 dice "Incompleto". | ☐ Sí<br>☐ No |
| CP-98 | Prueba RE.5 (los 14 rótulos y su posición) | D-RA2 cargado (con SC). | 1. Lea los rótulos alrededor del radar, empezando por el de arriba y siguiendo el sentido de las agujas del reloj. | Los 14 rótulos, en este orden: Movilización 1 (arriba, centrado), Movilización 2, Movilización 3, Traducción 1 (a la derecha, un poco por encima del centro), Traducción 2 (a la derecha, un poco por debajo), Traducción 3, Alineamiento 1, Alineamiento 2 (abajo, centrado), Motivación 1, Motivación 2, Motivación 3 (a la izquierda, un poco por debajo del centro), Gestión 1 (a la izquierda, un poco por encima), Gestión 2 y Gestión 3. Posiciones del texto (unidades del recuadro): Movilización 1 (320, 69), Movilización 2 (415, 91), Movilización 3 (492, 152), Traducción 1 (534, 240), Traducción 2 (534, 338), Traducción 3 (492, 426), Alineamiento 1 (415, 487), Alineamiento 2 (320, 509), Motivación 1 (225, 487), Motivación 2 (148, 426), Motivación 3 (106, 338), Gestión 1 (106, 240), Gestión 2 (148, 152) y Gestión 3 (225, 91). Los de la derecha empiezan en ese punto, los de la izquierda terminan en él y los de arriba y abajo van centrados. Hay 14 líneas finas que salen del centro, una por rótulo. | ☐ Sí<br>☐ No |
| CP-99 | Prueba RE.4 (anillos y números de la escala) | D-RA2 cargado (con SC). | 1. Observe los círculos y los números 0 a 5. | Hay cinco círculos concéntricos centrados en (320, 285), de radios 40, 80, 120, 160 y 200 (el último, el borde exterior, con un trazo más marcado). Los números 0, 1, 2, 3, 4 y 5 van sobre una línea horizontal a la izquierda del centro, a 289 de altura, entre los rótulos de Motivación 3 (abajo) y Gestión 1 (arriba): el 0 en x 310, el 1 en 277, el 2 en 237, el 3 en 197, el 4 en 157 y el 5 en 117, cada uno justo a la izquierda del círculo que le corresponde. Ningún número toca un rótulo de punta. | ☐ Sí<br>☐ No |
| CP-100 | Prueba RE.4 (posición de los puntos de D-RA1) | D-RA1 cargado (con SC). | 1. Ubique cada uno de los cuatro puntos. | Los cuatro puntos están sobre la punta de su componente y a esta distancia del centro: Movilización 1 a 0.30 del radio (50.0 % y 34.6 % del recuadro), Traducción 1 a 0.60 (68.3 % y 39.7 %), Traducción 3 en el centro (50.0 % y 43.8 %) y Gestión 3 a 1.00 (36.4 % y 16.1 %). Cada punto lleva a su lado su puntaje con dos decimales: 1.50, 3.00, 0.00 y 5.00. | ☐ Sí<br>☐ No |
| CP-101 | Prueba RE.4 (posición de los puntos de D-RA2) | D-RA2 cargado (con SC). | 1. Ubique los 14 puntos. | Hay 14 puntos, cada uno sobre la punta de su componente, a la fracción del radio y en la posición de la tabla de "Posición de los puntos de D-RA2" de la sección 1 (por ejemplo, Movilización 3 a 0.85, en 70.8 % y 27.5 % del recuadro, y Alineamiento 2 a 0.05, casi en el centro). Los de mayor alejamiento son Alineamiento 1 (4.75) y Motivación 3 (4.50); los más cercanos al centro, Alineamiento 2 (0.25) y Gestión 1 (0.50). | ☐ Sí<br>☐ No |
| CP-102 | Prueba RE.4 (los dos extremos de la escala) | D-RA1 cargado (con SC). | 1. Mire el punto de Traducción 3 (0.00).<br>2. Mire el punto de Gestión 3 (5.00) y el rótulo "Gestión 3". | El punto de Traducción 3 está exactamente en el centro del radar. El punto de Gestión 3 está sobre el círculo exterior (el de radio 200), en (233, 105), y no toca el rótulo "Gestión 3", que termina en (225, 91), ni su propio número "5.00". | ☐ Sí<br>☐ No |
| CP-103 | Escenario RE-1 (texto emergente de cada punto) | D-RA2 cargado (con SC). | 1. Deje el cursor quieto sobre cada punto un par de segundos, al menos sobre los de Movilización 1, Traducción 3 y Gestión 3. | Aparece un texto emergente con el nombre del componente y su puntaje: "Movilización 1: 2.00", "Traducción 3: 2.33" y "Gestión 3: 1.50" (y, en general, "<componente>: <puntaje con dos decimales>"). Los textos están en el gráfico; que el navegador los muestre, y cuánto tarda, es comportamiento del navegador y no se pudo simular. | ☐ Sí<br>☐ No |
| CP-104 | Prueba RE.6 (pie del gráfico) | D-RA1 cargado (con SC). | 1. Lea el texto que está debajo del radar. | Seis líneas centradas, todas dentro del recuadro y ninguna cortada: "Escala invertida: 0 en el centro (el objetivo ideal ya se cumple)", "y 5 en el borde exterior (máximo alejamiento).", "Un punto que se aleja del centro señala un problema", "en ese componente, no una fortaleza.", "Solo se grafican los componentes con todas sus" y "características calificadas." (a 535, 553, 571, 589, 607 y 625 de altura en el recuadro de 650). | ☐ Sí<br>☐ No |
| CP-105 | Prueba RE.5 (los puntos no se unen) | D-RA2 cargado (con SC). | 1. Mire el gráfico buscando una línea o un polígono que una los puntos. | No hay ninguna línea ni área que una los 14 puntos entre sí. Las únicas líneas son las 14 que salen del centro hacia los rótulos y los cinco círculos. | ☐ Sí<br>☐ No |
| CP-106 | Prueba RE.13 (los números de puntaje no se pisan) | D-RA2 cargado (con SC). | 1. Mire los 14 números de puntaje que acompañan a los puntos, sobre todo los de la zona alrededor del centro del radar, donde están los puntos de menor puntaje (0.25, 0.50 y 0.75).<br>2. Mire el "0" de la escala (a la izquierda del centro). | Se dibujan los 14 números, uno junto a cada punto, y cada uno queda en una de las ocho posiciones alrededor de su punto (arriba a la derecha si cabe y, si no, abajo a la derecha, arriba a la izquierda, abajo a la izquierda, a la derecha, a la izquierda, arriba o abajo). Ningún número toca otro número, un rótulo de punta, un número de la escala ni un punto. La única superposición que queda es el "0" de la escala, que queda parcialmente bajo el punto de Gestión 1 (0.50): es inherente al eje fijo de 0 a 5, y no se corrige. Medido en el SVG, con el cuadro real de cada texto: 14 números dibujados de 14, 0 pares de textos que se tocan, 0 pares número-punto y 1 par dígito de la escala-punto. Antes de la corrección había 2, 2 y 1 respectivamente. | ☐ Sí<br>☐ No |
| CP-107 | Disposición: dos columnas a 1366 por 768 | D-RA2 cargado (con SC). Ventana de 1366 por 768, página arriba. | 1. Observe dónde están el formulario, el panel con el gráfico y la tabla de resultados. | El formulario ocupa la columna izquierda y el gráfico queda a la derecha, en un panel de 576 px de ancho (36rem); la tabla de resultados queda debajo de las dos columnas, a todo el ancho de la sección, después del último componente. No hay barra de desplazamiento horizontal. La sección ocupa todo el ancho de la ventana menos la barra de desplazamiento (el máximo es de 92rem, que a este ancho no se alcanza). | ☐ Sí<br>☐ No |
| CP-108 | Disposición: panel fijo a 1366 por 768 | Igual que CP-107. | 1. Baje la página hasta el 25 % de su recorrido, después hasta el 50 % y hasta el 75 %.<br>2. En cada posición observe el gráfico y los componentes del formulario a su izquierda. | En las tres posiciones el gráfico completo (su recuadro, de unos 585 px de alto) queda visible dentro de la ventana, pegado arriba a la derecha, mientras el formulario se desplaza a su izquierda. El gráfico mide 576 px de ancho. | ☐ Sí<br>☐ No |
| CP-109 | Disposición: el panel con errores cabe entero en la ventana | D-RA3 cargado (con SC). Ventana de 1280 por 720. | 1. Baje la página hasta el 25 %, 50 % y 75 % de su recorrido.<br>2. En cada posición observe el panel de la derecha: el recuadro rojo y el gráfico.<br>3. Repita los pasos 1 y 2 con una ventana más baja, de 1366 por 650 (la de un portátil de 1366 por 768, que deja unos 650 px útiles). | En todas las posiciones y en las dos ventanas, el panel completo (el recuadro rojo con los tres errores y el gráfico) queda dentro de la ventana, pegado arriba a la derecha, sin que ninguna de sus partes quede cortada por el borde superior ni por el inferior. Los tres errores se leen enteros, sin barra de desplazamiento propia. Con errores el gráfico se escala hacia abajo, conservando su proporción, solo lo que hace falta para caber: es más chico que el que se ve sin errores (CP-108) y más chico en la ventana más baja. Al principio de la página (arriba de todo) el panel todavía no está fijo y su parte de abajo puede quedar cortada por el borde inferior de la ventana: es normal y se corrige solo al empezar a bajar. Medido a 1280 por 720 y a 1366 por 650, con la página desplazada 2000 px: el panel completo dentro de la ventana en las dos. | ☐ Sí<br>☐ No |
| CP-110 | Disposición: con muchos errores el recuadro rojo se desplaza por dentro | D-RA3 cargado (con SC). Ventana de 1280 por 720. | 1. Escriba un 9 en los primeros 12 campos de calificación (los de Movilización 1, 2 y 3).<br>2. Observe el panel de la derecha.<br>3. Desplace el contenido del recuadro rojo con la rueda del mouse puesta sobre él.<br>4. Baje la página hasta el 25 %, 50 % y 75 % de su recorrido. | El recuadro rojo lista 15 errores (los 12 de las calificaciones escritas más los 3 que ya traía D-RA3). No crece sin límite: ocupa como máximo cerca de un tercio del alto de la ventana (el 30 %), muestra una barra de desplazamiento propia y con ella se pueden leer los 15 mensajes. El gráfico entero sigue a la vista debajo del recuadro, escalado hacia abajo si hace falta, y el panel completo queda dentro de la ventana en las tres posiciones de la página. La página no se desplaza por culpa de los errores. Medido: con 15 errores, el contenido del recuadro rojo es más alto que el recuadro (se desplaza por dentro) en las ventanas de 1366 por 650, 1280 por 720, 1366 por 768 y 1920 por 1080, y el panel entero queda dentro de la ventana en todas. | ☐ Sí<br>☐ No |
| CP-111 | Disposición: dos columnas a 1920 por 1080 | D-RA2 cargado (con SC). Ventana de 1920 por 1080. | 1. Observe la sección.<br>2. Baje la página hasta el 25 %, 50 % y 75 % de su recorrido. | La sección se centra en la ventana con un ancho de 1472 px (92rem), con márgenes vacíos a los lados. El formulario queda a la izquierda, ocupando lo que queda entre el borde izquierdo de la sección y el panel de 576 px, que va a la derecha. En las tres posiciones el gráfico completo queda visible. No hay barra de desplazamiento horizontal. | ☐ Sí<br>☐ No |
| CP-112 | Disposición: una columna por debajo de 75rem | D-RA3 cargado (con SC). Ventana de 1190 por 800. | 1. Observe la sección de arriba abajo, sin saltarse nada. | Una sola columna, en este orden, de arriba abajo: el formulario con sus 56 campos, el recuadro rojo con los tres errores, el gráfico (centrado y de 576 px de ancho como máximo) y la tabla de resultados. El panel no queda fijo: se desplaza con la página. No hay barra de desplazamiento horizontal. El ancho de la sección es de 1024 px. | ☐ Sí<br>☐ No |
| CP-113 | Disposición: las otras ocho secciones conservan su ancho | RD. Ventana de 1366 por 768. | 1. Pulse BCG, EFI, EFE, MPC, PEYEA, MIE, GE y AE, uno por uno, y anote el ancho de cada sección (en la consola: `document.querySelector('main').getBoundingClientRect().width`).<br>2. Pulse RADAR y vuelva a hacer lo mismo con las otras ocho. | En los dos recorridos las ocho secciones miden 1024 px de ancho (64rem), antes y después de visitar RADAR. Solo RADAR se ensancha: ocupa todo el ancho de la ventana menos la barra de desplazamiento (hasta un máximo de 92rem). | ☐ Sí<br>☐ No |
| CP-114 | Prueba RE.10 y disposición (los errores quedan junto al gráfico) | D-RA2 cargado (con SC). RADAR visible. Ventana de 1366 por 768. | 1. Baje la página hasta el último componente (Gestión 3).<br>2. Haga clic en la última calificación (la cuarta de Gestión 3), seleccione su contenido y escriba 7.<br>3. Mire el lado derecho sin mover la página. | Sin mover la página aparece, arriba a la derecha, el recuadro rojo con "Gestión 3, característica 4: debe ser un número entero entre 0 y 5." y, debajo, el gráfico con 13 puntos (sin el de Gestión 3). El campo que se acaba de escribir y el recuadro rojo se ven a la vez en la ventana. Al reemplazar el 7 por 1, el recuadro desaparece y el gráfico vuelve a 14 puntos. | ☐ Sí<br>☐ No |
| CP-115 | Escenario RE-1 (el orden de la tecla Tab) | RD. RADAR visible. | 1. Haga clic en el primer campo de calificación.<br>2. Pulse Tab 10 veces y anote en qué afirmación está el cursor.<br>3. Pulse Tab 45 veces más y anote dónde está.<br>4. Pulse Tab una vez más. | Tras 10 pulsaciones el cursor está en la afirmación 3 de Movilización 3 (el campo 11 del formulario). Tras 45 más, en la afirmación 4 de Gestión 3 (el campo 56, el último). Con una pulsación más el cursor sale del último campo: no hay ningún otro control después, dentro de la sección. Tab recorre los 56 campos de arriba abajo, sin saltarse ninguno ni entrar en el panel de la derecha. | ☐ Sí<br>☐ No |
| CP-116 | Escenario RE-1 (el punto aparece apenas se completa el componente) | RD. RADAR visible. Ventana de 1366 por 768 o mayor. | 1. Cargue D-RA2 escribiendo (CA-R).<br>2. Después de escribir la última calificación de cada componente (y antes de pulsar Tab) mire el gráfico. | Mientras faltan calificaciones de un componente, este dice "Incompleto" en la tabla y su punto no está. En el instante en que se escribe la última de sus calificaciones, el componente pasa a "Completo", sin recuadro rojo, y aparece su punto. Tras terminar cada componente hay 1, 2, 3... hasta 14 puntos, con los puntajes de D-RA2. Anote cuánto tiempo y cuántas pulsaciones cuesta (56 calificaciones y 55 pulsaciones de Tab). | ☐ Sí<br>☐ No |
| CP-117 | Requisito de usabilidad (experiencia de escribir las 56 calificaciones) | RD. Ventana de 1366 por 768 o mayor. | 1. Escriba los 56 valores de D-RA2 con Tab (CA-R), como lo haría un estudiante.<br>2. Anote la experiencia: cuánto tarda, si pierde el lugar, si el gráfico y los errores quedan a la vista, si los campos se leen bien, si escribir 0, 1... 5 en un campo de texto es cómodo.<br>3. Anote si le serviría otro control para calificar (un selector con los seis niveles, o seis botones, en lugar de un campo de texto). | No se fija un tiempo como criterio de aprobación, porque ese umbral nunca se acordó con el profesor. Se marca "Sí" si la persona pudo escribir las 56 calificaciones sin perder el lugar y sin que el módulo falle, y se anotan sus observaciones. En la simulación se escribieron los 56 valores en orden, con un evento de entrada por campo, sin errores y sin recuadro rojo en ningún momento; cómo se siente no se pudo simular. La pregunta de otro control para las calificaciones está abierta para el profesor (punto 2 de los [VERIFICAR] de esta fase). | ☐ Sí<br>☐ No |
| CP-118 | Prueba RE.11 (recarga conserva calificaciones, errores y sección) | D-RA3 cargado (con SC). RADAR visible. | 1. Pulse F5.<br>2. Revise la sección, los campos con 6, 2.5 y abc, los errores, los estados y el gráfico. | Después de F5 sigue abierta la sección RADAR, con el botón RADAR resaltado. Los campos conservan lo escrito (6, 2.5 y abc, y los demás valores de D-RA3). El recuadro rojo muestra los mismos tres errores, en el mismo orden. Los estados son los de CP-93 y el gráfico tiene los mismos dos puntos. | ☐ Sí<br>☐ No |
| CP-119 | Prueba RE.11 (ir a otra sección y volver) | Continuación de CP-118. | 1. Pulse BCG.<br>2. Pulse RADAR otra vez. | Al pulsar BCG se ve la sección BCG sin cambios. Al volver a RADAR se ven los mismos valores, los mismos tres errores, los mismos estados y los dos puntos de CP-93. | ☐ Sí<br>☐ No |
| CP-120 | Prueba RE.11 (reinicio de datos) | D-RA3 cargado (con SC). | 1. Haga un RD.<br>2. Pulse "RADAR". | Después del RD se abre BCG, y RADAR vuelve al estado limpio de CP-83: los 56 campos en blanco, la pista de resultados, sin recuadro rojo y sin gráfico. | ☐ Sí<br>☐ No |
| CP-121 | Prueba RE.11 (estado guardado dañado) | D-RA3 cargado (con SC), con un valor en BCG: escriba 500 en los ingresos de la primera división de BCG. | 1. En la consola escriba y ejecute: `(() => { const g = Persistencia.cargar(); g.matrizActiva = 'RADAR'; g.datos.RADAR = { calificaciones: ['1', '2'] }; localStorage.setItem('mtx.estado', JSON.stringify(g)); location.reload(); })();`<br>2. Mire RADAR y después BCG. | Tras la recarga se abre RADAR con los 56 campos en blanco, "Complete o corrija los datos para ver el resultado.", sin recuadro rojo ni gráfico (el estado dañado se reemplaza por uno vacío, sin aviso). Al pulsar BCG, el valor 500 de la primera división sigue ahí: los datos de los demás módulos se conservan. | ☐ Sí<br>☐ No |
| CP-122 | Escenario RE-4 (exportación: nombre y hojas) | D-RA2 cargado (con SC). RADAR activa. Un programa de hojas de cálculo (Excel, LibreOffice u otro). | 1. Pulse "Exportar a Excel".<br>2. Abra el archivo descargado con el programa de hojas de cálculo.<br>3. Mire los nombres de las hojas. | Se descarga `Mtx-RADAR.xlsx` y se abre sin errores, sin recuadro rojo en la página. Tiene dos hojas, "Datos" y "Resultados", en este orden. | ☐ Sí<br>☐ No |
| CP-123 | Escenario RE-4 (exportación: hoja Datos) | Continuación de CP-122. | 1. En la hoja "Datos", lea la fila 1, las filas 2, 3 y 14, y la última fila.<br>2. Compruebe cuántas filas y columnas hay. | La hoja ocupa de A1 a E57: la fila 1 más 56 filas, cinco columnas. Fila 1: Etapa, Componente, Característica, Afirmación, Calificación. Fila 2: Movilización, Movilización 1, Característica 1, "La Estrategia está definida y formalizada por escrito", 1. Fila 3: Movilización, Movilización 1, Característica 2, "Existe alto conocimiento de la Misión y Visión por parte del Empresario y de los niveles Ejecutivos", 2. Fila 14: Traducción, Traducción 1, Característica 1, "La Empresa tiene definidas las áreas de trabajo", 2. Fila 57: Gestión, Gestión 3, Característica 4, "La empresa tiene una reunión anual de redefinición del la Estrategia", 1. La columna E (Calificación) tiene, de la fila 2 a la 57, estos valores: 1, 2, 2, 3, 0, 1, 1, 1, 4, 4, 5, 4, 2, 3, 3, 2, 3, 1, 1, 2, 1, 2, 2, 3, 5, 5, 4, 5, 0, 0, 0, 1, 3, 3, 3, 3, 2, 1, 2, 2, 4, 4, 5, 5, 1, 0, 1, 0, 3, 2, 3, 3, 2, 2, 1, 1. | ☐ Sí<br>☐ No |
| CP-124 | Escenario RE-4 (exportación: hoja Resultados) | Continuación de CP-122. | 1. En la hoja "Resultados", lea la fila 1, las filas 2 y 7 y las 14 filas de componentes. | La hoja ocupa de A1 a E15: la fila 1 más 14 filas, cinco columnas. Fila 1: Etapa, Componente, Título del componente, Estado, Puntaje. Fila 2: Movilización, Movilización 1, "LA VISION, MISION Y ESTRATEGIA ESTÁN CLARAMENTE DEFINIDAS", Completo, 2. Fila 7: Traducción, Traducción 3, "LAS METAS SON ESTABLECIDAS PARA CADA INDICADOR Y LAS INICIATIVAS ESTRATEGICAS SON CLARAMENTE DEFINIDAS", Completo, 2.3333333333333335. Las 14 filas dicen "Completo" y los puntajes (columna E, de la fila 2 a la 15) son: 2, 0.75, 4.25, 2.6, 1.25, 2.3333333333333335, 4.75, 0.25, 3, 1.75, 4.5, 0.5, 2.75 y 1.5. | ☐ Sí<br>☐ No |
| CP-125 | Escenario RE-4 (exportación: el puntaje es numérico y sin redondear) | Continuación de CP-122, con la hoja "Resultados" abierta. | 1. En una celda libre escriba una fórmula que compruebe si es número, sobre las celdas de puntaje (por ejemplo `=ESNUMERO(E7)` en Excel en español; en inglés `=ISNUMBER(E7)`).<br>2. En otra celda escriba `=E7*3`.<br>3. Haga lo mismo con `=ESNUMERO(E2)`. | La primera fórmula da VERDADERO: el puntaje es una celda numérica, no texto. `=E7*3` da 7: Traducción 3 guarda 7 entre 3 sin redondear (2.3333333333333335), aunque la celda se muestre con menos decimales según el formato. `=ESNUMERO(E2)` también da VERDADERO. Los textos (etapa, componente, título y estado) no son números. | ☐ Sí<br>☐ No |
| CP-126 | Escenario RE-4 (exportación con estados mezclados: hoja Datos) | D-RA3 cargado (con SC). RADAR activa. Un programa de hojas de cálculo. | 1. Pulse "Exportar a Excel".<br>2. En la hoja "Datos" del archivo descargado lea las celdas de calificación (columna E) de las filas 14 a 22 y la de la fila 35. | La exportación no se bloquea: se descarga `Mtx-RADAR.xlsx` aunque haya componentes inválidos, y en la página no cambia nada (siguen los tres errores). En la columna E: E14 = 0, E15 = 1, E16 = 6 (el valor fuera de rango tal cual), E17 y E18 vacías, E19 = 2.5 (tal cual), E20 a E22 = 1. La calificación abc (fila 35, Motivación 1, característica 2) se exporta como el texto "abc": E35 es una celda de texto, no un error de hoja de cálculo (en Excel no aparece #¡VALOR!). Las calificaciones sin valor (por ejemplo la fila 34) quedan vacías. | ☐ Sí<br>☐ No |
| CP-127 | Escenario RE-4 (exportación con estados mezclados: hoja Resultados) | Continuación de CP-126. | 1. En la hoja "Resultados" lea las columnas Componente, Estado y Puntaje de las 14 filas. | Estado y puntaje de cada componente: Movilización 1 Completo 1; Movilización 2 Incompleto; Movilización 3 Vacío; Traducción 1 Inválido; Traducción 2 Inválido; Traducción 3, Alineamiento 1, Alineamiento 2 Vacío; Motivación 1 Inválido; Motivación 2, Motivación 3, Gestión 1 y Gestión 2 Vacío; Gestión 3 Completo 0. Motivación 1, cuya calificación abc es no numérica, figura "Inválido", como en la página. El puntaje solo tiene valor en las filas "Completo" (la de Movilización 1 y la de Gestión 3); en las demás la celda está vacía. La columna C trae el título del profesor de cada componente (por ejemplo, el de la fila 5, Traducción 1: "LA ESTRATEGIA ESTA EXPLICITADA A TRAVES DE UN MAPA ESTRATEGICO COMO PARTE DEL PROCESO DE PLANEAMIENTO: LOS OBJETIVOS ESTRATÉGICOS"). | ☐ Sí<br>☐ No |
| CP-128 | Historia 7 y módulo RADAR: Edge | Windows con Edge. Misma copia local de `index.html`. RD en Edge. Ventana de al menos 1366 px de ancho. | 1. Abra `index.html` con Edge.<br>2. Cargue D-RA2 con SC (o con CA-R).<br>3. Repita lo de CP-94 (14 completos y puntajes), CP-98 (los 14 rótulos) y CP-101 (los 14 puntos).<br>4. Repita lo de CP-103 (texto emergente de un punto).<br>5. Repita lo de CP-107 y CP-108 (dos columnas y panel fijo). | Se cumple lo esperado en CP-94, CP-98, CP-101, CP-103, CP-107 y CP-108: mismos puntajes, el gráfico con sus 14 puntos y 14 rótulos, el texto emergente con el nombre y el puntaje, y la disposición de dos columnas con el gráfico fijo a la derecha mientras se baja la página. La disposición de dos columnas depende del selector `:has()` y de `position: sticky`: si Edge no los soporta, la sección mantiene el ancho de 64rem y la columna del formulario queda angosta pero usable. Sin mensajes de error en la consola. Anote el navegador y la versión en el registro. | ☐ Sí<br>☐ No |
| CP-129 | Historia 7 y módulo RADAR: Firefox | Windows con Firefox. Misma copia local de `index.html`. RD en Firefox. Ventana de al menos 1366 px de ancho. | 1. Abra `index.html` con Firefox.<br>2. Cargue D-RA2 con SC (o con CA-R).<br>3. Repita lo de CP-94, CP-98, CP-101, CP-103, CP-107 y CP-108. | Se cumple lo esperado en CP-94, CP-98, CP-101, CP-103, CP-107 y CP-108, igual que en CP-128. Anote el navegador y la versión en el registro. | ☐ Sí<br>☐ No |
| CP-130 | Prueba RE.13 (puntos muy juntos, "empresa sana") | RD. | 1. Cargue D-RA4 (con SC).<br>2. Mire la zona del centro del radar y cuente los números de puntaje que se dibujan.<br>3. Deje el cursor sobre un par de puntos del centro (por ejemplo, los de Alineamiento 2 y Movilización 1) y compare con la tabla de resultados. | Los 14 puntos están dibujados, todos muy cerca del centro y muchos tocándose entre sí (a un puntaje de 0.25 corresponde un 5 % del radio). Cada uno conserva su texto emergente "<componente>: <puntaje>": por ejemplo "Movilización 1: 0.25", "Traducción 1: 0.20", "Traducción 3: 0.33", "Alineamiento 2: 0.00", "Gestión 2: 0.00" y "Gestión 3: 0.50"; los dos de 0.00 están exactamente en el centro. La tabla de resultados trae los 14 puntajes, en el orden de las puntas: 0.25, 0.25, 0.25, 0.20, 0.25, 0.33, 0.25, 0.00, 0.25, 0.25, 0.25, 0.25, 0.00 y 0.50. De los 14 números solo se dibujan los que caben sin tocar otro texto ni un punto; en la medición se dibujan 4 (Gestión 3, 0.50; Traducción 3, 0.33; Motivación 2, 0.25; y Alineamiento 1, 0.25) y los otros 10 no se dibujan (su puntaje queda en el texto emergente y en la tabla). Esta colocación no depende de la tipografía del equipo: usa un ancho estimado por carácter, así que el número de números dibujados es el mismo en todos los navegadores. Ninguno de los números dibujados toca otro texto ni un punto (0 pares). Observación: los propios puntos se tocan entre sí (en la medición, 46 pares de puntos) y 4 dígitos de la escala quedan parcialmente bajo un punto; es inherente al eje fijo de 0 a 5 con puntajes bajos y no se corrige. | ☐ Sí<br>☐ No |

### 4. Resumen de cobertura

| Bloque | Casos | Rango |
|---|---|---|
| A. Estado limpio y contenido del formulario | 9 | CP-78 a CP-86 |
| B. Estados por componente y cálculo | 10 | CP-87 a CP-96 |
| C. Gráfico de `dibujarRadar` | 11 | CP-97 a CP-106 y CP-130 |
| D. Disposición y usabilidad | 11 | CP-107 a CP-117 |
| E. Recarga y persistencia | 4 | CP-118 a CP-121 |
| F. Exportación | 6 | CP-122 a CP-127 |
| G. Edge y Firefox | 2 | CP-128 y CP-129 |
| Total | 53 | CP-78 a CP-130 |

Lo que Construction II dejó pendiente en el punto 6 de su sección y dónde se cubre:

| Pendiente de Construction II | Casos |
|---|---|
| Resultado visual real de `dibujarRadar` (rótulos sin pisarse, aspecto de puntos y anillos, el radar con muy pocos puntos) | CP-97 a CP-106 |
| Aspecto del formulario con sus 56 campos | CP-78 a CP-82, CP-85, CP-86 |
| Usabilidad de escribir 56 calificaciones: orden de tabulación y si conviene otro control | CP-115, CP-116, CP-117 |
| Longitud de la página y mensaje de error y gráfico lejos de los campos | CP-107 a CP-112, CP-114 |
| Apertura de `Mtx-RADAR.xlsx` en un programa de hojas de cálculo, con componentes incompletos e inválidos | CP-122 a CP-127 |
| Navegadores distintos de Chrome | CP-128, CP-129 |
| Archivo descargado de internet | CP-50 a CP-52 del Módulo 1 (a nivel de archivo; no se repite para RADAR) |
| Disposición de dos columnas en el navegador real del equipo (Chrome, Edge y Firefox), incluido el selector `:has()` | CP-107 a CP-113 (Chrome u otro navegador ya probado), CP-128, CP-129 |

Pendientes que este plan no puede cubrir, declarados aparte: un criterio de tiempo aceptable para escribir las 56 calificaciones (nunca se acordó con el profesor, por eso CP-116 y CP-117 solo anotan la experiencia); el uso desde un celular (no hay requisito, ver las observaciones); la comparación del archivo exportado con uno equivalente del profesor (la estructura de las hojas la fijó el juez y no existe un archivo de referencia del profesor para este modelo); y la comparación del texto fijo con el original más allá de las siete afirmaciones y los 14 títulos de CP-80 y CP-85.

Las pruebas RE.1 a RE.13 de Elaboration I y de la ronda correctiva, que `tests/elaboration1.test.js` ya verifica sin pantalla, se cubren aquí en lo que tienen de visible: RE.1 y RE.2 en CP-87 a CP-89 y CP-94, RE.3 en CP-83 y CP-96, RE.4 y RE.5 en CP-97 a CP-106 y CP-130, RE.13 en CP-106 y CP-130, RE.6 en CP-79 y CP-104, RE.7 en CP-80 a CP-82, RE.9 en CP-95, RE.10 en CP-90 a CP-93 y CP-114, RE.11 en CP-118 a CP-121 y RE.12 en CP-107 a CP-112. RE.8 (los dos componentes de Alineamiento con el mismo título) se cubre en CP-86, en lo que tiene de visible. No se escribió ningún caso para funcionalidad que no existe: no hay botones para agregar ni quitar componentes, ni otro control de calificación, ni versión para celular.

### 5. Correcciones de código hechas durante esta fase

Esta es la excepción a que la Construction III no escribe código. El juez verificó este plan contra la aplicación real, confirmó dos de los hallazgos que el plan había registrado y encontró un tercer problema que el plan no veía, y decidió corregir los tres en una ronda corta de código dentro de esta fase, para que el plan describa el comportamiento final. Cada corrección es un commit aparte, con sus pruebas, y el plan se actualizó en un cuarto commit.

**A. Exportación de un valor no numérico.** En la rama RADAR de `hojasExportacion`, una calificación como abc salía como una celda de error (tipo `e`, valor 15), porque `aNumero` devuelve NaN y SheetJS escribe un NaN como error. Ahora una calificación que no es un número se exporta como el texto escrito; los números y los vacíos quedan como antes. Se amplió la prueba X.7.

**B. Números de puntaje que se pisaban.** Cada número se dibujaba siempre arriba a la derecha de su punto, y con puntajes bajos se amontonaban. Ahora `dibujarRadar` coloca cada número en la primera de ocho posiciones alrededor de su punto que no pise ningún otro texto ni ningún punto; si ninguna queda libre, el número no se dibuja (el punto sigue con su texto emergente y el número sigue en la tabla de resultados). Se reparte primero a los puntos más alejados del centro. Los anchos de texto que usa la colocación son estimados (7 unidades por carácter), así que el resultado no depende de la tipografía. Se agregó la prueba RE.13. Además de la referencia del juez se descartan las posiciones que quedarían fuera del recuadro del SVG.

**C. El panel de la derecha no siempre cabía.** El panel tenía un alto máximo y un desplazamiento propio: con errores y una ventana baja, el gráfico quedaba cortado. Ahora el panel no lleva alto máximo: el recuadro rojo tiene como máximo el 30 % del alto de la ventana y se desplaza por dentro, y el gráfico se escala hacia abajo, conservando su proporción, hasta el alto de la ventana menos un margen (el 70 % menos el margen cuando hay errores). Es solo CSS, dentro de la consulta de medios de 75rem; por debajo de ese ancho nada cambia. No se prueba con el arnés: la prueba RE.12 sigue valiendo porque el marcado no cambió.

Mediciones (navegador integrado del escritorio de Claude, un Chromium, sobre una copia servida desde un servidor temporal fuera del repositorio, ya detenido y borrado).

| Criterio | Medición |
|---|---|
| K1. Exportación (con el libro que genera SheetJS, leído antes de la descarga) | Con D-RA3: E35 es de tipo texto con el valor "abc" (antes, tipo error con valor 15); E16 es numérica con 6 y E19 numérica con 2.5; hay 0 celdas de tipo error en las dos hojas (antes, 1). Con D-RA2 las 360 celdas del archivo son idénticas a las de antes de la corrección. |
| K2. Números de puntaje (1366 por 768; cuadro real de cada texto y de cada punto) | D-RA1: 4 números de 4 puntos, 0 pares texto-texto, 0 pares número-punto. D-RA2: 14 de 14, 0 y 0 (más 1 par dígito de la escala-punto, inherente). D-RA4: 4 de 14, 0 y 0 (más 4 pares dígito de la escala-punto), con los 14 textos emergentes y la tabla con los 14 puntajes. En los 36 conjuntos al azar (3 familias por 12 semillas): 0 pares texto-texto, 0 pares número-punto y 0 elementos fuera del recuadro en todos; se dibujan 4 a 6 de 14 números en "sano", 11 a 14 en "medio" y 14 de 14 en "mixto". Antes de la corrección: D-RA2 tenía 2 pares texto-texto y 2 número-punto; "sano" tenía de 42 a 79 y de 30 a 48, "medio" de 0 a 6 y de 2 a 7, y "mixto" de 0 a 2 y de 0 a 2 (el juez midió, con otra tipografía, 3 pares en D-RA2 y de 22 a 75 y de 28 a 47 en "sano"). |
| K3. Panel completo dentro de la ventana (página desplazada 2000 px) | El borde superior y el inferior del panel quedaron dentro de la ventana, con D-RA2, D-RA3 y D-RA3 con 12 errores más, a 1366 por 650, 1280 por 720, 1366 por 768 y 1920 por 1080 (y también a 1366 por 560). Con 15 errores el recuadro rojo se desplaza por dentro (altura del contenido 378 px, contra la del recuadro: 195, 216, 230 y 324 px en esas cuatro ventanas) y el gráfico entero sigue dentro de la ventana. Con D-RA2 (sin errores) el gráfico mide 576 por 585 px en las cuatro ventanas. Con errores se escala: con D-RA3, 409 por 415 px a 1366 por 650, 457 por 464 a 1280 por 720, 490 por 498 a 1366 por 768 y 576 por 585 a 1920 por 1080. A 1190 por 800, con D-RA3 y 12 errores más, la sección sigue en una columna: el panel no es fijo, el gráfico mide 576 por 585 px, el recuadro rojo no tiene alto máximo ni desplazamiento propio, el orden es formulario, errores, gráfico y resultados, y no hay desborde horizontal. |

Mutaciones, cada una en una copia fuera del repositorio: volver a exportar con `n(valor)` a secas hace fallar X.7 (el texto abc sale como celda de tipo error, 44 de 45 correctas); volver a colocar siempre el número en la primera posición hace fallar RE.13 (D-RA2: 2 números que pisan un punto; D-RA4: 14 números dibujados y 70 textos y 38 puntos pisados; 45 de 46 correctas); y quitar la comprobación de superposición de la colocación da el mismo fallo de RE.13.

Commits de la ronda: A, `16bcee4` ("Construction III (Modulo 3): exportar a Excel como texto una calificacion no numerica de RADAR"); B, `de7db41` ("Construction III (Modulo 3): dibujarRadar coloca los numeros de puntaje sin que se pisen entre si ni con los puntos"); C, `45d963a` ("Construction III (Modulo 3): el panel de RADAR cabe siempre en la ventana, con los errores desplazables y el grafico escalado"); y el de esta actualización del plan. Con la ronda las pruebas son 46 de 46 (eran 45): la prueba nueva es RE.13 y X.7 ganó aserciones.

### Observaciones y puntos nuevos marcados [VERIFICAR] en esta fase

Observaciones al preparar el plan (no se cambió código ni pruebas):

1. **Longitud de la página.** Con las 56 afirmaciones a la vista, la página de RADAR es muy larga: unos 5300 a 6300 px en las ventanas de escritorio medidas (5322 a 1920 por 1080, 5610 a 1366 por 768, 5940 a 1280 por 720 y 6320 a 1200 por 800, cada una con sus datos de prueba; la de 1200 por 800 se midió en la ronda de las dos columnas), y unos 10 000 px a 390 px de ancho. La revisión del juez la midió entre 5600 y 6900 px según la ventana: la diferencia entre las dos mediciones viene de la tipografía de cada equipo, que cambia cuántas líneas ocupa cada afirmación. La ronda de las dos columnas no acortó la página; lo que mejoró es la visibilidad: con una ventana de 75rem o más, el recuadro de errores y el gráfico acompañan a la persona mientras recorre el formulario (CP-108, CP-109, CP-111 y CP-114). Por debajo de ese ancho, el gráfico y los errores siguen al final de la página.
2. **Celulares.** A 390 px de ancho la sección va en una columna, el gráfico mide unos 324 px y sus rótulos quedan en unos 6 px de alto. No hay ningún caso para el celular porque no hay requisito de usarlo; queda como observación.
3. **Gráfico cortado.** Con errores y una ventana baja, el panel de la derecha no cabía entero: el recuadro rojo y el gráfico quedaban cortados por el borde inferior o con una barra de desplazamiento propia. Se corrigió en esta fase (punto 5): el recuadro rojo se desplaza por dentro si son muchos errores y el gráfico se escala hacia abajo solo cuando no cabe. Queda una salvedad, descrita en CP-109: arriba de todo de la página el panel todavía no está fijo y, con errores y una ventana baja, su parte de abajo puede quedar cortada hasta que se empieza a bajar.
4. **Puntos cercanos al centro.** Con D-RA2 había rótulos que se tocaban entre sí y con los puntos. Se corrigió en esta fase (punto 5): cada número de puntaje se coloca donde no pisa otro texto ni un punto y, si no hay lugar (puntos muy juntos, como en D-RA4), no se dibuja: el punto conserva su texto emergente y el número sigue en la tabla de resultados (CP-106 y CP-130). Quedan dos observaciones, ambas inherentes al eje fijo de 0 a 5: con puntajes bajos los propios puntos se tocan entre sí, y algún dígito de la escala (el "0" con D-RA2) queda parcialmente bajo un punto.
5. **Valor no numérico en el Excel.** Una calificación como abc se exportaba como una celda de error de hoja de cálculo (en Excel, #¡VALOR!). Se corrigió en esta fase (punto 5), por decisión del juez: se exporta el texto tal cual ("abc"); 6 y 2.5 siguen exportándose como números (CP-126).
6. **Una calificación borrada queda guardada como texto vacío**, no como `null` (el formulario guarda lo que escribe la persona). Para la aplicación es lo mismo que "sin calificar" y no se ve en pantalla. Ya está documentado en los criterios de Construction II.
7. **Tecla Tab y escritura.** En la simulación se usaron teclas reales para escribir los cuatro valores de Movilización 1, los valores no válidos, el borrado con Retroceso, F5 y la tecla Tab para recorrer los 56 campos (10 pulsaciones llevan al campo 11, y 45 más al 56). La escritura completa de las 56 calificaciones de D-RA2 se reprodujo escribiendo cada valor con un evento de entrada en su campo, en orden, y no con teclas reales; el efecto sobre el estado es el mismo, pero la sensación de escribir 56 valores con Tab no se pudo simular.
8. **Botón y título de la sección.** Igual que en AE, el botón de la navegación y el título de la sección dicen "RADAR" y el nombre "Radar Estratégico" solo aparece como texto emergente del botón.
9. Un texto que contiene una tabulación dentro de un solo campo se trata como un valor no numérico y da el mismo error que CP-90 (se observó al intentar escribir varios valores de una vez con un solo envío de texto).

Puntos [VERIFICAR] nuevos:

1. **Criterios de usabilidad no acordados.** Ningún criterio de usabilidad (cuánto tarda escribir, cuánta página es demasiada, qué tamaño de pantalla debe atenderse) fue acordado con el profesor. CP-116 y CP-117 solo anotan la experiencia.
2. **Control para las calificaciones.** Si el profesor prefiere otro control para calificar de 0 a 5 (un selector con los seis niveles con su texto, o seis botones) en lugar de un campo de texto. Es una pregunta para el profesor; CP-117 recoge la opinión de quien ejecuta el plan.
3. **Exportación de un valor no numérico.** Resuelto en esta fase con el criterio de que el archivo exporta el texto que escribió la persona, de acuerdo con lo ya documentado en Construction II ("el valor tal cual si está fuera de rango"). Está en CP-126; lo decidió el juez y no hace falta confirmarlo con el profesor.
4. **Selector `:has()` y `position: sticky` en Edge y Firefox.** Ya registrado en el punto 7 de Construction II de este módulo. Lo prueban CP-128 y CP-129.
5. **Quién ejecuta este plan, con qué frecuencia y cómo se registran los fallos:** igual que el punto 5 de la Construction III del Módulo 1.
6. **Uso en celular.** Si el curso necesita usar el módulo en un celular. Mientras no haya requisito no hay casos.
7. Siguen abiertos los puntos de las fases anteriores de este módulo: el título repetido de los dos componentes de Alineamiento, la afirmación repetida de Motivación 3 y Gestión 1 y las erratas del texto del profesor (ver "Puntos [VERIFICAR] para el profesor" en Construction II), y las calificaciones con decimales (el 2.5 se rechaza, como en AE, EFI, EFE y MPC).
