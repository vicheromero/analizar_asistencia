/**
 * Se ejecuta cuando la hoja de cálculo se abre.
 * Crea un menú personalizado en la UI de Google Sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Reportes de Asistencia')
    .addItem('Generar Reporte', 'reorganizarMarcaciones')
    .addToUi();
}

/**
 * --- NUEVA FUNCIÓN ---
 * Reemplaza texto en una hoja específica.
 * Útil para corregir errores de codificación de caracteres (ej: Ã‘ por Ñ).
 */
function reemplazarTextoEnHoja(hoja, buscar, reemplazar) {
  if (!hoja) return; // Si la hoja no existe, no hacer nada
  try {
    const textFinder = hoja.createTextFinder(buscar);
    textFinder.replaceAllWith(reemplazar);
  } catch (err) {
    console.warn(`No se pudo reemplazar texto en la hoja "${hoja.getName()}". Error: ${err.message}`);
  }
}

/**
 * Función principal para reorganizar las marcaciones de asistencia.
 * Lee desde la hoja "Marcaciones", procesa los datos y genera dos hojas nuevas:
 * "Marcaciones Reorganizadas" y "Resumen de Asistencia".
 */
function reorganizarMarcaciones() {
  // --- MEJORA: Añadir feedback al usuario y manejo de errores ---
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Iniciando la generación del reporte. Esto puede tardar unos segundos...');
    const startTime = new Date(); // Medir tiempo

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- MEJORA (Solicitada por el usuario) ---
    // Corregir error de 'Ñ' en las hojas de origen *antes* de leer los datos.
    const textoBuscar = "PEÃ‘AHERRERA SALCEDO FABRICIO MARCELO";
    const textoReemplazar = "PEÑAHERRERA SALCEDO FABRICIO MARCELO";
    
    reemplazarTextoEnHoja(ss.getSheetByName("Marcaciones"), textoBuscar, textoReemplazar);
    reemplazarTextoEnHoja(ss.getSheetByName("Turnos"), textoBuscar, textoReemplazar);
    reemplazarTextoEnHoja(ss.getSheetByName("Ausentismo"), textoBuscar, textoReemplazar);
    // --- Fin de la mejora ---

    const hojaOriginal = ss.getSheetByName("Marcaciones");
    if (!hojaOriginal) {
      ui.alert("No se encontró la hoja 'Marcaciones'. Asegúrate de que la hoja exista.");
      return;
    }

    // Eliminar hojas previas
    let hojaResultados = ss.getSheetByName("Marcaciones Reorganizadas");
    if (hojaResultados) ss.deleteSheet(hojaResultados);
    let hojaResumen = ss.getSheetByName("Resumen de Asistencia");
    if (hojaResumen) ss.deleteSheet(hojaResumen);

    hojaResultados = ss.insertSheet("Marcaciones Reorganizadas");
    hojaResumen = ss.insertSheet("Resumen de Asistencia");

    const datos = hojaOriginal.getDataRange().getValues();
    const encabezados = datos[0];

    const nuevosEncabezados = [
      "Nombre y Apellido", "Fecha", "marcacion1", "Atraso",
      "marcacion2", "marcacion3", "T. Almuerzo", "marcacion4", "Día laborado"
    ];

    hojaResultados.getRange(1, 1, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]).setFontWeight("bold");

    // --- Obtener datos de configuración ---
    const turnosEmpleados = obtenerTurnosEmpleados(ss);
    const ausentismosEmpleados = obtenerAusentismosEmpleados(ss);
    const feriados = obtenerFeriados(ss);

    // --- Paso 1: Detectar rango de fechas del mes ---
    const indiceFecha = encabezados.indexOf("Fecha");
    if (indiceFecha === -1) {
       ui.alert("No se encontró la columna 'Fecha' en la hoja 'Marcaciones'.");
       return;
    }
    const indiceNombre = encabezados.indexOf("Nombre y Apellido");
    if (indiceNombre === -1) {
       ui.alert("No se encontró la columna 'Nombre y Apellido' en la hoja 'Marcaciones'.");
       return;
    }
    const indiceTipo = encabezados.indexOf("Tipo de Registro");
    if (indiceTipo === -1) {
       ui.alert("No se encontró la columna 'Tipo de Registro' en la hoja 'Marcaciones'.");
       return;
    }
    const indiceHora = encabezados.indexOf("Hora");
    if (indiceHora === -1) {
       ui.alert("No se encontró la columna 'Hora' en la hoja 'Marcaciones'.");
       return;
    }


    const fechas = datos.slice(1).map(f => {
      let d = new Date(f[indiceFecha]);
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    });
    
    if (fechas.length === 0) {
      ui.alert("No se encontraron datos en la hoja 'Marcaciones'.");
      return;
    }

    const ref = fechas[0];
    const primerDiaMes = new Date(ref.getFullYear(), ref.getMonth(), 1);
    const ultimoDiaMes = new Date(ref.getFullYear(), ref.getMonth() + 1, 0);

    // --- Paso 2: Guardar marcaciones en un objeto para acceso rápido ---
    const marcacionesPorPersonaFecha = {};
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const nombre = fila[indiceNombre];
      const fecha = new Date(fila[indiceFecha]);
      const tipo = fila[indiceTipo];
      const hora = fila[indiceHora];

      const clave = `${nombre}_${fecha.toDateString()}`;
      if (!marcacionesPorPersonaFecha[clave]) {
        marcacionesPorPersonaFecha[clave] = { ingreso: "", salida: "", inicio: "", fin: "" };
      }
      if (tipo === "Ingreso") marcacionesPorPersonaFecha[clave].ingreso = hora;
      if (tipo === "Salida") marcacionesPorPersonaFecha[clave].salida = hora;
      if (tipo === "Inicio descanso") marcacionesPorPersonaFecha[clave].inicio = hora;
      if (tipo === "Fin descanso") marcacionesPorPersonaFecha[clave].fin = hora;
    }

    // --- Paso 3: Generar filas para todos los días del mes por cada técnico ---
    const nombres = [...new Set(datos.slice(1).map(f => f[indiceNombre]))];
    const datosReorganizados = [];

    nombres.forEach(nombre => {
      for (let d = new Date(primerDiaMes); d <= ultimoDiaMes; d.setDate(d.getDate() + 1)) {
        const clave = `${nombre}_${d.toDateString()}`;
        const registro = marcacionesPorPersonaFecha[clave] || { ingreso: "", salida: "", inicio: "", fin: "" };

        const esLaborable = esDiaLaborable(new Date(d), nombre, turnosEmpleados);
        const tieneMarcaciones = registro.ingreso || registro.salida || registro.inicio || registro.fin;
        const estaEnAusentismo = verificarAusentismo(nombre, new Date(d), ausentismosEmpleados);
        const esFeriado = verificarFeriado(new Date(d), feriados);

        let ingreso, salida, inicioDescanso, finDescanso;

        if (esFeriado && !tieneMarcaciones) {
          ingreso = "Feriado";
          salida = "Feriado";
          inicioDescanso = "Feriado";
          finDescanso = "Feriado";
        } else if (esFeriado && tieneMarcaciones) {
          ingreso = registro.ingreso || "Horas extras";
          salida = registro.salida || "Horas extras";
          inicioDescanso = registro.inicio || "Horas extras";
          finDescanso = registro.fin || "Horas extras";
        } else if (estaEnAusentismo && !tieneMarcaciones) {
          ingreso = "Permiso";
          salida = "Permiso";
          inicioDescanso = "Permiso";
          finDescanso = "Permiso";
        } else if (estaEnAusentismo && tieneMarcaciones) {
          ingreso = registro.ingreso || "Permiso";
          salida = registro.salida || "Permiso";
          inicioDescanso = registro.inicio || "Permiso";
          finDescanso = registro.fin || "Permiso";
        } else if (!esLaborable && !tieneMarcaciones) {
          ingreso = "Día de descanso";
          salida = "Día de descanso";
          inicioDescanso = "Día de descanso";
          finDescanso = "Día de descanso";
        } else if (!esLaborable && tieneMarcaciones) {
          ingreso = registro.ingreso || "Horas extras";
          salida = registro.salida || "Horas extras";
          inicioDescanso = registro.inicio || "Horas extras";
          finDescanso = registro.fin || "Horas extras";
        } else {
          ingreso = registro.ingreso || "Falta marcación";
          salida = registro.salida || "Falta marcación";
          inicioDescanso = registro.inicio || "Falta marcación";
          finDescanso = registro.fin || "Falta marcación";
        }

        datosReorganizados.push([
          nombre,
          new Date(d),
          ingreso,
          "", // Atraso (fórmula)
          salida,
          inicioDescanso,
          "", // T. Almuerzo (fórmula)
          finDescanso,
          ""  // Día laborado (fórmula)
        ]);
      }
    });

    // Ordenar por nombre y fecha
    datosReorganizados.sort((a, b) => {
      if (a[0] < b[0]) return -1;
      if (a[0] > b[0]) return 1;
      return a[1] - b[1];
    });

    // --- Paso 4: Escribir datos en hoja ---
    if (datosReorganizados.length > 0) {
      hojaResultados.getRange(2, 1, datosReorganizados.length, nuevosEncabezados.length).setValues(datosReorganizados);
    }

    // --- Paso 5: Dar formato y aplicar fórmulas ---
    const filas = datosReorganizados.length;
    if (filas > 0) {
      // Formato de Fechas y Horas
      hojaResultados.getRange(2, 2, filas, 1).setNumberFormat("dd/MM/yyyy");
      const columnasHora = [3, 4, 5, 6, 7, 8];
      for (const columna of columnasHora) {
        hojaResultados.getRange(2, columna, filas, 1).setNumberFormat("HH:mm:ss");
      }

      // --- OPTIMIZACIÓN ---
      // Aplicar fórmulas y formatos en lote
      aplicarFormulasConTurnos(hojaResultados, filas, turnosEmpleados, feriados);
      aplicarFormatoCondicional(hojaResultados, filas, turnosEmpleados, feriados);
      // aplicarFormatoHorasExtras se llama *dentro* de aplicarFormatoCondicional
      
      hojaResultados.autoResizeColumns(1, nuevosEncabezados.length);
    }

    // --- Resumen de Asistencia (Optimizado) ---
    crearResumenAsistencia(hojaResumen, nombres); // Usar 'nombres' que ya está calculado

    // --- MEJORA: Feedback de éxito ---
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    ui.alert(`¡Reporte generado con éxito en ${duration.toFixed(1)} segundos!`);

  } catch (err) {
    // --- MEJORA: Feedback de error ---
    console.error(err);
    ui.alert('Se produjo un error: ' + err.message + ' (Línea: ' + err.lineNumber + ')');
  }
}

/**
 * Obtiene los turnos de empleados desde la hoja "Turnos"
 */
function obtenerTurnosEmpleados(ss) {
  const turnos = {};
  try {
    const hojaTurnos = ss.getSheetByName("Turnos");
    if (!hojaTurnos) {
      console.warn("Hoja 'Turnos' no encontrada. Se usará lógica por defecto.");
      return turnos;
    }
    const datosTurnos = hojaTurnos.getDataRange().getValues();
    
    for (let i = 1; i < datosTurnos.length; i++) {
      const fila = datosTurnos[i];
      const nombre = fila[0];
      if (nombre) {
        turnos[nombre] = {
          lunes: fila[1] === 1 ? 1 : 0, // Asegurar que sea 1 o 0
          martes: fila[2] === 1 ? 1 : 0,
          miercoles: fila[3] === 1 ? 1 : 0,
          jueves: fila[4] === 1 ? 1 : 0,
          viernes: fila[5] === 1 ? 1 : 0,
          sabado: fila[6] === 1 ? 1 : 0,
          domingo: fila[7] === 1 ? 1 : 0
        };
      }
    }
  } catch (error) {
    console.warn("Error al leer hoja de turnos:", error.message);
  }
  return turnos;
}

/**
 * Obtiene los feriados desde la hoja "Feriado"
 */
function obtenerFeriados(ss) {
  const feriados = [];
  try {
    const hojaFeriado = ss.getSheetByName("Feriado");
    if (!hojaFeriado) {
      console.warn("Hoja 'Feriado' no encontrada.");
      return feriados;
    }
    // Leer solo la columna A
    const datosFeriado = hojaFeriado.getRange(1, 1, hojaFeriado.getLastRow(), 1).getValues();
    
    for (let i = 0; i < datosFeriado.length; i++) {
      const fecha = datosFeriado[i][0]; // Columna A
      if (fecha && fecha instanceof Date) {
        const fechaNormalizada = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
        feriados.push(fechaNormalizada);
      }
    }
  } catch (error) {
    console.warn("Error al leer hoja de feriados:", error.message);
  }
  return feriados;
}

/**
 * Verifica si una fecha es feriado
 */
function verificarFeriado(fecha, feriados) {
  const fechaComparar = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  const fechaTime = fechaComparar.getTime();
  for (const feriado of feriados) {
    if (fechaTime === feriado.getTime()) {
      return true;
    }
  }
  return false;
}

/**
 * Obtiene los ausentismos de empleados desde la hoja "Ausentismo"
 */
function obtenerAusentismosEmpleados(ss) {
  const ausentismos = [];
  try {
    const hojaAusentismo = ss.getSheetByName("Ausentismo");
    if (!hojaAusentismo) {
      console.warn("Hoja 'Ausentismo' no encontrada.");
      return ausentismos;
    }
    const datosAusentismo = hojaAusentismo.getDataRange().getValues();
    const encabezadosAusentismo = datosAusentismo[0];
    
    const indiceNombre = encabezadosAusentismo.indexOf("Nombre Empleado");
    const indiceInicio = encabezadosAusentismo.indexOf("Inicio de validez");
    const indiceFin = encabezadosAusentismo.indexOf("Fin de validez");
    const indiceDias = encabezadosAusentismo.indexOf("Días de absentismo");
    
    if (indiceNombre === -1 || indiceInicio === -1 || indiceFin === -1 || indiceDias === -1) {
      console.warn("No se encontraron todas las columnas necesarias en la hoja Ausentismo");
      return ausentismos;
    }
    
    for (let i = 1; i < datosAusentismo.length; i++) {
      const fila = datosAusentismo[i];
      const nombre = fila[indiceNombre];
      const fechaInicio = new Date(fila[indiceInicio]);
      const fechaFin = new Date(fila[indiceFin]);
      // Manejar números con comas
      const diasAbsentismo = parseFloat(String(fila[indiceDias]).replace(',', '.')) || 0; 
      
      if (nombre && diasAbsentismo > 0 && fechaInicio instanceof Date && fechaFin instanceof Date) {
        ausentismos.push({
          nombre: nombre,
          fechaInicio: new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), fechaInicio.getDate()),
          fechaFin: new Date(fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate()),
          dias: diasAbsentismo
        });
      }
    }
  } catch (error) {
    console.warn("Error al leer hoja de ausentismos:", error.message);
  }
  return ausentismos;
}

/**
 * Verifica si un empleado está en período de ausentismo en una fecha específica
 */
function verificarAusentismo(nombreEmpleado, fecha, ausentismosEmpleados) {
  for (const ausentismo of ausentismosEmpleados) {
    if (ausentismo.nombre === nombreEmpleado) {
      if (fecha >= ausentismo.fechaInicio && fecha <= ausentismo.fechaFin) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Verifica si es un día laborable para el empleado según sus turnos.
 */
function esDiaLaborable(fecha, nombre, turnosEmpleados) {
  if (!turnosEmpleados[nombre]) {
    const diaSemana = fecha.getDay();
    return diaSemana >= 1 && diaSemana <= 5; // Lunes a Viernes por defecto
  }
  
  const turnos = turnosEmpleados[nombre];
  const diaSemana = fecha.getDay();
  
  switch (diaSemana) {
    case 0: return turnos.domingo === 1;
    case 1: return turnos.lunes === 1;
    case 2: return turnos.martes === 1;
    case 3: return turnos.miercoles === 1;
    case 4: return turnos.jueves === 1;
    case 5: return turnos.viernes === 1;
    case 6: return turnos.sabado === 1;
    default: return false;
  }
}

/**
 * --- OPTIMIZADO ---
 * Aplica las fórmulas de Atraso, T. Almuerzo y Día Laborado en lote.
 */
function aplicarFormulasConTurnos(hojaResultados, filas, turnosEmpleados, feriados) {
  // 1. Leer los datos necesarios (Nombre y Fecha) en una sola llamada
  const data = hojaResultados.getRange(2, 1, filas, 2).getValues(); // Col A y B

  // 2. Preparar arrays para las nuevas fórmulas/valores
  const formulasAtraso = []; // Col D
  const formulasAlmuerzo = []; // Col G
  const formulasDiaLaborado = []; // Col I

  for (let i = 0; i < filas; i++) {
    const filaSheet = i + 2; // Fila real en la hoja (para las fórmulas A1)
    
    // Fórmulas de Atraso (Col D)
    formulasAtraso.push([
      `=IF(OR(C${filaSheet}="",C${filaSheet}="Falta marcación",C${filaSheet}="Día de descanso",C${filaSheet}="Horas extras",C${filaSheet}="Permiso",C${filaSheet}="Feriado"),"",IF(C${filaSheet}>TIME(8,0,59),C${filaSheet}-TIME(8,0,0),""))`
    ]);
    
    // Fórmulas de Tiempo de Almuerzo (Col G)
    formulasAlmuerzo.push([
      `=IF(OR(E${filaSheet}="Falta marcación",E${filaSheet}="Horas extras",F${filaSheet}="Falta marcación",F${filaSheet}="Horas extras"),"Falta marcación",IF(OR(E${filaSheet}="Día de descanso",F${filaSheet}="Día de descanso",E${filaSheet}="Permiso",F${filaSheet}="Permiso",E${filaSheet}="Feriado",F${filaSheet}="Feriado"),"",IF(AND(E${filaSheet}<>"",E${filaSheet}<>"Falta marcación",F${filaSheet}<>"",F${filaSheet}<>"Falta marcación"),F${filaSheet}-E${filaSheet},"Falta marcación")))`
    ]);

    // Lógica de Día Laborado (Col I)
    const nombre = data[i][0]; // Col A
    const fecha = data[i][1]; // Col B
    
    if (nombre && fecha && fecha instanceof Date) {
      const esLaborable = esDiaLaborable(new Date(fecha), nombre, turnosEmpleados);
      const esFeriado = verificarFeriado(new Date(fecha), feriados);
      
      if (esFeriado) {
        formulasDiaLaborado.push([0]);
      } else if (esLaborable) {
        formulasDiaLaborado.push([
          `=IF(AND(C${filaSheet}<>"",C${filaSheet}<>"Falta marcación",C${filaSheet}<>"Día de descanso",C${filaSheet}<>"Horas extras",C${filaSheet}<>"Permiso",C${filaSheet}<>"Feriado",H${filaSheet}<>"",H${filaSheet}<>"Falta marcación",H${filaSheet}<>"Día de descanso",H${filaSheet}<>"Horas extras",H${filaSheet}<>"Permiso",H${filaSheet}<>"Feriado"),1,0)`
        ]);
      } else {
        formulasDiaLaborado.push([0]);
      }
    } else {
      formulasDiaLaborado.push([0]); // Valor por defecto si falta nombre o fecha
    }
  }

  // 3. Escribir todas las fórmulas en la hoja en tres llamadas
  hojaResultados.getRange(2, 4, filas, 1).setValues(formulasAtraso);
  hojaResultados.getRange(2, 7, filas, 1).setValues(formulasAlmuerzo);
  hojaResultados.getRange(2, 9, filas, 1).setValues(formulasDiaLaborado);
}


/**
 * Aplica el formato condicional a la hoja de resultados.
 */
function aplicarFormatoCondicional(hojaResultados, filas, turnosEmpleados, feriados) {
  let reglas = [];

  // Formato para atrasos (Col D)
  const rangoAtrasos = hojaResultados.getRange(2, 4, filas, 1);
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.00694) // > 10 min
      .setBackground("#FF0000")
      .setRanges([rangoAtrasos])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.000001, 0.00694) // 1 seg a 10 min
      .setBackground("#FFFF00")
      .setRanges([rangoAtrasos])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Falta marcación")
      .setBackground("#FF0000")
      .setRanges([rangoAtrasos])
      .build()
  );

  // Formato para tiempo de almuerzo (Col G)
  const rangoAlmuerzo = hojaResultados.getRange(2, 7, filas, 1);
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.041667) // > 1 hora
      .setBackground("#FFFF00")
      .setRanges([rangoAlmuerzo])
      .build()
  );

  // Rangos para marcaciones (C, E, F, H)
  const rangoMarcacion1 = hojaResultados.getRange(2, 3, filas, 1);
  const rangoMarcacion2 = hojaResultados.getRange(2, 5, filas, 1);
  const rangoMarcacion3 = hojaResultados.getRange(2, 6, filas, 1);
  const rangoMarcacion4 = hojaResultados.getRange(2, 8, filas, 1);
  const rangosMarcaciones = [rangoMarcacion1, rangoMarcacion2, rangoMarcacion3, rangoMarcacion4];

  reglas.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Falta marcación").setBackground("#FF0000").setRanges(rangosMarcaciones).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Día de descanso").setBackground("#FFA500").setRanges(rangosMarcaciones).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Horas extras").setBackground("#006400").setFontColor("#FFFFFF").setRanges(rangosMarcaciones).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Permiso").setBackground("#0066CC").setFontColor("#FFFFFF").setRanges(rangosMarcaciones).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Feriado").setBackground("#800080").setFontColor("#FFFFFF").setRanges(rangosMarcaciones).build()
  );

  // Formato para días laborados = 0 (Col I)
  const rangoDiaLaborado = hojaResultados.getRange(2, 9, filas, 1);
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground("#D3D3D3")
      .setRanges([rangoDiaLaborado])
      .build()
  );

  // Aplicar reglas de formato condicional
  hojaResultados.setConditionalFormatRules(reglas);

  // --- OPTIMIZACIÓN ---
  // Aplicar formato de Horas Extras (que no es condicional) en lote
  aplicarFormatoHorasExtras(hojaResultados, filas, turnosEmpleados, feriados);
}

/**
 * --- OPTIMIZADO ---
 * Aplica formato verde a marcaciones reales en días no laborables (Horas Extras).
 * Lee y escribe los colores de fondo y fuente en lote.
 */
function aplicarFormatoHorasExtras(hojaResultados, filas, turnosEmpleados, feriados) {
  
  // 1. Definir el rango de interés (Columnas A hasta H)
  const rangoDatos = hojaResultados.getRange(2, 1, filas, 8); 
  
  // 2. Leer todos los valores, fondos y colores de fuente de una vez
  const values = rangoDatos.getValues();
  const backgrounds = rangoDatos.getBackgrounds();
  const fontColors = rangoDatos.getFontColors();

  // Helper function para verificar si es una marcación real
  const isRealMarking = (valor) => {
    return (valor instanceof Date && valor.getTime() !== 0) || (typeof valor === "number" && valor !== 0);
  }

  // 3. Procesar los arrays en JavaScript (muy rápido)
  for (let i = 0; i < filas; i++) {
    const nombre = values[i][0]; // Col A (índice 0)
    const fecha = values[i][1];  // Col B (índice 1)
    
    if (nombre && fecha && fecha instanceof Date) {
      const esLaborable = esDiaLaborable(new Date(fecha), nombre, turnosEmpleados);
      const esFeriado = verificarFeriado(new Date(fecha), feriados);
      
      if (!esLaborable || esFeriado) {
        // Es día no laborable o feriado. Buscar marcaciones reales.
        const marcacion1 = values[i][2]; // Col C (índice 2)
        const marcacion2 = values[i][4]; // Col E (índice 4)
        const marcacion3 = values[i][5]; // Col F (índice 5)
        const marcacion4 = values[i][7]; // Col H (índice 7)
        
        // Modificar los arrays de colores si se cumple la condición
        if (isRealMarking(marcacion1)) {
          backgrounds[i][2] = "#006400"; // Fondo verde
          fontColors[i][2] = "#FFFFFF";  // Fuente blanca
        }
        if (isRealMarking(marcacion2)) {
          backgrounds[i][4] = "#006400";
          fontColors[i][4] = "#FFFFFF";
        }
        if (isRealMarking(marcacion3)) {
          backgrounds[i][5] = "#006400";
          fontColors[i][5] = "#FFFFFF";
        }
        if (isRealMarking(marcacion4)) {
          backgrounds[i][7] = "#006400";
          fontColors[i][7] = "#FFFFFF";
        }
      }
    }
  }
  
  // 4. Escribir todos los cambios de formato de una vez
  rangoDatos.setBackgrounds(backgrounds);
  rangoDatos.setFontColors(fontColors);
}


/**
 * --- OPTIMIZADO ---
 * Crea el resumen de asistencia en su hoja correspondiente.
 * Escribe todos los datos y fórmulas en un solo `setValues`.
 */
function crearResumenAsistencia(hojaResumen, nombresEmpleados) {
  if (!nombresEmpleados || nombresEmpleados.length === 0) return; // No crear resumen si no hay datos

  const encabezadosResumen = ["Nombre y Apellido", "Días asistidos", "Calculo minutos (días*8*60*0.8)"];
  hojaResumen.getRange(1, 1, 1, 3).setValues([encabezadosResumen]).setFontWeight("bold");

  const datosResumen = [];
  
  // 1. Preparar el array de datos y fórmulas
  for (let i = 0; i < nombresEmpleados.length; i++) {
    const nombreEmpleado = nombresEmpleados[i];
    const filaSheet = i + 2; // Fila real de la hoja de resumen
    
    // Escapar comillas dobles para la fórmula SUMIF
    const nombreEmpleadoEscapado = nombreEmpleado.replace(/"/g, '""'); 
    
    const formulaDias = `=SUMIF('Marcaciones Reorganizadas'!A:A,"${nombreEmpleadoEscapado}",'Marcaciones Reorganizadas'!I:I)`;
    const formulaMinutos = `=B${filaSheet}*8*60*0.8`;
    
    datosResumen.push([nombreEmpleado, formulaDias, formulaMinutos]);
  }
  
  // 2. Escribir todos los datos y fórmulas de una sola vez
  if (datosResumen.length > 0) {
     hojaResumen.getRange(2, 1, datosResumen.length, 3).setValues(datosResumen);
  }
  
  hojaResumen.autoResizeColumns(1, 3);
}
