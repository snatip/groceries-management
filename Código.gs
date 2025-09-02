// Nombre de la hoja de resumen y datos
const RESUMEN_SHEET = "Resumen";
const DATOS_SHEET = "Datos";
const GRUPOS_SHEET = "Grupos";
const HORARIOS_SHEET = "Horarios";

// Cache para evitar múltiples lecturas de la hoja Datos
let personasCache = null;
let cacheTimestamp = null;
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutos

/**
 * Carga la configuración de personas desde la hoja "Datos"
 */
function cargarConfiguracionPersonas() {
  const now = new Date().getTime();
  
  // Usar cache si es válido
  if (personasCache && cacheTimestamp && (now - cacheTimestamp < CACHE_DURATION)) {
    return personasCache;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const datosSheet = ss.getSheetByName(DATOS_SHEET);
    
    if (!datosSheet) {
      throw new Error(`La hoja "${DATOS_SHEET}" no existe. Por favor, créala con las columnas: Nombre, Iniciales, Foto`);
    }
    
    const data = datosSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error(`La hoja "${DATOS_SHEET}" debe tener al menos una fila de datos además del encabezado`);
    }
    
    const header = data[0];
    const nombreCol = header.findIndex(col => col.toString().toLowerCase().includes('nombre'));
    const inicialesCol = header.findIndex(col => col.toString().toLowerCase().includes('iniciales'));
    const fotoCol = header.findIndex(col => col.toString().toLowerCase().includes('foto'));
    
    if (nombreCol === -1 || inicialesCol === -1) {
      throw new Error(`La hoja "${DATOS_SHEET}" debe tener columnas 'Nombre' e 'Iniciales'`);
    }
    
    const personas = {};
    const personasArray = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nombre = row[nombreCol]?.toString().trim();
      const iniciales = row[inicialesCol]?.toString().trim();
      const foto = fotoCol !== -1 ? row[fotoCol]?.toString().trim() : '';
      
      if (nombre && iniciales) {
        personas[nombre] = iniciales;
        personasArray.push({
          nombre: nombre,
          iniciales: iniciales,
          foto: foto || ''
        });
      }
    }
    
    personasCache = {
      mapeoIniciales: personas,
      listaPersonas: personasArray
    };
    cacheTimestamp = now;
    
    return personasCache;
    
  } catch (error) {
    console.error('Error cargando configuración de personas:', error.message);
    // Fallback al mapeo original si hay error
    const fallbackPersonas = {
      "Santi": "S",
      "Lucía": "L", 
      "Polito": "MP",
      "Carlos": "C",
      "Erik": "E",
      "Julia": "J",
      "Marina D": "MD",
      "Pablo": "P",
      "Sebas": "S",
      "Ura": "U",
      "Amecor": "A",
      "JP": "JP"
    };
    
    const fallbackArray = Object.keys(fallbackPersonas).map(nombre => ({
      nombre: nombre,
      iniciales: fallbackPersonas[nombre],
      foto: ''
    }));
    
    return {
      mapeoIniciales: fallbackPersonas,
      listaPersonas: fallbackArray
    };
  }
}

/**
 * Crea un menú personalizado en la UI de Google Sheets al abrir el archivo.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestión de Compras')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Configuración Inicial')
        .addItem('Configurar Todo (Libro Nuevo)', 'configurarTodoInicial')
        .addItem('Solo Crear Hoja de Datos', 'crearHojaDatos')
        .addItem('Solo Crear Compra de Ejemplo', 'crearCompraEjemplo')
    )
    .addSeparator()
    .addItem('Nueva Compra', 'mostrarDialogoNuevaCompra')
    .addSeparator()
    .addItem('Crear/Actualizar Hoja de Resumen', 'crearHojaResumenConFormulas')
    .addItem('Limpiar Cache de Personas', 'limpiarCachePersonas')
    .addToUi();
}

/**
 * ¡FUNCIÓN NUEVA!
 * Configura todo desde cero en un libro nuevo: hoja de datos, compra de ejemplo y resumen
 */
function configurarTodoInicial() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Configuración Inicial',
    '¿Estás seguro de que quieres configurar todo desde cero?\n\nEsto creará:\n- Hoja "Datos" con personas de ejemplo\n- Compra de ejemplo "Supermercado_2024"\n- Hoja de Resumen',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // 1. Crear hoja de datos
    crearHojaDatos();

    // 1.5. Crear hoja de grupos
    crearHojaGrupos();
    
    // 2. Crear compra de ejemplo
    crearCompraEjemplo();
    
    // 3. Crear hoja de resumen
    crearHojaResumenConFormulas();
    
    ui.alert('¡Configuración completada!', 'Se han creado todas las hojas necesarias:\n- Datos (con 4 personas de ejemplo)\n- Supermercado_2024 (compra de ejemplo)\n- Resumen\n\nYa puedes usar la aplicación web.', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Hubo un problema durante la configuración: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * ¡FUNCIÓN NUEVA!
 * Crea la hoja "Datos" con personas de ejemplo
 */
function crearHojaDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Eliminar hoja "Datos" si existe
  const existingDatos = ss.getSheetByName(DATOS_SHEET);
  if (existingDatos) {
    ss.deleteSheet(existingDatos);
  }
  
  // Crear nueva hoja "Datos"
  const datosSheet = ss.insertSheet(DATOS_SHEET);
  
  // Configurar encabezados
  datosSheet.getRange("A1:C1").setValues([["Nombre", "Iniciales", "Foto"]]);
  datosSheet.getRange("A1:C1").setFontWeight("bold");
  datosSheet.getRange("A1:C1").setBackground("#4285f4");
  datosSheet.getRange("A1:C1").setFontColor("white");
  
  // Datos de ejemplo
  const datosEjemplo = [
    ["Santi", "S", ""],
    ["Lucía", "L", ""],
    ["Polito", "MP", ""],
    ["Carlos", "C", ""]
  ];
  
  datosSheet.getRange(2, 1, datosEjemplo.length, 3).setValues(datosEjemplo);
  
  // Formatear tabla
  datosSheet.autoResizeColumns(1, 3);
  const lastRow = datosEjemplo.length + 1;
  datosSheet.getRange(1, 1, lastRow, 3).setBorder(true, true, true, true, true, true);
  datosSheet.getRange(2, 1, datosEjemplo.length, 3).setBorder(true, true, true, true, false, true);
  
  /*
  // Añadir nota explicativa
  datosSheet.getRange("A7").setValue("INSTRUCCIONES:");
  datosSheet.getRange("A7").setFontWeight("bold");
  datosSheet.getRange("A8:A11").setValues([
    ["- Añade aquí las personas que participan en las compras"],
    ["- La columna 'Foto' es opcional (URL de imagen)"],
    ["- Después de modificar esta hoja, usa 'Limpiar Cache' del menú"],
    ["- Máximo recomendado: 10-12 personas"]
  ]);
  datosSheet.getRange("A8:A11").setFontStyle("italic");
  datosSheet.getRange("A8:A11").setFontColor("#666666");
  */

  // Limpiar cache para que se recarguen los datos
  limpiarCachePersonas();
}

/**
 * ¡FUNCIÓN NUEVA!
 * Crea la hoja "Grupos" para gestionar grupos de personas
 */
function crearHojaGrupos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Eliminar hoja "Grupos" si existe
  const existingGrupos = ss.getSheetByName(GRUPOS_SHEET);
  if (existingGrupos) {
    ss.deleteSheet(existingGrupos);
  }
  
  // Crear nueva hoja "Grupos"
  const gruposSheet = ss.insertSheet(GRUPOS_SHEET);
  
  // Configurar encabezados
  gruposSheet.getRange("A1:C1").setValues([["ID_Grupo", "Nombre_Grupo", "Miembros"]]);
  gruposSheet.getRange("A1:C1").setFontWeight("bold");
  gruposSheet.getRange("A1:C1").setBackground("#4285f4");
  gruposSheet.getRange("A1:C1").setFontColor("white");
  
  // Formatear tabla
  gruposSheet.autoResizeColumns(1, 3);
  gruposSheet.getRange(1, 1, 1, 3).setBorder(true, true, true, true, true, true);
  
  /*
  // Añadir nota explicativa
  gruposSheet.getRange("A4").setValue("INSTRUCCIONES:");
  gruposSheet.getRange("A4").setFontWeight("bold");
  gruposSheet.getRange("A5:A8").setValues([
    ["- Esta hoja gestiona los grupos de comidas automáticamente"],
    ["- No modificar manualmente"],
    ["- Cada persona puede estar en máximo un grupo"],
    ["- Los grupos se gestionan desde la aplicación web"]
  ]);
  gruposSheet.getRange("A5:A8").setFontStyle("italic");
  gruposSheet.getRange("A5:A8").setFontColor("#666666");*/
}

/**
 * ¡FUNCIÓN NUEVA!
 * Crea una compra de ejemplo con datos realistas
 */
function crearCompraEjemplo() {
  const nombreCompra = "Supermercado_2024";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Eliminar hoja si existe
  const existingSheet = ss.getSheetByName(nombreCompra);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  // Obtener personas de la hoja Datos
  const config = cargarConfiguracionPersonas();
  const personas = config.listaPersonas.map(p => p.nombre);
  
  if (personas.length === 0) {
    throw new Error("No hay personas configuradas en la hoja Datos");
  }
  
  // Crear nueva hoja
  const sheet = ss.insertSheet(nombreCompra);
  
  // Productos de ejemplo
  const productosEjemplo = [
    ["Leche", 1.20],
    ["Pan", 0.85],
    ["Huevos (docena)", 2.45],
    ["Pollo (1kg)", 4.50],
    ["Arroz (1kg)", 1.15],
    ["Pasta", 0.95],
    ["Tomates (1kg)", 2.10],
    ["Aceite Oliva", 3.80],
    ["Yogures", 2.25],
    ["Detergente", 3.50]
  ];
  
  crearEstructuraCompra(sheet, personas, productosEjemplo);
  
  // Añadir algunos tics de ejemplo (productos compartidos típicos)
  // Leche - todos la comparten
  sheet.getRange(2, 4, 1, personas.length).setValue(true);
  // Pan - todos menos el último
  if (personas.length > 1) {
    sheet.getRange(3, 4, 1, personas.length - 1).setValue(true);
  }
  // Huevos - primeros 2
  if (personas.length >= 2) {
    sheet.getRange(4, 4, 1, 2).setValue(true);
  }
  // Pollo - solo el primero
  sheet.getRange(5, 4).setValue(true);
}

/**
 * ¡FUNCIÓN NUEVA!
 * Muestra un diálogo para crear una nueva compra
 */
function mostrarDialogoNuevaCompra() {
  const ui = SpreadsheetApp.getUi();
  
  // Verificar que existe la hoja Datos
  const config = cargarConfiguracionPersonas();
  if (config.listaPersonas.length === 0) {
    ui.alert('Error', 'Primero debes configurar la hoja "Datos" con las personas.\n\nUsa: Gestión de Compras > Configuración Inicial > Solo Crear Hoja de Datos', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt(
    'Nueva Compra',
    'Introduce el nombre de la nueva compra:\n(Ej: "Mercadona_Enero", "Carrefour_2024", etc.)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const nombreCompra = response.getResponseText().trim();
    
    if (!nombreCompra) {
      ui.alert('Error', 'El nombre de la compra no puede estar vacío.', ui.ButtonSet.OK);
      return;
    }
    
    // Verificar que el nombre no esté en uso
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSheetByName(nombreCompra)) {
      ui.alert('Error', `Ya existe una hoja con el nombre "${nombreCompra}".`, ui.ButtonSet.OK);
      return;
    }
    
    try {
      crearNuevaCompra(nombreCompra);
      ui.alert('¡Éxito!', `La compra "${nombreCompra}" se ha creado correctamente.\n\nAhora puedes:\n1. Añadir productos y precios manualmente\n2. Marcar quién comparte cada producto\n3. Ver los resultados en la aplicación web`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', `No se pudo crear la compra: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * ¡FUNCIÓN NUEVA!
 * Crea una nueva hoja de compra con la estructura correcta
 */
function crearNuevaCompra(nombreCompra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener personas configuradas
  const config = cargarConfiguracionPersonas();
  const personas = config.listaPersonas.map(p => p.nombre);
  
  if (personas.length === 0) {
    throw new Error("No hay personas configuradas en la hoja Datos");
  }
  
  // Crear nueva hoja
  const sheet = ss.insertSheet(nombreCompra);
  
  // Crear estructura básica sin productos
  crearEstructuraCompra(sheet, personas, []);
  
  // Añadir filas vacías para que el usuario pueda empezar a introducir datos
  const filasVacias = [
    ["", ""],
    ["", ""],
    ["", ""],
    ["", ""],
    ["", ""]
  ];
  
  sheet.getRange(2, 1, filasVacias.length, 2).setValues(filasVacias);
  
  // Seleccionar la primera celda para que el usuario pueda empezar a escribir
  sheet.getRange("A2").activate();
}

/**
 * ¡FUNCIÓN AUXILIAR NUEVA!
 * Crea la estructura común de una hoja de compra
 */
function crearEstructuraCompra(sheet, personas, productos) {
  // Configurar encabezados
  const headers = ["Producto", "Subtotal", "Total pp"].concat(personas);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatear encabezados
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.getRange(1, 1, 1, headers.length).setBackground("#4285f4");
  sheet.getRange(1, 1, 1, headers.length).setFontColor("white");
  
  // Si hay productos, añadirlos
  if (productos && productos.length > 0) {
    // Añadir productos con subtotales
    const productosData = productos.map(p => [p[0], p[1]]);
    sheet.getRange(2, 1, productos.length, 2).setValues(productosData);
    
    // Crear fórmulas para "Total pp" (columna C)
    for (let i = 0; i < productos.length; i++) {
      const row = i + 2;
      const personasRange = `${columnToLetter(4)}${row}:${columnToLetter(3 + personas.length)}${row}`;
      const formula = `=IF(COUNTIF(${personasRange}, TRUE)>0, B${row}/COUNTIF(${personasRange}, TRUE), 0)`;
      sheet.getRange(row, 3).setFormula(formula);
    }
  }
  
  // Formatear columnas
  sheet.getRange("B:B").setNumberFormat("0.00\" €\""); // Subtotal
  sheet.getRange("C:C").setNumberFormat("0.00\" €\""); // Total pp
  
  // Configurar columnas de personas como checkboxes
  const personasStartCol = 4;
  const lastRow = Math.max(productos.length + 1, 10);

  for (let i = 0; i < personas.length; i++) {
    const col = personasStartCol + i;
    sheet.getRange(2, col, Math.max(productos.length, lastRow-1), 1).insertCheckboxes();
  }
  
  // Auto-redimensionar columnas
  sheet.autoResizeColumns(1, headers.length);
  
  // Configurar bordes
  
  
  /*
  // Añadir nota explicativa al final
  const notaRow = lastRow + 2;
  sheet.getRange(notaRow, 1).setValue("INSTRUCCIONES:");
  sheet.getRange(notaRow, 1).setFontWeight("bold");
  sheet.getRange(notaRow + 1, 1, 4, 1).setValues([
    ["1. Escribe el nombre del producto en la columna A"],
    ["2. Introduce el precio en la columna B (Subtotal)"],
    ["3. Marca con ✓ las personas que comparten cada producto"],
    ["4. La columna 'Total pp' se calcula automáticamente"]
  ]);
  sheet.getRange(notaRow + 1, 1, 4, 1).setFontStyle("italic");
  sheet.getRange(notaRow + 1, 1, 4, 1).setFontColor("#666666");*/
}

/**
 * Limpia el cache de personas para forzar recarga desde la hoja Datos
 */
function limpiarCachePersonas() {
  personasCache = null;
  cacheTimestamp = null;
  SpreadsheetApp.getUi().alert('Cache de personas limpiado correctamente');
}

/**
 * Devuelve el HTML de la aplicación web.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Devuelve los nombres de todas las hojas de compra.
 */
function getSheetNames() {
  const excludeSheets = [RESUMEN_SHEET, DATOS_SHEET, GRUPOS_SHEET, HORARIOS_SHEET];
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    .filter(n => !excludeSheets.includes(n));
}

/**
 * Devuelve la lista de personas configuradas
 */
function getPersonasConfiguracion() {
  const config = cargarConfiguracionPersonas();
  return config.listaPersonas;
}

/**
 * Devuelve el mapeo de iniciales para uso en el frontend
 */
function getInicialesPersonas() {
  const config = cargarConfiguracionPersonas();
  return config.mapeoIniciales;
}

/**
 * Obtiene todos los grupos existentes
 */
function getGrupos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gruposSheet = ss.getSheetByName(GRUPOS_SHEET);
    
    if (!gruposSheet) {
      return [];
    }
    
    const data = gruposSheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      nombre: row[1],
      miembros: row[2] ? row[2].split(',').map(m => m.trim()).filter(m => m) : []
    })).filter(grupo => grupo.nombre); // Solo grupos con nombre
  } catch (error) {
    console.error('Error obteniendo grupos:', error);
    return [];
  }
}

/**
 * Obtiene el grupo de una persona específica
 */
function getGrupoPersona(nombrePersona) {
  try {
    const grupos = getGrupos();
    return grupos.find(grupo => grupo.miembros.includes(nombrePersona)) || null;
  } catch (error) {
    console.error('Error obteniendo grupo de persona:', error);
    return null;
  }
}

/**
 * Crea un nuevo grupo
 */
function crearGrupo(nombreGrupo, miembrosSeleccionados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let gruposSheet = ss.getSheetByName(GRUPOS_SHEET);
    
    if (!gruposSheet) {
      crearHojaGrupos();
      gruposSheet = ss.getSheetByName(GRUPOS_SHEET);
    }
    
    // Verificar que el nombre no esté en uso
    const grupos = getGrupos();
    if (grupos.some(g => g.nombre.toLowerCase() === nombreGrupo.toLowerCase())) {
      throw new Error('Ya existe un grupo con ese nombre');
    }
    
    // Remover personas de otros grupos primero
    if (miembrosSeleccionados && miembrosSeleccionados.length > 0) {
      miembrosSeleccionados.forEach(miembro => {
        salirDeGrupo(miembro);
      });
    }
    
    // Generar ID único
    const id = 'grupo_' + Date.now();
    const miembrosStr = miembrosSeleccionados ? miembrosSeleccionados.join(', ') : '';
    
    // Añadir nuevo grupo
    gruposSheet.appendRow([id, nombreGrupo, miembrosStr]);
    
    return { success: true, message: 'Grupo creado exitosamente' };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Une a una persona a un grupo existente
 */
function unirseAGrupo(nombrePersona, idGrupo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gruposSheet = ss.getSheetByName(GRUPOS_SHEET);
    
    if (!gruposSheet) {
      throw new Error('No existe la hoja de grupos');
    }
    
    // Primero salir del grupo actual
    salirDeGrupo(nombrePersona);
    
    const data = gruposSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idGrupo) {
        const miembrosActuales = data[i][2] ? data[i][2].split(',').map(m => m.trim()).filter(m => m) : [];
        
        if (!miembrosActuales.includes(nombrePersona)) {
          miembrosActuales.push(nombrePersona);
          gruposSheet.getRange(i + 1, 3).setValue(miembrosActuales.join(', '));
        }
        
        return { success: true, message: 'Te has unido al grupo exitosamente' };
      }
    }
    
    throw new Error('Grupo no encontrado');
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Saca a una persona de su grupo actual
 */
function salirDeGrupo(nombrePersona) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gruposSheet = ss.getSheetByName(GRUPOS_SHEET);
    
    if (!gruposSheet) {
      return { success: true, message: 'No hay grupos' };
    }
    
    const data = gruposSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const miembrosActuales = data[i][2] ? data[i][2].split(',').map(m => m.trim()).filter(m => m) : [];
      
      if (miembrosActuales.includes(nombrePersona)) {
        const nuevosMiembros = miembrosActuales.filter(m => m !== nombrePersona);
        gruposSheet.getRange(i + 1, 3).setValue(nuevosMiembros.join(', '));
        break;
      }
    }
    
    return { success: true, message: 'Has salido del grupo exitosamente' };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Calcula los gastos totales de una persona específica
 */
function calcularGastosTotalesPersona(nombrePersona) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resumenSheet = ss.getSheetByName(RESUMEN_SHEET);
    
    if (!resumenSheet) {
      return 0; // Si no hay hoja de resumen, gastos = 0
    }
    
    const data = resumenSheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
    let totalGastos = 0;
    
    // Sumar todos los gastos de la persona (excluyendo filas de "__TOTAL__")
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const persona = row[1]; // Columna B (Persona)
      const importe = parseFloat(row[2]) || 0; // Columna C (Total a Pagar)
      
      if (persona === nombrePersona && persona !== "__TOTAL__") {
        totalGastos += importe;
      }
    }
    
    return totalGastos;
    
  } catch (error) {
    console.error(`Error calculando gastos para ${nombrePersona}:`, error.message);
    return 0;
  }
}

/**
 * FUNCIÓN MODIFICADA: Ahora incluye información sobre quién más tiene cada producto
 * Devuelve la lista de productos para una persona en una compra concreta,
 * incluyendo tanto las iniciales como los nombres completos de otras personas que también tienen el producto.
 */
function getData(nombre, sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!ss) throw new Error(`La hoja "${sheetName}" no existe.`);
    const data = ss.getDataRange().getValues();
    const header = data[0];
    const colIndex = header.indexOf(nombre);
    if (colIndex === -1) return []; // La persona no está en esta compra

    const config = cargarConfiguracionPersonas();

    // Obtener índices de todas las personas (columnas a partir de la 3)
    const personasInfo = header.slice(2).map((persona, index) => {
      const nombrePersona = persona.toString().trim();
      const personaConfig = config.listaPersonas.find(p => p.nombre === nombrePersona);
      
      return {
        nombre: nombrePersona,
        colIndex: index + 2,
        iniciales: personaConfig ? personaConfig.iniciales : nombrePersona.substring(0, 2).toUpperCase(),
        foto: personaConfig ? personaConfig.foto : ''
      };
    }).filter(p => p.nombre && p.nombre.toLowerCase() !== "total pp");

    return data.slice(1).map(row => {
      if (!row[0]) return null; // Ignora filas vacías de producto
      
      // Encontrar quién más tiene este producto (excluyendo al usuario actual)
      const otrasPersonasConProducto = personasInfo
        .filter(persona => persona.nombre !== nombre && !!row[persona.colIndex])
        .map(persona => ({
          iniciales: persona.iniciales,
          nombreCompleto: persona.nombre,
          foto: persona.foto
        }));

      return {
        producto: row[0],
        subtotal: row[1],
        checked: !!row[colIndex],
        otrasPersonas: otrasPersonasConProducto
      };
    }).filter(item => item !== null); // Ignora filas vacías
  } catch (e) {
    console.error(`Error en getData: ${e.message}`);
    return [];
  }
}

/**
 * Actualiza una casilla (tick) en la hoja de una compra.
 * Se ha eliminado la llamada a actualizarResumen() ya que las fórmulas lo hacen automáticamente.
 */
function updateData(nombre, producto, valor, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = ss.getDataRange().getValues();
  const header = data[0];
  const colIndex = header.indexOf(nombre);
  if (colIndex === -1) return;

  // Encontrar la fila del producto
  const rowIndex = data.slice(1).findIndex(row => row[0] === producto);
  if (rowIndex !== -1) {
    // +2 porque getRange es 1-based y saltamos la cabecera
    ss.getRange(rowIndex + 2, colIndex + 1).setValue(valor);
  }
}

/**
 * Obtiene los datos de la hoja de Resumen para mostrarlos en la app.
 * Ahora es mucho más simple: solo lee los valores que ya están calculados por las fórmulas.
 */
function getResumenData() {
  const resumenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESUMEN_SHEET);
  if (!resumenSheet) {
    // Si la hoja no existe, la crea vacía y devuelve un array vacío.
    // El usuario debe generarla desde el menú de la hoja de cálculo.
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(RESUMEN_SHEET).appendRow(["Compra", "Persona", "Total a Pagar"]);
    return [];
  }
  
  const data = resumenSheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Solo cabecera o vacía

  return data.slice(1).map(row => ({
    compra: row[0],
    persona: row[1],
    total: parseFloat(row[2]) || 0
  }));
}

/**
 * ¡FUNCIÓN ACTUALIZADA!
 * Crea o reemplaza la hoja de resumen con fórmulas robustas.
 * Ahora los totales van como filas separadas identificables para la web app.
 */
function crearHojaResumenConFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName() !== RESUMEN_SHEET && s.getName() !== DATOS_SHEET && s.getName() !== GRUPOS_SHEET && s.getName() !== HORARIOS_SHEET);
  
  const oldResumenSheet = ss.getSheetByName(RESUMEN_SHEET);
  if (oldResumenSheet) {
    ss.deleteSheet(oldResumenSheet);
  }

  const resumenSheet = ss.insertSheet(RESUMEN_SHEET, 0);
  resumenSheet.appendRow(["Compra", "Persona", "Total a Pagar"]);
  resumenSheet.getRange("A1:C1").setFontWeight("bold");

  const config = cargarConfiguracionPersonas();
  const resumenData = [];
  let currentRow = 2; // Empezamos en la fila 2 (después del header)

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    const header = data[0];
    // Los nombres de las personas comienzan desde la 3ª columna (índice 2)
    const personas = header.slice(2);
    const subtotalCol = "C:C"; // Columna de subtotales
    
    const personasFiltradasData = [];

    personas.forEach((persona, index) => {
      // Ignorar la columna "Total pp" y columnas sin nombre.
      const nombreLimpio = persona.toString().trim();
      if (!nombreLimpio || nombreLimpio.toLowerCase() === "total pp") {
        return; // Saltar esta iteración y no añadir la fila al resumen
      }
      
      const personaColIndex = index + 3; 
      const personaCol = columnToLetter(personaColIndex) + ":" + columnToLetter(personaColIndex);

      const formula = `=SUMIF('${sheetName}'!${personaCol},TRUE,'${sheetName}'!${subtotalCol})`;
      
      personasFiltradasData.push([sheetName, nombreLimpio, formula]);
    });

    // Agregar las filas de personas para esta compra
    personasFiltradasData.forEach(row => {
      resumenData.push(row);
      resumenSheet.getRange(currentRow, 1).setValue(row[0]);   // Sheet name
      resumenSheet.getRange(currentRow, 2).setValue(row[1]);   // Person
      resumenSheet.getRange(currentRow, 3).setFormula(row[2]); // Formula
      currentRow++;
    });

    // Agregar fila de total para esta compra si tiene personas
    // Usamos un prefijo especial "__TOTAL__" para identificar estas filas
    if (personasFiltradasData.length > 0) {
      const startRow = currentRow - personasFiltradasData.length;
      const endRow = currentRow - 1;
      const totalFormula = `=SUM(C${startRow}:C${endRow})`;
      
      resumenSheet.getRange(currentRow, 1).setValue(sheetName);
      resumenSheet.getRange(currentRow, 2).setValue("__TOTAL__");
      resumenSheet.getRange(currentRow, 3).setFormula(totalFormula);
      
      // Formatear la fila de total
      resumenSheet.getRange(currentRow, 1, 1, 3).setFontWeight("bold");
      resumenSheet.getRange(currentRow, 1, 1, 3).setBackground("#e8f4fd");
      
      currentRow++;
      
      // Agregar una fila vacía como separador entre compras
      currentRow++;
    }
  });

  if (resumenData.length > 0) {
    resumenSheet.getRange("C:C").setNumberFormat("0.00\" €\"");
    resumenSheet.autoResizeColumns(1, 3);
    
    // Aplicar bordes para mejor visualización
    const lastRow = currentRow - 1;
    resumenSheet.getRange(1, 1, lastRow, 3).setBorder(true, true, true, true, true, true);
  }

  //SpreadsheetApp.getUi().alert('¡Hoja de Resumen creada/actualizada con éxito!');
}

/**
 * Helper para convertir un índice de columna numérico a su letra (ej: 1 -> A, 3 -> C)
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Actualiza la hoja de resumen desde la aplicación web
 */
function actualizarResumenDesdeApp() {
  try {
    crearHojaResumenConFormulas();
    return { success: true, message: 'Resumen actualizado correctamente' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Crea una nueva compra desde la aplicación web
 */
/**
 * Crea una nueva compra desde la aplicación web
 */
function crearCompraDesdeApp(nombreCompra, productos, asignaciones) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Verificar que el nombre no esté en uso
    if (ss.getSheetByName(nombreCompra)) {
      throw new Error(`Ya existe una compra con el nombre "${nombreCompra}"`);
    }
    
    // Obtener personas configuradas
    const config = cargarConfiguracionPersonas();
    const personas = config.listaPersonas.map(p => p.nombre);
    
    if (personas.length === 0) {
      throw new Error("No hay personas configuradas en la hoja Datos");
    }
    
    // Crear nueva hoja
    const sheet = ss.insertSheet(nombreCompra);
    
    // Configurar encabezados
    const headers = ["Producto", "Subtotal", "Total pp"].concat(personas);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Formatear encabezados
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#4285f4");
    sheet.getRange(1, 1, 1, headers.length).setFontColor("white");
    
    // Añadir productos
    if (productos && productos.length > 0) {
      const productosData = productos.map(p => [p.nombre, p.precio]);
      sheet.getRange(2, 1, productos.length, 2).setValues(productosData);
      
      // Crear fórmulas para "Total pp" (columna C)
      for (let i = 0; i < productos.length; i++) {
        const row = i + 2;
        const personasRange = `${columnToLetter(4)}${row}:${columnToLetter(3 + personas.length)}${row}`;
        const formula = `=IF(COUNTIF(${personasRange}, TRUE)>0, B${row}/COUNTIF(${personasRange}, TRUE), 0)`;
        sheet.getRange(row, 3).setFormula(formula);
      }
      
      // Aplicar asignaciones de productos
      if (asignaciones && asignaciones.length > 0) {
        asignaciones.forEach((asignacion, productoIndex) => {
          const row = productoIndex + 2;
          asignacion.forEach(persona => {
            const personaColIndex = personas.indexOf(persona) + 4;
            if (personaColIndex >= 4) {
              sheet.getRange(row, personaColIndex).setValue(true);
            }
          });
        });
      }
    }
    
    // Formatear columnas
    sheet.getRange("B:B").setNumberFormat("0.00\" €\"");
    sheet.getRange("C:C").setNumberFormat("0.00\" €\"");
    
    // Configurar columnas de personas como checkboxes
    const personasStartCol = 4;
    const lastRow = Math.max(productos.length + 1, 10);
    
    for (let i = 0; i < personas.length; i++) {
      const col = personasStartCol + i;
      sheet.getRange(2, col, Math.max(productos.length, lastRow-1), 1).insertCheckboxes();
    }
    
    // Auto-redimensionar columnas
    sheet.autoResizeColumns(1, headers.length);
    
    return { success: true, message: 'Compra creada exitosamente' };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// === FUNCIONES PARA GESTIÓN DE HORARIOS Y ASISTENCIA ===

/**
 * Obtiene o crea la hoja de horarios
 */
function getHojaHorarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let horariosSheet = ss.getSheetByName(HORARIOS_SHEET);
  
  if (!horariosSheet) {
    // Crear la hoja si no existe
    horariosSheet = ss.insertSheet(HORARIOS_SHEET);
    
    // Configurar encabezados
    const headers = [
      "Nombre", 
      "Asistencia", 
      "Trayecto1_Salida", 
      "Trayecto1_Llegada", 
      "Trayecto1_Transporte",
      "Trayecto2_Salida", 
      "Trayecto2_Llegada", 
      "Trayecto2_Transporte",
      "Trayecto3_Salida", 
      "Trayecto3_Llegada", 
      "Trayecto3_Transporte",
      "Trayecto4_Salida", 
      "Trayecto4_Llegada", 
      "Trayecto4_Transporte",
      "Notas"
    ];
    
    horariosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    horariosSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    horariosSheet.getRange(1, 1, 1, headers.length).setBackground("#4285f4");
    horariosSheet.getRange(1, 1, 1, headers.length).setFontColor("white");
    horariosSheet.autoResizeColumns(1, headers.length);
  }
  
  return horariosSheet;
}

/**
 * Guarda la confirmación de asistencia de una persona
 */
function saveAsistencia(nombre, asistio) {
  try {
    const horariosSheet = getHojaHorarios();
    const data = horariosSheet.getDataRange().getValues();
    
    // Buscar si la persona ya existe en la hoja
    let personaRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nombre) {
        personaRow = i;
        break;
      }
    }
    
    if (personaRow === -1) {
      // Añadir nueva fila para la persona
      const config = cargarConfiguracionPersonas();
      const personaConfig = config.listaPersonas.find(p => p.nombre === nombre);
      
      if (personaConfig) {
        const newRow = [nombre, asistio];
        // Añadir celdas vacías para los trayectos y notas
        for (let i = 0; i < 14; i++) {
          newRow.push("");
        }
        horariosSheet.appendRow(newRow);
      } else {
        return { success: false, message: "Persona no encontrada en la configuración" };
      }
    } else {
      // Actualizar asistencia existente
      horariosSheet.getRange(personaRow + 1, 2).setValue(asistio);
    }
    
    return { success: true, message: "Asistencia guardada correctamente" };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene la confirmación de asistencia de una persona
 */
function getAsistencia(nombre) {
  try {
    const horariosSheet = getHojaHorarios();
    const data = horariosSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nombre) {
        return data[i][1] === true;
      }
    }
    
    return false; // Por defecto, no ha confirmado asistencia
    
  } catch (error) {
    return false;
  }
}

/**
 * Obtiene el contador de personas que han confirmado asistencia
 */
function getContadorAsistencia() {
  try {
    const horariosSheet = getHojaHorarios();
    const data = horariosSheet.getDataRange().getValues();
    
    let contador = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === true) {
        contador++;
      }
    }
    
    return contador;
    
  } catch (error) {
    return 0;
  }
}

/**
 * Guarda los datos de viaje de una persona
 */
function saveDatosViaje(nombre, datosViaje) {
  try {
    const horariosSheet = getHojaHorarios();
    const data = horariosSheet.getDataRange().getValues();
    
    // Buscar si la persona ya existe en la hoja
    let personaRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nombre) {
        personaRow = i;
        break;
      }
    }
    
    if (personaRow === -1) {
      // Añadir nueva fila para la persona
      const config = cargarConfiguracionPersonas();
      const personaConfig = config.listaPersonas.find(p => p.nombre === nombre);
      
      if (personaConfig) {
        const asistencia = getAsistencia(nombre);
        const newRow = [nombre, asistencia];
        
        // Añadir datos de los trayectos
        newRow.push(
          datosViaje.trayecto1.salida || "",
          datosViaje.trayecto1.llegada || "",
          datosViaje.trayecto1.transporte || "",
          datosViaje.trayecto2.salida || "",
          datosViaje.trayecto2.llegada || "",
          datosViaje.trayecto2.transporte || "",
          datosViaje.trayecto3.salida || "",
          datosViaje.trayecto3.llegada || "",
          datosViaje.trayecto3.transporte || "",
          datosViaje.trayecto4.salida || "",
          datosViaje.trayecto4.llegada || "",
          datosViaje.trayecto4.transporte || "",
          datosViaje.notas || ""
        );
        
        horariosSheet.appendRow(newRow);
      } else {
        return { success: false, message: "Persona no encontrada en la configuración" };
      }
    } else {
      // Actualizar datos de viaje existentes
      const rowNum = personaRow + 1;
      
      // Trayecto 1
      horariosSheet.getRange(rowNum, 3).setValue(datosViaje.trayecto1.salida || "");
      horariosSheet.getRange(rowNum, 4).setValue(datosViaje.trayecto1.llegada || "");
      horariosSheet.getRange(rowNum, 5).setValue(datosViaje.trayecto1.transporte || "");
      
      // Trayecto 2
      horariosSheet.getRange(rowNum, 6).setValue(datosViaje.trayecto2.salida || "");
      horariosSheet.getRange(rowNum, 7).setValue(datosViaje.trayecto2.llegada || "");
      horariosSheet.getRange(rowNum, 8).setValue(datosViaje.trayecto2.transporte || "");
      
      // Trayecto 3
      horariosSheet.getRange(rowNum, 9).setValue(datosViaje.trayecto3.salida || "");
      horariosSheet.getRange(rowNum, 10).setValue(datosViaje.trayecto3.llegada || "");
      horariosSheet.getRange(rowNum, 11).setValue(datosViaje.trayecto3.transporte || "");
      
      // Trayecto 4
      horariosSheet.getRange(rowNum, 12).setValue(datosViaje.trayecto4.salida || "");
      horariosSheet.getRange(rowNum, 13).setValue(datosViaje.trayecto4.llegada || "");
      horariosSheet.getRange(rowNum, 14).setValue(datosViaje.trayecto4.transporte || "");
      
      // Notas
      horariosSheet.getRange(rowNum, 15).setValue(datosViaje.notas || "");
    }
    
    return { success: true, message: "Datos de viaje guardados correctamente" };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene los datos de viaje de una persona
 */
function getDatosViaje(nombre) {
  try {
    const horariosSheet = getHojaHorarios();
    const data = horariosSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nombre) {
        const row = data[i];
        return {
          trayecto1: {
            salida: row[2] || "",
            llegada: row[3] || "",
            transporte: row[4] || ""
          },
          trayecto2: {
            salida: row[5] || "",
            llegada: row[6] || "",
            transporte: row[7] || ""
          },
          trayecto3: {
            salida: row[8] || "",
            llegada: row[9] || "",
            transporte: row[10] || ""
          },
          trayecto4: {
            salida: row[11] || "",
            llegada: row[12] || "",
            transporte: row[13] || ""
          },
          notas: row[14] || ""
        };
      }
    }
    
    return {}; // Retornar objeto vacío si no hay datos
    
  } catch (error) {
    return {};
  }
}

/**
 * Función para reparar la estructura de la hoja de horarios
 */
function repararHojaHorarios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let horariosSheet = ss.getSheetByName(HORARIOS_SHEET);
    
    if (!horariosSheet) {
      // Crear la hoja si no existe
      horariosSheet = ss.insertSheet(HORARIOS_SHEET);
      
      // Configurar encabezados
      const headers = [
        "Nombre", 
        "Asistencia", 
        "Trayecto1_Salida", 
        "Trayecto1_Llegada", 
        "Trayecto1_Transporte",
        "Trayecto2_Salida", 
        "Trayecto2_Llegada", 
        "Trayecto2_Transporte",
        "Trayecto3_Salida", 
        "Trayecto3_Llegada", 
        "Trayecto3_Transporte",
        "Trayecto4_Salida", 
        "Trayecto4_Llegada", 
        "Trayecto4_Transporte",
        "Notas"
      ];
      
      horariosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      horariosSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      horariosSheet.getRange(1, 1, 1, headers.length).setBackground("#4285f4");
      horariosSheet.getRange(1, 1, 1, headers.length).setFontColor("white");
      
      return { success: true, message: "Hoja de horarios creada correctamente" };
    }
    
    const data = horariosSheet.getDataRange().getValues();
    const expectedCols = 15;
    let reparaciones = 0;
    
    // Verificar y reparar cada fila
    for (let i = 1; i < data.length; i++) {
      if (!data[i] || data[i].length !== expectedCols) {
        // Crear una nueva fila con la longitud correcta
        const newRow = new Array(expectedCols).fill("");
        
        // Copiar datos existentes
        if (data[i]) {
          for (let j = 0; j < Math.min(data[i].length, expectedCols); j++) {
            newRow[j] = data[i][j];
          }
        }
        
        // Reemplazar la fila
        horariosSheet.getRange(i + 1, 1, 1, expectedCols).setValues([newRow]);
        reparaciones++;
      }
    }
    
    return { 
      success: true, 
      message: `Hoja reparada correctamente. ${reparaciones} filas reparadas.`,
      reparaciones: reparaciones
    };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Función de depuración para verificar la estructura de la hoja de horarios
 */
function debugHojaHorarios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const horariosSheet = ss.getSheetByName(HORARIOS_SHEET);
    
    if (!horariosSheet) {
      return { error: "La hoja de horarios no existe", sheetExists: false };
    }
    
    const data = horariosSheet.getDataRange().getValues();
    const range = horariosSheet.getDataRange();
    
    const debugInfo = {
      sheetName: horariosSheet.getName(),
      sheetExists: true,
      totalRows: range.getNumRows(),
      totalCols: range.getNumColumns(),
      expectedCols: 15, // Nombre + Asistencia + 12 trayectos + Notas
      headers: data[0] || [],
      sampleRows: [],
      allRowsLength: []
    };

        // Verificar cada fila
    for (let i = 0; i < data.length; i++) {
      debugInfo.allRowsLength.push({
        row: i + 1,
        length: data[i] ? data[i].length : 0
      });
      
      // Mostrar primeras 5 filas como muestra
      if (i > 0 && i < 6) {
        debugInfo.sampleRows.push({
          row: i + 1,
          data: data[i] || [],
          length: data[i] ? data[i].length : 0
        });
      }
    }
    
    // Verificar si hay filas con longitud incorrecta
    const rowsWithWrongLength = debugInfo.allRowsLength.filter(row => row.length !== debugInfo.expectedCols);
    if (rowsWithWrongLength.length > 0) {
      debugInfo.rowsWithWrongLength = rowsWithWrongLength;
    }
    
    return debugInfo;
  } catch (error) {
    return { 
      error: error.message, 
      stack: error.stack,
      errorMessage: error.toString()
    };
  }
}

function getAllHorarios() {
  try {
    console.log('Iniciando getAllHorarios...');
    const horariosSheet = getHojaHorarios();
    console.log('Hoja de horarios obtenida:', horariosSheet.getName());
    
    const data = horariosSheet.getDataRange().getValues();
    console.log('Datos de hoja obtenidos, filas:', data.length);
    
    const config = cargarConfiguracionPersonas();
    console.log('Configuración de personas cargada, personas:', config.listaPersonas.length);
    
    const todosHorarios = [];
    
    // Obtener todas las personas de la configuración
    config.listaPersonas.forEach(personaConfig => {
      const nombre = personaConfig.nombre;
      
      // Buscar datos en la hoja de horarios
      let horariosPersona = {
        nombre: nombre,
        foto: personaConfig.foto || "",
        asistencia: false,
        datosViaje: null
      };
      
      for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][0] === nombre) {
          const row = data[i];
          console.log('Encontrada persona:', nombre, 'fila:', i + 1, 'longitud:', row ? row.length : 'null');
          
          horariosPersona.asistencia = row[1] === true;
          horariosPersona.datosViaje = {
            trayecto1: {
              salida: formatearHora(row[2]),      // ← CORREGIDO
              llegada: formatearHora(row[3]),    // ← CORREGIDO
              transporte: row[4] || ""
            },
            trayecto2: {
              salida: formatearHora(row[5]),      // ← CORREGIDO
              llegada: formatearHora(row[6]),    // ← CORREGIDO
              transporte: row[7] || ""
            },
            trayecto3: {
              salida: formatearHora(row[8]),      // ← CORREGIDO
              llegada: formatearHora(row[9]),    // ← CORREGIDO
              transporte: row[10] || ""
            },
            trayecto4: {
              salida: formatearHora(row[11]),     // ← CORREGIDO
              llegada: formatearHora(row[12]),    // ← CORREGIDO
              transporte: row[13] || ""
            },
            notas: row[14] || ""
          };
          break;
        }
      }
      
      todosHorarios.push(horariosPersona);
    });
    
    console.log('Retornando todosHorarios:', todosHorarios.length, 'personas');
    return todosHorarios;
    
  } catch (error) {
    console.error('Error en getAllHorarios:', error.message, error.stack);
    return [];
  }
}

/**
 * Convierte un objeto Date de Google Sheets a string de hora
 */
function formatearHora(dateObj) {
  if (!dateObj) return "";
  
  // Si es un string, devolverlo directamente
  if (typeof dateObj === 'string') return dateObj;
  
  // Si es un objeto Date, formatearlo
  if (typeof dateObj === 'object' && dateObj.getTime) {
    try {
      // Usar Utilities.formatDate para formatear la hora
      return Utilities.formatDate(dateObj, "GMT", "HH:mm");
    } catch (e) {
      // Si falla, intentar formatear manualmente
      const hours = dateObj.getHours().toString().padStart(2, '0');
      const minutes = dateObj.getMinutes().toString().padStart(2, '0');
      return `${hours}:${minutes}`;
    }
  }
  
  // Para cualquier otro caso, devolver string vacío
  return "";
}

/**
 * Convierte un objeto Date de Google Sheets a string de hora
 */
function formatearHora(dateObj) {
  if (!dateObj) return "";
  
  // Si es un string, devolverlo directamente
  if (typeof dateObj === 'string') return dateObj;
  
  // Si es un objeto Date, formatearlo
  if (typeof dateObj === 'object' && dateObj.getTime) {
    try {
      // Usar Utilities.formatDate para formatear la hora
      return Utilities.formatDate(dateObj, "GMT", "HH:mm");
    } catch (e) {
      // Si falla, intentar formatear manualmente
      const hours = dateObj.getHours().toString().padStart(2, '0');
      const minutes = dateObj.getMinutes().toString().padStart(2, '0');
      return `${hours}:${minutes}`;
    }
  }
  
  // Para cualquier otro caso, devolver string vacío
  return "";
}
