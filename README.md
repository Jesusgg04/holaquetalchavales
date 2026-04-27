# Automatizacion de CSV con GAS(Google As Scripsts)
# Utilizar las funciones de leerCelda, leerCeldas, setCelda
# Funcion 1 
/**
 *Inteligencia de Negocios
 * Jesus Luna
 */

// Funciones principales(aqui se definen las funciones)


function getCeldas(rango) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Si no hay rango, toma toda la hoja con datos automáticamente
  if (!rango) { return hoja.getDataRange().getValues(); }
  return hoja.getRange(rango).getValues();
}


function setCelda(zelda, balor) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (zelda) { hoja.getRange(zelda).setValue(balor); }
}


// Primera funcion (metodos de contactto)


function ls(arg) {
  var tabla = getCeldas("A:AV"); 
  var encabezados = tabla[0];
  var resultados = [encabezados];

// Aqui se "cuenta" en que columna está el sitio Web, Correo y Telefono(Numero real - 1, ya que la columna A siempre es 0)

  var colTel = 27;    // Columna AB
  var colCorreo = 28; // Columna AC
  var colWeb = 29;    // Columna AD

  for (var i = 1; i < tabla.length; i++) {
    var fila = tabla[i];
    if (fila[0] == "" || fila[0] == null) { continue; }

    var tieneTel = fila[colTel] && fila[colTel].toString().trim() != "";
    var tieneCorr = fila[colCorreo] && fila[colCorreo].toString().trim() != "";
    var tieneWeb = fila[colWeb] && fila[colWeb].toString().trim() != "";

    var cumple = false;
    if (arg == 't' && tieneTel) cumple = true;
    if (arg == 'c' && tieneCorr) cumple = true;
    if (arg == 'w' && tieneWeb) cumple = true;
    if (arg == 'a' && tieneTel && tieneCorr && tieneWeb) cumple = true;

    if (cumple) { resultados.push(fila); }
  }
  
  enviarAHoja(resultados, "Reporte_Contacto_" + arg);
}


// Funcion 2 Viabilidad


function lsV(tipoVialidad, nombreVialidad) {
  var tabla = getCeldas("A:AV");
  var resultados = [tabla[0]];


  // Columna de viabilidad y nombre de ella(6 y 7)
  var colTipo = 6;
  var colNombre = 7;

  for (var i = 1; i < tabla.length; i++) {
    var fila = tabla[i];
    if (fila[0] == "" || fila[0] == null) { continue; }

    var t = fila[colTipo].toString().toLowerCase().trim();
    var n = fila[colNombre].toString().toLowerCase().trim();


  // Filtra por tipo exacto y que el nombre contenga el texto buscado
    if (t == tipoVialidad.toLowerCase().trim() && n.indexOf(nombreVialidad.toLowerCase().trim()) > -1) {
      resultados.push(fila);
    }
    
  }
  enviarAHoja(resultados, "Reporte_Vialidad");
}


// Funcion para que los resultados se escriban

function enviarAHoja(matriz, nombreHoja) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
  
  hoja.clear();
  if (matriz.length > 1) {
    hoja.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
    setCelda("Z1", "Último reporte generado: " + nombreHoja);
  } else {
    SpreadsheetApp.getUi().alert("No se encontraron datos para " + nombreHoja);
  }
}

// Prueba

function ejecutarTodo() {
  // Aquí pruebas la que quieras
  ls('c'); // Busca correos
  // lsV("CALLE", "DELANTE"); // Ejemplo vialidad
}

### ¿Cómo funciona el código? (Flujo de trabajo)

# **Paso 1 - Recolección:** El programa toma todos los datos del INEGI que están en la hoja principal.
# **Paso 2 - Inspección:** Revisa negocio por negocio. Si el negocio tiene la información que buscamos (teléfono, correo o calle), lo guarda. Si no, lo ignora.
# **Paso 3 - Entrega:** Crea una hoja nueva y pega ahí solamente los negocios que encontramos, avisando al usuario cuando el proceso termina.
