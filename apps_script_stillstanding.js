// ═══════════════════════════════════════════════════════════════════
// STILL STANDING — Apps Script
// Versión: Kit vendible (genérico)
// 
// INSTRUCCIONES DE INSTALACIÓN (para quien hace el setup):
//
// 1. Abre el Google Sheets del cliente
// 2. Menú: Extensiones → Apps Script
// 3. Borra todo el código que aparece y pega este archivo completo
// 4. Clic en "Guardar" (ícono de disco)
// 5. Menú: Implementar → Nueva implementación
//    - Tipo: Aplicación web
//    - Ejecutar como: Yo (tu cuenta)
//    - Quién tiene acceso: Cualquier usuario
// 6. Clic en "Implementar" → copia la URL que aparece
// 7. Pega esa URL en la pantalla de configuración de la app Still Standing
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = "Balance";

// ── UTILIDADES ───────────────────────────────────────────────────────
function r(v) { return Math.round(v * 100) / 100; }

function formatFecha(fecha) {
  return fecha.getDate().toString().padStart(2, "0") + "/" +
    (fecha.getMonth() + 1).toString().padStart(2, "0");
}

function getRangoMetodo(metodo, mesEfectivo, anioEfectivo) {
  const mes = mesEfectivo;
  const anio = anioEfectivo;
  let desde, hasta;
  // Métodos con corte mensual calendario
  if (metodo === "Efectivo" || metodo === "Transferencia") {
    desde = new Date(anio, mes, 1);
    hasta = new Date(anio, mes + 1, 0);
  }
  // Métodos con corte 15→14 (ej. tarjetas de crédito)
  else if (metodo === "Tarjeta 1" || metodo === "Tarjeta 2") {
    desde = new Date(anio, mes - 1, 15);
    hasta = new Date(anio, mes, 14);
  }
  // Método con corte 19→18
  else if (metodo === "Tarjeta 3") {
    desde = new Date(anio, mes - 1, 19);
    hasta = new Date(anio, mes, 18);
  } else {
    // Default: mes calendario
    desde = new Date(anio, mes, 1);
    hasta = new Date(anio, mes + 1, 0);
  }
  return { desde, hasta };
}

// ── doGet: BACKEND DE LA APP WEB ─────────────────────────────────────
function doGet(e) {
  const cmd = e.parameter.cmd ? e.parameter.cmd.toLowerCase().trim() : "";
  const mesParam = e.parameter.mes ? parseInt(e.parameter.mes) : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const egresosSheet = ss.getSheetByName("Egresos");
  const ingresosSheet = ss.getSheetByName("Ingresos");

  if (mesParam) {
    const fechaMes = new Date(new Date().getFullYear(), mesParam - 1, 1);
    sheet.getRange("B1").setValue(fechaMes);
    SpreadsheetApp.flush();
  }

  const mesEfectivo = mesParam ? mesParam - 1 : new Date().getMonth();
  const anioEfectivo = new Date().getFullYear();

  // ── HISTORIAL ────────────────────────────────────────────────────
  if (cmd === "historial") {
    const ultimaFila = egresosSheet.getLastRow();
    const desde = Math.max(2, ultimaFila - 49);
    const data = egresosSheet.getRange(desde, 1, ultimaFila - desde + 1, 5).getValues();
    const resultado = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const fila = data[i];
      if (!fila[0]) continue;
      const fecha = fila[0];
      if (fecha.getMonth() !== mesEfectivo || fecha.getFullYear() !== anioEfectivo) continue;
      resultado.push({
        fila: desde + i,
        fecha: formatFecha(fecha),
        monto: r(fila[1]),
        categoria: fila[2],
        metodo: fila[3],
        descripcion: fila[4]
      });
    }
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── ELIMINAR ─────────────────────────────────────────────────────
  if (cmd === "eliminar") {
    const fila = parseInt(e.parameter.fila);
    if (!fila || fila < 2) return ContentService.createTextOutput("Error: fila inválida")
      .setMimeType(ContentService.MimeType.TEXT);
    egresosSheet.deleteRow(fila);
    return ContentService.createTextOutput("✅ Registro eliminado")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // ── EDITAR ───────────────────────────────────────────────────────
  if (cmd === "editar") {
    const fila = parseInt(e.parameter.fila);
    const monto = parseFloat(e.parameter.monto);
    const categoria = e.parameter.categoria;
    const metodo = e.parameter.metodo;
    const descripcion = e.parameter.descripcion;
    const fecha = e.parameter.fecha;
    if (!fila || fila < 2) return ContentService.createTextOutput("Error: fila inválida")
      .setMimeType(ContentService.MimeType.TEXT);
    const partes = fecha.split('-');
    egresosSheet.getRange(fila, 1).setValue(new Date(partes[0], partes[1] - 1, partes[2]));
    egresosSheet.getRange(fila, 2).setValue(monto);
    egresosSheet.getRange(fila, 3).setValue(categoria);
    egresosSheet.getRange(fila, 4).setValue(metodo);
    egresosSheet.getRange(fila, 5).setValue(descripcion);
    return ContentService.createTextOutput("✅ Registro actualizado")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // ── REGISTRAR GASTO ──────────────────────────────────────────────
  if (cmd === "registrar") {
    const fecha = e.parameter.fecha;
    const monto = parseFloat(e.parameter.monto);
    const categoria = e.parameter.categoria;
    const metodo = e.parameter.metodo;
    const descripcion = e.parameter.descripcion;
    if (!fecha || !monto || !categoria || !metodo || !descripcion)
      return ContentService.createTextOutput("Error: faltan datos")
        .setMimeType(ContentService.MimeType.TEXT);
    const ultimaFila = egresosSheet.getLastRow() + 1;
    const partes = fecha.split('-');
    egresosSheet.getRange(ultimaFila, 1).setValue(new Date(partes[0], partes[1] - 1, partes[2]));
    egresosSheet.getRange(ultimaFila, 2).setValue(monto);
    egresosSheet.getRange(ultimaFila, 3).setValue(categoria);
    egresosSheet.getRange(ultimaFila, 4).setValue(metodo);
    egresosSheet.getRange(ultimaFila, 5).setValue(descripcion);
    return ContentService.createTextOutput("✅ Gasto registrado")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // ── REGISTRAR INGRESO ─────────────────────────────────────────────
  if (cmd === "registrar_ingreso") {
    const fecha = e.parameter.fecha;
    const monto = parseFloat(e.parameter.monto);
    const fuente = e.parameter.fuente;
    const descripcion = e.parameter.descripcion;
    if (!fecha || !monto || !fuente || !descripcion)
      return ContentService.createTextOutput("Error: faltan datos")
        .setMimeType(ContentService.MimeType.TEXT);
    const ultimaFila = ingresosSheet.getLastRow() + 1;
    const partes = fecha.split('-');
    ingresosSheet.getRange(ultimaFila, 1).setValue(new Date(partes[0], partes[1] - 1, partes[2]));
    ingresosSheet.getRange(ultimaFila, 2).setValue(descripcion);
    ingresosSheet.getRange(ultimaFila, 3).setValue(fuente);
    ingresosSheet.getRange(ultimaFila, 4).setValue(monto);
    return ContentService.createTextOutput("✅ Ingreso registrado")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // ── CHART ─────────────────────────────────────────────────────────
  if (cmd === "chart") {
    const egresosData = egresosSheet.getRange("A2:C1000").getValues();
    const categorias = {};
    egresosData.forEach(row => {
      const fecha = row[0];
      if (!(fecha instanceof Date)) return;
      if (fecha.getMonth() !== mesEfectivo || fecha.getFullYear() !== anioEfectivo) return;
      const cat = row[2];
      if (!cat) return;
      categorias[cat] = (categorias[cat] || 0) + row[1];  // row[1] = Monto (col B)
    });
    // También leer monto desde col B de Egresos
    const gastosArr = Object.entries(categorias).map(([nombre, valor]) => ({
      nombre, valor: r(valor)
    }));
    const sueldo = r(sheet.getRange("B2").getValue());
    return ContentService.createTextOutput(JSON.stringify({ sueldo, gastos: gastosArr }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── SUSCRIPCIONES ─────────────────────────────────────────────────
  if (cmd === "suscripciones") {
    const subSheet = ss.getSheetByName("Config_Suscripciones");
    if (!subSheet) return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
    const data = subSheet.getRange("A2:F20").getValues();
    const diaHoy = new Date().getDate();
    const resultado = [];
    data.forEach(row => {
      if (!row[0]) return;
      const diasFaltantes = row[2] - diaHoy;
      const alerta = (diasFaltantes >= 0 && diasFaltantes <= 7) ? " ⚠️" : "";
      resultado.push(row[0] + alerta + "  $" + row[1] + " — día " + row[2] + " — " + row[4]);
    });
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── PRESUPUESTOS ──────────────────────────────────────────────────
  if (cmd === "presupuestos") {
    const presSheet = ss.getSheetByName("Config_Presupuestos");
    if (!presSheet) return ContentService.createTextOutput(JSON.stringify({ error: "Hoja Config_Presupuestos no encontrada" }))
      .setMimeType(ContentService.MimeType.JSON);
    const presData = presSheet.getRange("A2:C20").getValues();
    const egresosData = egresosSheet.getRange("A2:E1000").getValues();
    const gastosReales = {};
    egresosData.forEach(row => {
      const fecha = row[0];
      if (!(fecha instanceof Date)) return;
      if (fecha.getMonth() !== mesEfectivo || fecha.getFullYear() !== anioEfectivo) return;
      const cat = row[2];
      if (!cat) return;
      gastosReales[cat] = (gastosReales[cat] || 0) + row[1];
    });
    const resultado = [];
    presData.forEach(row => {
      if (!row[0] || !row[1]) return;
      const categoria = row[0];
      const limite = row[1];
      const icono = row[2] || "📌";
      const gastado = r(gastosReales[categoria] || 0);
      const pct = limite > 0 ? Math.round((gastado / limite) * 100) : 0;
      resultado.push({ categoria, icono, limite: r(limite), gastado, pct });
    });
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── HISTÓRICO ─────────────────────────────────────────────────────
  if (cmd === "historico") {
    const mesesAtras = e.parameter.n ? parseInt(e.parameter.n) : 6;
    const egresosData = egresosSheet.getRange("A2:E1000").getValues();
    const ingresosData = ingresosSheet ? ingresosSheet.getRange("A2:D1000").getValues() : [];
    const hoy = new Date();
    const resultado = [];
    for (let i = mesesAtras - 1; i >= 0; i--) {
      const fecha = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const m = fecha.getMonth();
      const a = fecha.getFullYear();
      const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
      let egresos = 0;
      egresosData.forEach(row => {
        const f = row[0];
        if (!(f instanceof Date)) return;
        if (f.getMonth() === m && f.getFullYear() === a) egresos += row[1];
      });
      let ingresos = 0;
      ingresosData.forEach(row => {
        const f = row[0];
        if (!(f instanceof Date)) return;
        if (f.getMonth() === m && f.getFullYear() === a) ingresos += row[3];
      });
      resultado.push({
        label: meses[m] + " " + a,
        mes: m + 1, anio: a,
        ingresos: r(ingresos),
        egresos: r(egresos),
        balance: r(ingresos - egresos)
      });
    }
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── INGRESOS DETALLE ──────────────────────────────────────────────
  if (cmd === "ingresos_detalle") {
    if (!ingresosSheet) return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
    const data = ingresosSheet.getRange("A2:D1000").getValues();
    const resultado = [];
    data.forEach(row => {
      const fecha = row[0];
      if (!(fecha instanceof Date)) return;
      if (fecha.getMonth() !== mesEfectivo || fecha.getFullYear() !== anioEfectivo) return;
      resultado.push({
        fecha: formatFecha(fecha),
        descripcion: row[1],
        fuente: row[2],
        monto: r(row[3])
      });
    });
    resultado.sort((a, b) => {
      const [da, ma] = a.fecha.split('/').map(Number);
      const [db, mb] = b.fecha.split('/').map(Number);
      return (ma * 100 + da) - (mb * 100 + db);
    });
    const total = r(resultado.reduce((s, x) => s + x.monto, 0));
    return ContentService.createTextOutput(JSON.stringify({ items: resultado, total }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── CONFIG: categorías y métodos desde Sheets ────────────────────
  // La app lee esto al arrancar para construir los dropdowns dinámicamente
  if (cmd === "config") {
    const presSheet = ss.getSheetByName("Config_Presupuestos");
    const metodosSheet = ss.getSheetByName("Config_Metodos");

    const categorias = [];
    if (presSheet) {
      presSheet.getRange("A2:C20").getValues().forEach(row => {
        if (row[0]) categorias.push({ nombre: row[0], icono: row[2] || "📌" });
      });
    }

    const metodos = [];
    if (metodosSheet) {
      metodosSheet.getRange("A2:C20").getValues().forEach(row => {
        if (row[0]) metodos.push({ nombre: row[0], tipo: row[1] || "calendario", dia: row[2] || 1 });
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ categorias, metodos }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── INFO ──────────────────────────────────────────────────────────
  if (cmd === "info") {
    const url = ss.getUrl();
    return ContentService.createTextOutput(JSON.stringify({ sheets_url: url }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── BALANCE (texto) ───────────────────────────────────────────────
  if (cmd === "balance") {
    const respuesta = "💰 Balance Neto: $" + r(sheet.getRange("B4").getValue());
    return ContentService.createTextOutput(respuesta).setMimeType(ContentService.MimeType.TEXT);
  }

  return ContentService.createTextOutput("Comando no reconocido.")
    .setMimeType(ContentService.MimeType.TEXT);
}
