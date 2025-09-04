/**
 * @OnlyCurrentDoc
 */

/**
 * The sheet names used in the spreadsheet.
 * @enum {string}
 */
const SHEET_NAMES = {
  INVENTORY: 'Inventario', // Will be deprecated
  INVENTARIO_ESTIMADO: 'Inventario Estimado',
  INVENTARIO_REAL: 'Inventario Real',
  ACQUISITIONS: 'Adquisiciones',
  SALES: 'Ventas',
  SKU: 'SKU',
  DISCREPANCIES: 'Discrepancias',
  HISTORICAL_INVENTORY: 'Inventario Histórico',
  TODAY_REPORT: 'REPORTE HOY',
  // Esta es la hoja donde se guardan los reportes de clientes generados.
  REPORTED_CLIENTS: 'ReporteClientes',
  PURCHASE_REQUESTS: 'Solicitudes'
};

/**
 * URLs for the external spreadsheets.
 * @enum {string}
 */
const SOURCE_URLS = {
  OPERACION: 'https://docs.google.com/spreadsheets/d/1hPyDsDHo6Sll6mYY_4YGcPJ4I9FPpG1kQINcidMM-s4/edit',
  ACQUISITIONS_SOURCE: 'https://docs.google.com/spreadsheets/d/1vCZejbBPMh73nbAhdZNYFOlvJvRoMA7PVSCUiLl8MMQ/edit?gid=1415653435#gid=1415653435',
};

/**
 * Creates the initial sheet structure and custom menu.
 * This function runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 DASHBOARD PRINCIPAL')
    .addItem('📊 Abrir Dashboard', 'showDashboard')
    .addToUi();

  ui.createMenu('🔧 Herramientas de Inventario')
    .addItem('🧮 Generar Inventario Estimado', 'generarInventarioEstimado')
    .addSeparator()
    .addItem('📝 Generar Reporte de Cliente', 'showReportGeneratorUI')
    .addItem('🛒 Solicitar producto', 'showPurchaseRequestUI')
    .addSeparator()
    .addItem('⚡ Forzar Actualización de Datos', 'forceRefreshAllImports')
    .addItem('Simular Datos Históricos', 'simulateHistoricalData')
    .addSeparator()
    .addItem('⚠️ Reiniciar Sistema (Puesta en Marcha Blanca)', 'resetSystem')
    .addToUi();

  createSheetsIfNeeded();
  setupImportFormulas();
}

/**
 * Creates the necessary sheets if they don't already exist.
 */
function createSheetsIfNeeded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = Object.values(SHEET_NAMES);

  sheetNames.forEach(name => {
    if (!ss.getSheetByName(name)) {
      const newSheet = ss.insertSheet(name);
      if (name === SHEET_NAMES.ACQUISITIONS) {
        newSheet.appendRow(['Producto Base', 'Formato de Compra', 'Cantidad a Comprar', 'Correccion a Comprar', 'N Formato', 'N Cant', 'N Unidad']);
      } else if (name === SHEET_NAMES.DISCREPANCIES) {
        newSheet.appendRow(['Timestamp', 'Producto Base', 'Stock Esperado', 'Stock Real', 'Discrepancia', 'Unidad de Stock', 'Nota']);
      } else if (name === SHEET_NAMES.HISTORICAL_INVENTORY) {
        newSheet.appendRow(['Timestamp', 'Producto Base', 'Cantidad Stock Real', 'Unidad Venta']);
      } else if (name === SHEET_NAMES.TODAY_REPORT) {
        newSheet.appendRow(['SKU', 'Producto', 'Total Adquirido Hoy', 'Total Vendido Hoy', 'Stock Esperado', 'Último Stock Real', 'Discrepancia', 'Notas']);
      } else if (name === SHEET_NAMES.INVENTARIO_ESTIMADO) {
        newSheet.appendRow(['Producto Base', 'Último Stock (fecha)', 'Stock Esperado', 'Unidad de Inventario']);
      } else if (name === SHEET_NAMES.INVENTARIO_REAL) {
        newSheet.appendRow(['Fecha', 'Producto Base', 'Stock Esperado', 'Stock Real', 'Discrepancia', 'Unidad', 'Notas']);
      } else if (name === SHEET_NAMES.REPORTED_CLIENTS) {
        newSheet.appendRow(['Fecha Reporte', 'Nº Pedido', 'Nombre Cliente', 'Teléfono', 'Email', 'Nombre Producto', 'Cantidad']);
      } else if (name === SHEET_NAMES.PURCHASE_REQUESTS) {
        newSheet.appendRow(['Timestamp','Cantidad a Solicitar','Producto Base','Formato Adquisición','Cantidad Adquisición','Unidad Adquisición']);
        newSheet.setFrozenRows(1);
      }
    }
  });
}

/**
 * Sets up the IMPORTRANGE formulas in the respective sheets.
 */
function setupImportFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const acquisitionsSheet = ss.getSheetByName(SHEET_NAMES.ACQUISITIONS);
  if (acquisitionsSheet.getRange('A2').getFormula() === '') {
    acquisitionsSheet.getRange('A2').setFormula('=IMPORTRANGE("' + SOURCE_URLS.ACQUISITIONS_SOURCE + '"; "RESUMEN_Adquisiciones!B2:D")');
  }
  if (acquisitionsSheet.getRange('D2').getFormula() === '') {
    acquisitionsSheet.getRange('D2').setFormula('=IMPORTRANGE("' + SOURCE_URLS.ACQUISITIONS_SOURCE + '"; "RESUMEN_Adquisiciones!I2:L")');
  }
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES);
   if (salesSheet.getRange('A1').getFormula() === '') {
      salesSheet.getRange('A1').setFormula('=IMPORTRANGE("' + SOURCE_URLS.OPERACION + '"; "Orders!A:L")');
  }
  const skuSheet = ss.getSheetByName(SHEET_NAMES.SKU);
  if (skuSheet.getRange('A1').getFormula() === '') {
      skuSheet.getRange('A1').setFormula('=IMPORTRANGE("' + SOURCE_URLS.OPERACION + '"; "SKU!A:K")');
  }
}

/**
 * Forces a refresh of all IMPORTRANGE formulas by clearing and resetting them.
 * This function now clears the target sheets to prevent "array result was not expanded" errors.
 */
function forceRefreshAllImports() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Iniciando la actualización forzada de datos. Esto puede tomar unos momentos...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Handle sheets that are fully populated by IMPORTRANGE from A1
    const sheetsToClearFully = [SHEET_NAMES.SALES, SHEET_NAMES.SKU];
    sheetsToClearFully.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const cell = sheet.getRange('A1');
        const formula = cell.getFormula();
        if (formula && formula.toUpperCase().includes('IMPORTRANGE')) {
          sheet.clear(); // Clear the entire sheet
          SpreadsheetApp.flush();
          Utilities.sleep(1000);
          sheet.getRange('A1').setFormula(formula); // Restore formula
        }
      }
    });

    // Handle Acquisitions sheet which has headers in row 1
    const acqSheet = ss.getSheetByName(SHEET_NAMES.ACQUISITIONS);
    if (acqSheet) {
      // Store formulas before clearing
      const formulaA2 = acqSheet.getRange('A2').getFormula();
      const formulaD2 = acqSheet.getRange('D2').getFormula();

      // Clear from row 2 downwards
      if (acqSheet.getMaxRows() > 1) {
        acqSheet.getRange(2, 1, acqSheet.getMaxRows() - 1, acqSheet.getMaxColumns()).clearContent();
      }
      SpreadsheetApp.flush();
      Utilities.sleep(1000);

      // Restore formulas
      if (formulaA2 && formulaA2.toUpperCase().includes('IMPORTRANGE')) {
        acqSheet.getRange('A2').setFormula(formulaA2);
      }
      if (formulaD2 && formulaD2.toUpperCase().includes('IMPORTRANGE')) {
        acqSheet.getRange('D2').setFormula(formulaD2);
      }
    }

    SpreadsheetApp.flush();
    ui.alert('¡Actualización forzada completada! Los datos deberían estar al día.');
  } catch (e) {
    ui.alert('Ocurrió un error durante la actualización: ' + e.message);
  }
}

// Wrapper functions to allow a modal to trigger another modal

function showReportGeneratorUI() {
  const html = HtmlService.createTemplateFromFile('ReportGenerator')
    .evaluate()
    .setWidth(800)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generador de Reportes');
}


function getSalesDataForReporting() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES);
    const salesData = salesSheet.getDataRange().getValues();
    const headers = salesData.shift(); // Remove headers

    const orders = salesData.reduce((acc, row) => {
      const orderId = row[0];
      if (!orderId) return acc;

      if (!acc[orderId]) {
        acc[orderId] = {
          orderId: orderId,
          clientName: row[1],
          email: row[2],
          phone: row[3],
          products: []
        };
      }
      acc[orderId].products.push({
        productName: row[9],
        quantity: row[10]
      });
      return acc;
    }, {});

    return Object.values(orders);
  } catch (e) {
    return { error: e.message };
  }
}

function getReportedClientsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = ss.getSheetByName(SHEET_NAMES.REPORTED_CLIENTS);

    if (!reportSheet) {
      return [];
    }

    const lastRow = reportSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    const data = reportSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    return data.map(row => {
      const reportDate = row[0] instanceof Date ? row[0].toISOString().split('T')[0] : row[0];
      return {
        reportDate: reportDate,
        orderId: row[1],
        clientName: row[2],
        phone: row[3],
        email: row[4],
        productName: row[5],
        quantity: row[6]
      };
    });
  } catch (e) {
    console.error("Error in getReportedClientsData: " + e.message);
    return { error: e.message };
  }
}

function saveReport(reports) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const correctHeaders = ['Fecha Reporte', 'Nº Pedido', 'Nombre Cliente', 'Teléfono', 'Email', 'Nombre Producto', 'Cantidad'];
    let reportSheet = ss.getSheetByName(SHEET_NAMES.REPORTED_CLIENTS);

    if (!reportSheet) {
      reportSheet = ss.insertSheet(SHEET_NAMES.REPORTED_CLIENTS);
      reportSheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders]);
    } else {
      if (reportSheet.getLastRow() === 0) {
        reportSheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders]);
      } else {
        const currentHeaders = reportSheet.getRange(1, 1, 1, correctHeaders.length).getValues()[0];
        if (JSON.stringify(currentHeaders) !== JSON.stringify(correctHeaders)) {
          reportSheet.clear();
          reportSheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders]);
          SpreadsheetApp.flush();
        }
      }
    }

    const lastRow = reportSheet.getLastRow();
    let existingOrderIds = new Set();
    if (lastRow > 1) {
      const orderIds = reportSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(id => id);
      existingOrderIds = new Set(orderIds);
    }

    const newReports = reports.filter(report => !existingOrderIds.has(report.orderId));

    if (newReports.length === 0) {
      return { success: true, message: "Todos los pedidos seleccionados ya habían sido reportados. No se guardó nada nuevo." };
    }

    const reportDate = new Date();
    const rowsToAppend = [];
    newReports.forEach(report => {
      report.products.forEach(p => {
        rowsToAppend.push([
          reportDate,
          report.orderId,
          report.clientName,
          report.phone,
          report.email,
          p.productName,
          p.quantity
        ]);
      });
    });

    reportSheet.getRange(reportSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    const savedCount = newReports.length;
    const skippedCount = reports.length - savedCount;
    let message = `Se guardaron ${savedCount} nuevos reportes.`;
    if (skippedCount > 0) {
      message += ` Se omitieron ${skippedCount} reportes que ya existían.`;
    }

    return { success: true, message: message };
  } catch (e) {
    return { success: false, message: `Error al guardar los reportes: ${e.message}` };
  }
}

function generarInventarioEstimado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const acquisitionsData = ss.getSheetByName(SHEET_NAMES.ACQUISITIONS).getDataRange().getValues().slice(1);
  const salesData = ss.getSheetByName(SHEET_NAMES.SALES).getDataRange().getValues().slice(1);
  const skuData = ss.getSheetByName(SHEET_NAMES.SKU).getDataRange().getValues().slice(1);
  const historicalData = ss.getSheetByName(SHEET_NAMES.HISTORICAL_INVENTORY).getDataRange().getValues().slice(1);

  const converter = new SkuConverter(skuData);
  const acquisitionsByBase = new Map();
  const salesByBase = new Map();
  const salesSummaryByBase = new Map();
  const acquisitionsSummaryByBase = new Map();
  const allBaseProductsSet = new Set(converter.getAllBaseProducts());

  salesData.forEach(row => {
    const nombreProducto = row[9];
    if (!nombreProducto || typeof nombreProducto !== 'string') return;
    const quantitySold = parseFloat(row[10]);
    if (isNaN(quantitySold)) return;

    const saleInfo = converter.getSaleConversion(nombreProducto.trim());
    if (!saleInfo) return;
    const { baseProduct, amountInInventoryUnit } = saleInfo;
    allBaseProductsSet.add(baseProduct);
    const totalSoldInInventoryUnit = quantitySold * amountInInventoryUnit;
    salesByBase.set(baseProduct, (salesByBase.get(baseProduct) || 0) + totalSoldInInventoryUnit);
    if (!salesSummaryByBase.has(baseProduct)) salesSummaryByBase.set(baseProduct, new Map());
    const currentSummaryQty = salesSummaryByBase.get(baseProduct).get(nombreProducto.trim()) || 0;
    salesSummaryByBase.get(baseProduct).set(nombreProducto.trim(), currentSummaryQty + quantitySold);
  });

  acquisitionsData.forEach(row => {
    let [productoBase, formatoCompra, cantComprar, corrCant, corrFormato, corrNCant, corrNUnidad] = row;
    if (!productoBase || typeof productoBase !== 'string') return;
    const baseProductKey = productoBase.trim();
    allBaseProductsSet.add(baseProductKey);

    let totalAcquiredInInventoryUnit = 0;
    let acquisitionFormat = '';
    let acquisitionQty = 0;
    if (corrCant && corrFormato && typeof corrFormato === 'string' && corrNCant && corrNUnidad) {
      acquisitionQty = parseFloat(corrCant);
      const purchaseConversion = converter.getPurchaseConversion(baseProductKey, corrFormato.trim());
      totalAcquiredInInventoryUnit = acquisitionQty * purchaseConversion;
      acquisitionFormat = `${corrCant} ${corrFormato}`;
    } else if (formatoCompra && cantComprar) {
      acquisitionQty = parseFloat(cantComprar);
      const formatoName = SkuConverter.parsePurchaseFormat(formatoCompra);
      const purchaseConversion = converter.getPurchaseConversion(baseProductKey, formatoName);
      totalAcquiredInInventoryUnit = acquisitionQty * purchaseConversion;
      acquisitionFormat = formatoCompra;
    }
    if (totalAcquiredInInventoryUnit > 0) {
      acquisitionsByBase.set(baseProductKey, (acquisitionsByBase.get(baseProductKey) || 0) + totalAcquiredInInventoryUnit);
      if (!acquisitionsSummaryByBase.has(baseProductKey)) acquisitionsSummaryByBase.set(baseProductKey, []);
      acquisitionsSummaryByBase.get(baseProductKey).push({ qty: acquisitionQty, format: acquisitionFormat });
    }
  });

  const latestHistoricalStock = historicalData.reduce((map, row) => {
    let [timestamp, baseProduct, realStock] = row;
    if (baseProduct && typeof baseProduct === 'string' && timestamp) {
      const baseProductKey = baseProduct.trim();
      const ts = new Date(timestamp);
      if (!map.has(baseProductKey) || ts > map.get(baseProductKey).timestamp) {
        map.set(baseProductKey, { stock: parseFloat(realStock) || 0, timestamp: ts });
      }
    }
    return map;
  }, new Map());

  const inventorySheet = ss.getSheetByName(SHEET_NAMES.INVENTARIO_ESTIMADO);

  const inventoryOutput = [];
  allBaseProductsSet.forEach(baseProduct => {
    const lastStockInfo = latestHistoricalStock.get(baseProduct);
    const lastStock = lastStockInfo ? lastStockInfo.stock : 0;
    const lastStockString = lastStockInfo ? `${lastStockInfo.stock.toFixed(2)} (${lastStockInfo.timestamp.toLocaleDateString()})` : 'N/A';
    const acquiredTotal = acquisitionsByBase.get(baseProduct) || 0;
    const soldTotal = salesByBase.get(baseProduct) || 0;
    const expectedStock = lastStock + acquiredTotal - soldTotal;
    const inventoryUnit = converter.getInventoryUnit(baseProduct);

    inventoryOutput.push([
      baseProduct, lastStockString, expectedStock.toFixed(2), inventoryUnit
    ]);
  });

  // Clear previous data and write new estimates
  inventorySheet.getRange(2, 1, inventorySheet.getMaxRows() - 1, inventorySheet.getMaxColumns()).clearContent();
  if (inventoryOutput.length > 0) {
    inventorySheet.getRange(2, 1, inventoryOutput.length, inventoryOutput[0].length).setValues(inventoryOutput);
  }
}


function launchRealInventoryEntry() {
  const html = HtmlService.createTemplateFromFile('RealInventoryEntry.html')
    .evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Inventario Real');
}

function getEstimatedInventoryForModal() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const estimatedSheet = ss.getSheetByName(SHEET_NAMES.INVENTARIO_ESTIMADO);
    if (estimatedSheet.getLastRow() < 2) {
        return [];
    }
    const data = estimatedSheet.getRange(2, 1, estimatedSheet.getLastRow() - 1, 4).getValues();
    return data.map(row => ({
        baseProduct: row[0],
        lastStock: row[1],
        expectedStock: row[2],
        unit: row[3]
    }));
}

function saveRealInventory(inventoryDataFromModal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicalSheet = ss.getSheetByName(SHEET_NAMES.HISTORICAL_INVENTORY);
  const discrepanciesSheet = ss.getSheetByName(SHEET_NAMES.DISCREPANCIES);
  const realInventorySheet = ss.getSheetByName(SHEET_NAMES.INVENTARIO_REAL);
  const timestamp = new Date();

  const rowsToAppendHistorical = [];
  const rowsToAppendDiscrepancy = [];
  const rowsToAppendReal = [];

  inventoryDataFromModal.forEach(item => {
    const baseProduct = item.baseProduct;
    const expectedStock = parseFloat(item.expectedStock);
    const actualStock = parseFloat(item.actualStock);
    const unit = item.unit;
    const note = item.note || '';
    const discrepancy = actualStock - expectedStock;

    // Row for Inventario Real sheet
    rowsToAppendReal.push([timestamp, baseProduct, expectedStock.toFixed(2), actualStock.toFixed(2), discrepancy.toFixed(2), unit, note]);

    // Row for Inventario Histórico sheet
    rowsToAppendHistorical.push([timestamp, baseProduct, actualStock, unit]);

    // Row for Discrepancias sheet (if any)
    if (discrepancy !== 0) {
      rowsToAppendDiscrepancy.push([timestamp, baseProduct, expectedStock.toFixed(2), actualStock.toFixed(2), discrepancy.toFixed(2), unit, note]);
    }
  });

  if (rowsToAppendReal.length > 0) {
    realInventorySheet.getRange(realInventorySheet.getLastRow() + 1, 1, rowsToAppendReal.length, rowsToAppendReal[0].length).setValues(rowsToAppendReal);
  }

  if (rowsToAppendHistorical.length > 0) {
    historicalSheet.getRange(historicalSheet.getLastRow() + 1, 1, rowsToAppendHistorical.length, rowsToAppendHistorical[0].length).setValues(rowsToAppendHistorical);
  }

  if (rowsToAppendDiscrepancy.length > 0) {
    discrepanciesSheet.getRange(discrepanciesSheet.getLastRow() + 1, 1, rowsToAppendDiscrepancy.length, rowsToAppendDiscrepancy[0].length).setValues(rowsToAppendDiscrepancy);
  }

  return { success: true, message: "Inventario real guardado con éxito." };
}

function simulateHistoricalData() {
  const ui = SpreadsheetApp.getUi();
  try {
    synchronizeInventory();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historicalSheet = ss.getSheetByName(SHEET_NAMES.HISTORICAL_INVENTORY);
    const inventorySheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
    const inventoryData = inventorySheet.getDataRange().getValues().slice(1);

    if (inventoryData.length === 0) {
      ui.alert('No hay datos en la hoja "Inventario" para simular. Sincroniza primero.');
      return;
    }

    const historicalRows = [];
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);

    inventoryData.forEach(row => {
      const [baseProduct, , , inventoryUnit] = row;
      const simulatedStock = Math.floor(Math.random() * 10) + 1;
      historicalRows.push([yesterday, baseProduct, simulatedStock, inventoryUnit]);
    });

    if (historicalRows.length > 0) {
      historicalSheet.getRange(historicalSheet.getLastRow() + 1, 1, historicalRows.length, historicalRows[0].length).setValues(historicalRows);
      ui.alert('Se han generado ' + historicalRows.length + ' registros de prueba para "ayer" en "Inventario Histórico".');
    } else {
      ui.alert('No se generaron datos.');
    }
  } catch (e) {
    ui.alert('Error al simular datos: ' + e.message);
  }
}


function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Confirmación de Reinicio Total',
    'Esta acción borrará los datos de inventario y discrepancias. Es irreversible. Por favor, introduce la clave "11" para confirmar.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() == '11') {
    ui.alert('Clave correcta. Iniciando el proceso de reinicio...');
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetsToClear = [SHEET_NAMES.INVENTORY, SHEET_NAMES.HISTORICAL_INVENTORY, SHEET_NAMES.DISCREPANCIES];

      sheetsToClear.forEach(name => {
        const sheet = ss.getSheetByName(name);
        if (sheet) {
          const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns());
          range.clearContent();
        }
      });

      forceRefreshAllImports();

      // Give some time for imports to refresh before syncing
      Utilities.sleep(5000);

      synchronizeInventory();

      ui.alert('¡Reinicio completado! El sistema está listo para un nuevo día.');

    } catch (e) {
      ui.alert('Ocurrió un error durante el reinicio: ' + e.message);
    }
  } else {
    ui.alert('La clave es incorrecta o la operación fue cancelada. No se ha realizado ningún cambio.');
  }
}

/**
 * Cleans a phone number string by removing non-numeric characters.
 * @param {string} phone The phone number to clean.
 * @returns {string} A string containing only digits.
 */
function normalizePhoneNumber(phone) {
  if (!phone || typeof phone.toString !== 'function') {
    return '';
  }
  return phone.toString().replace(/\D/g, '');
}

// --- Funciones para el Dashboard ---

function showDashboard() {
  const html = HtmlService.createTemplateFromFile('dashboard.html')
    .evaluate()
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Inventario');
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get Estimated Inventory Data
  const estimatedSheet = ss.getSheetByName(SHEET_NAMES.INVENTARIO_ESTIMADO);
  const estimatedData = estimatedSheet.getLastRow() > 1 ? estimatedSheet.getRange(2, 1, estimatedSheet.getLastRow() - 1, 4).getValues() : [];
  const inventory = estimatedData.map(row => {
    return {
      baseProduct: row[0],
      lastInventory: row[1],
      expectedStock: row[2],
      unit: row[3]
    };
  });

  // Get Sales Data
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES);
  const salesData = salesSheet.getLastRow() > 1 ? salesSheet.getDataRange().getValues().slice(1) : [];
  const sales = salesData.map(row => {
      return {
          orderId: row[0],
          clientName: row[1],
          productName: row[9],
          quantity: row[10]
      };
  });

  // Get Acquisitions Data
  const acquisitionsSheet = ss.getSheetByName(SHEET_NAMES.ACQUISITIONS);
  const acquisitionsData = acquisitionsSheet.getLastRow() > 1 ? acquisitionsSheet.getDataRange().getValues().slice(1) : [];
  const acquisitions = acquisitionsData.map(row => {
      return {
          baseProduct: row[0],
          format: row[1],
          quantity: row[2]
      };
  });

  return {
    inventory: inventory,
    sales: sales,
    acquisitions: acquisitions
  };
}

// --- Fin de funciones para el nuevo Dashboard v2 ---

class SkuConverter {
  constructor(skuData) {
    this.inventoryUnitMap = new Map(); // Maps Producto Base -> Unidad Venta
    this.salesConversionMap = new Map(); // Maps Nombre Producto -> { baseProduct, amountInInventoryUnit }
    this.purchaseConversionMap = new Map(); // Maps Producto Base -> Formato -> amountInInventoryUnit
    this.categoryMap = new Map(); // Maps Producto Base -> Categoria
    this.baseProductToSkuMap = new Map(); // Maps Producto Base -> Nombre Producto (SKU)

    this._buildMaps(skuData);
  }

  _buildMaps(skuData) {
    // Pass 1: Determine the standard inventory unit, Category, and SKU for each Producto Base.
    skuData.forEach(row => {
      const [nombreProducto, productoBase, , , , categoria, , unidadVenta] = row;
      if (productoBase && typeof productoBase === 'string') {
        const baseProductKey = productoBase.trim();
        if (unidadVenta && typeof unidadVenta === 'string' && !this.inventoryUnitMap.has(baseProductKey)) {
          this.inventoryUnitMap.set(baseProductKey, unidadVenta.trim());
        }
        if (categoria && typeof categoria === 'string' && !this.categoryMap.has(baseProductKey)) {
          this.categoryMap.set(baseProductKey, categoria.trim());
        }
        if (nombreProducto && typeof nombreProducto === 'string' && !this.baseProductToSkuMap.has(baseProductKey)) {
          this.baseProductToSkuMap.set(baseProductKey, nombreProducto.trim());
        }
      }
    });

    // Pass 2: Build the conversion maps.
    skuData.forEach(row => {
      const [nombreProducto, productoBase, formato, cantCompra, unidadCompra, , cantVenta, unidadVenta] = row;

      // Sales Conversion
      if (nombreProducto && typeof nombreProducto === 'string' && productoBase && typeof productoBase === 'string' && cantVenta && unidadVenta) {
        const baseProductKey = productoBase.trim();
        const inventoryUnit = this.getInventoryUnit(baseProductKey);
        const convertedAmount = this.convert(parseFloat(cantVenta), unidadVenta, inventoryUnit);
        this.salesConversionMap.set(nombreProducto.trim(), {
          baseProduct: baseProductKey,
          amountInInventoryUnit: convertedAmount
        });
      }

      // Purchase Conversion
      if (productoBase && typeof productoBase === 'string' && formato && typeof formato === 'string' && cantCompra && unidadCompra) {
        const baseProductKey = productoBase.trim();
        if (!this.purchaseConversionMap.has(baseProductKey)) {
          this.purchaseConversionMap.set(baseProductKey, new Map());
        }
        const inventoryUnit = this.getInventoryUnit(baseProductKey);
        const convertedAmount = this.convert(parseFloat(cantCompra), unidadCompra, inventoryUnit);
        this.purchaseConversionMap.get(baseProductKey).set(formato.trim(), convertedAmount);
      }
    });
  }

  getInventoryUnit(baseProduct) {
    return this.inventoryUnitMap.get(baseProduct) || 'Unidad';
  }

  getCategory(baseProduct) {
    return this.categoryMap.get(baseProduct) || null;
  }

  getSku(baseProduct) {
    return this.baseProductToSkuMap.get(baseProduct) || null;
  }

  getSaleConversion(nombreProducto) {
    return this.salesConversionMap.get(nombreProducto) || null;
  }

  getPurchaseConversion(productoBase, formato) {
    if (this.purchaseConversionMap.has(productoBase)) {
      return this.purchaseConversionMap.get(productoBase).get(formato) || 0;
    }
    return 0;
  }

  getAllBaseProducts() {
    return [...this.inventoryUnitMap.keys()];
  }

  convert(amount, fromUnit, toUnit) {
    if (!fromUnit || !toUnit || fromUnit.toLowerCase() === toUnit.toLowerCase()) {
      return amount;
    }
    const unitMap = { 'kilo': 1000, 'kg': 1000, 'grs': 1, 'gr': 1, 'g': 1 };
    const from = fromUnit.toLowerCase();
    const to = toUnit.toLowerCase();
    if (unitMap[from] && unitMap[to]) {
      const amountInGrams = amount * unitMap[from];
      return amountInGrams / unitMap[to];
    }
    return amount;
  }

  static parsePurchaseFormat(formatString) {
    if (!formatString || typeof formatString !== 'string') return null;
    const match = formatString.match(/([^\(]+)/);
    if (match) {
      return match[1].trim();
    }
    return formatString.trim();
  }
}

// --- Funciones para la Solicitud de Compra ---

function showPurchaseRequestUI() {
  const html = HtmlService.createTemplateFromFile('PurchaseRequest')
    .evaluate()
    .setWidth(500)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nueva Solicitud de Compra');
}

function getBaseProductSuggestions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skuSheet = ss.getSheetByName(SHEET_NAMES.SKU);
    if (!skuSheet) {
      Logger.log('La hoja SKU no fue encontrada.');
      return [];
    }
    const skuData = skuSheet.getDataRange().getValues().slice(1); // Get all data, skip header
    const converter = new SkuConverter(skuData);
    const suggestions = converter.getAllBaseProducts();
    return suggestions;
  } catch (e) {
    Logger.log('Error in getBaseProductSuggestions: ' + e.message);
    return [];
  }
}

function savePurchaseRequest(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reqSheet = ss.getSheetByName(SHEET_NAMES.PURCHASE_REQUESTS);
    if (!reqSheet) {
      throw new Error(`Sheet "${SHEET_NAMES.PURCHASE_REQUESTS}" not found.`);
    }
    const newRow = [
      new Date(),
      data.cantidad,
      data.productoBase,
      data.formato,
      data.cantidadAdquisicion,
      data.unidad
    ];
    reqSheet.appendRow(newRow);
    return { success: true };
  } catch (e) {
    Logger.log('Error in savePurchaseRequest: ' + e.message);
    throw new Error('Failed to save request. ' + e.message);
  }
}
