const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
// =====================================================================
// ⚙️ CONFIGURACIÓN DE COLUMNAS (Índices numéricos: A=1, B=2, C=3... P=16)
// =====================================================================

// 1. Tu Archivo Maestro
const COL_MAESTRO_SKU = 2;    // Columna B
const COL_MAESTRO_STOCK = 16; // Columna P

// 2. Plantilla Coppel (.csv)
// ¡OJO! Abre tu CSV de Coppel y cambia estos números según tus columnas
const COL_COPPEL_SKU = 1;     // <-- Cambiar por la columna del SKU en Coppel
const COL_COPPEL_STOCK = 2;   // <-- Cambiar por la columna de Cantidad en Coppel

// 3. Plantilla Walmart (.xlsx)
// ¡OJO! Abre tu Excel de Walmart y cambia estos números según tus columnas
const COL_WALMART_SKU = 1;    // <-- Cambiar por la columna del SKU en Walmart
const COL_WALMART_STOCK = 2;  // <-- Cambiar por la columna de Cantidad en Walmart

// =====================================================================

const procesarActualizacionStock = async (req, res) => {
    try {
        console.log('⏳ Iniciando extracción y cruce de inventario...');
        
        // -------------------------------------------------------------
        // PASO 1: LEER TU ARCHIVO MAESTRO Y GUARDARLO EN MEMORIA
        // -------------------------------------------------------------
        const rutaMaestro = path.join(__dirname, '../../uploads/', req.file.filename);
        const libroMaestro = new ExcelJS.Workbook();
        await libroMaestro.xlsx.readFile(rutaMaestro);
        const hojaMaestro = libroMaestro.worksheets[0];

        const mapaStock = {}; // Diccionario ultra rápido para buscar SKUs
        let skusLeidos = 0;

        hojaMaestro.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Saltamos los encabezados (fila 1)
                let sku = row.getCell(COL_MAESTRO_SKU).value;
                let stock = row.getCell(COL_MAESTRO_STOCK).value;
                const celdaSKU = row.getCell(COL_MAESTRO_SKU);

                // Validamos que la celda no esté vacía
                if (sku !== null && sku !== undefined) {
                    // Limpiamos espacios accidentales y aseguramos que sea texto
                    sku = String(sku).trim(); 
                    // Aseguramos que el stock sea un número entero (si está vacío, ponemos 0)
                    stock = parseInt(stock) || 0; 
                    
                    mapaStock[sku] = stock;
                    skusLeidos++;
                }
            }
        });
        
        console.log(`✅ Maestro procesado: ${skusLeidos} SKUs guardados en memoria.`);

        // -------------------------------------------------------------
        // PASO 2: PROCESAR PLANTILLA DE COPPEL (Formato CSV)
        // -------------------------------------------------------------
        const rutaPlantillaCoppel = path.join(__dirname, '../../templates/plantilla_coppel_stock.csv');
        const libroCoppel = new ExcelJS.Workbook();
        
        // Leemos el CSV
        await libroCoppel.csv.readFile(rutaPlantillaCoppel);
        const hojaCoppel = libroCoppel.getWorksheet(1);

       let actualizadosCoppel = 0;
        hojaCoppel.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const celdaSku = row.getCell(COL_COPPEL_SKU)
                let skuCoppel = row.getCell(COL_COPPEL_SKU).value;
                
                if (skuCoppel !== null && skuCoppel !== undefined) {
                    skuCoppel = String(skuCoppel).trim();
                    
                    // BUSCAMOS: Si este SKU existe en nuestro diccionario maestro...
                    if (mapaStock[skuCoppel] !== undefined) {
                        row.getCell(COL_COPPEL_STOCK).value = mapaStock[skuCoppel];
                        actualizadosCoppel++;
                    } else {
                        // OJO: Si el SKU no viene en el maestro, forzamos el stock a 0
                        row.getCell(COL_COPPEL_STOCK).value = 0;
                    }
                }
            }
        });

        // Guardamos el nuevo archivo CSV procesado en la carpeta uploads
        const archivoSalidaCoppel = `Coppel_Stock_Listo_${Date.now()}.csv`;
        const rutaSalidaCoppel = path.join(__dirname, '../../uploads/', archivoSalidaCoppel);
        await libroCoppel.csv.writeFile(rutaSalidaCoppel);
        
        console.log(`✅ Coppel listo: ${actualizadosCoppel} SKUs actualizados.`);

        // -------------------------------------------------------------
        // PASO 3: PROCESAR PLANTILLA DE WALMART (Formato XLSX)
        // -------------------------------------------------------------
        const rutaPlantillaWalmart = path.join(__dirname, '../../templates/plantilla_walmart_stock.xlsx');
        const libroWalmart = new ExcelJS.Workbook();
        
        // Leemos el Excel
        await libroWalmart.xlsx.readFile(rutaPlantillaWalmart);
        const hojaWalmart = libroWalmart.worksheets[0];

        let actualizadosWalmart = 0;
        hojaWalmart.eachRow((row, rowNumber) => {
            if (rowNumber > 2) {
                const celdaSKU = row.getCell(COL_WALMART_SKU).value;
                let skuWalmart = row.getCell(COL_WALMART_SKU).value;
                
                if (skuWalmart !== null && skuWalmart !== undefined) {
                    celdaSKU.numFmt = '@';
                    skuWalmart = String(skuWalmart).trim();
                    
                    // BUSCAMOS en el maestro
                    if (mapaStock[skuWalmart] !== undefined) {
                        row.getCell(COL_WALMART_STOCK).value = mapaStock[skuWalmart];
                        actualizadosWalmart++;
                    } else {
                        // OJO: Si el SKU no viene en el maestro, forzamos el stock a 0
                        row.getCell(COL_WALMART_STOCK).value = 0;
                    }
                }
            }
        });

        // Guardamos el nuevo Excel procesado
        const archivoSalidaWalmart = `Walmart_Stock_Listo_${Date.now()}.xlsx`;
        const rutaSalidaWalmart = path.join(__dirname, '../../uploads/', archivoSalidaWalmart);
        await libroWalmart.xlsx.writeFile(rutaSalidaWalmart);
        
        console.log(`✅ Walmart listo: ${actualizadosWalmart} SKUs actualizados.`);

        // -------------------------------------------------------------
        // PASO 4: RESPONDER AL FRONTEND CON LOS RESULTADOS
        // -------------------------------------------------------------
        fs.unlinkSync(rutaMaestro);
        console.log(`✅ Archivo maestro limpiado: ${rutaMaestro}`);
        
        
        res.json({
            mensaje: '¡Cruce de inventario completado con éxito!',
            resumen: {
                total_skus_maestro: skusLeidos,
                actualizados_coppel: actualizadosCoppel,
                actualizados_walmart: actualizadosWalmart
            },
            archivos_generados: {
                coppel: archivoSalidaCoppel,
                walmart: archivoSalidaWalmart
            }
        });

    } catch (error) {
        console.error('Error crítico en el cruce de stock:', error);
        res.status(500).json({ 
            error: 'Falló el procesamiento. Asegúrate de tener plantilla_coppel_stock.csv y plantilla_walmart_stock.xlsx dentro de la carpeta templates/.' 
        });
    }
};

module.exports = {
    procesarActualizacionStock
};