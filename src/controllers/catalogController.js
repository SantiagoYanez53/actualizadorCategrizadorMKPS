const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// ============================================================================
// 1. DEFINICIÓN DE COORDENADAS (Índices basados en 1)
// ============================================================================

const MAESTRO_COL = {
    SKU: 1, UPC: 2, TITULO: 3, DESCRIPCION: 4, PRECIO: 5,
    MEDIDAS: 6, ANCHO: 7, ALTO: 8, LARGO: 9, PESO: 10,
    IMG1: 11, IMG2: 12, IMG3: 13, IMG4: 14, IMG5: 15, MARCA: 16
};

const COPPEL_CAT = {
    CATEGORIA: 1, SKU: 2, TITULO: 3, UPC: 4, MARCA: 5,
    MODELO: 6, COLOR: 7, DESCRIPCION: 9, ORIGEN: 10, MATERIAL: 11,
    MEDIDAS: 12, PESO: 13, IMG_START: 15, OFERTA_SKU: 24, OFERTA_UPC: 25,
    PRECIO: 29, CANTIDAD: 31, PESO_PAQ: 42, ALTO_PAQ: 43, LARGO_PAQ: 44,
    ANCHO_PAQ: 45, SAT: 46
};

const WALMART_CONFIGS = {
    'accesorios de instrumentos': {
        SKU: 4, UPC: 6, TITULO: 7, MARCA: 8, IMG_MAIN: 9,
        FEATURES: 10, DESC: 11, PRECIO: 12, FABRICANTE: 13, SAT: 14,
        GARANTIA: 24,       // X: "0"
        ORIGEN: 28,         // AB: "CN- China"
        CONTENIDO: 29,      // AC: Título
        COLOR: 32,          // AF: "Multicolor"
        IMG_SEC_START: 34,  // AH (Fotos en AH, AI, AJ, AK)
        HAZARD: 88,         // CJ: "No"
        NOM: 90             // CL: "Sí"
    },
    'instrumentos musicales': {
        SKU: 4, UPC: 6, TITULO: 7, MARCA: 8, IMG_MAIN: 9,
        FEATURES: 10, DESC: 11, PRECIO: 12, FABRICANTE: 13, SAT: 14,
        GARANTIA: 24,       // X: "0"
        ORIGEN: 28,         // AB: "CN- China"
        CONTENIDO: 29,      // AC: Título
        MATERIAL: 31,       // AE: Condicional (Madera / Plástico)
        CANT_AN: 40,        // AN: "1"
        IMG_SEC_START: 41,  // AO (Fotos en AO, AP, AQ, AR)
        HAZARD: 108,        // DD: "No"
        NOM: 110            // DF: "Sí"
    },
    'amplificadores y audio': { 
        SKU: 4, UPC: 6, TITULO: 7, MARCA: 8, IMG_MAIN: 9,
        FEATURES: 10, DESC: 11, PRECIO: 12, FABRICANTE: 13, SAT: 14,
        GARANTIA_AA: 27,    // AA: "1"
        ORIGEN: 28,         // AB: "CN- China"
        CONTENIDO: 29,      // AC: Título
        CANT_AE: 31,        // AE: "1"
        COLOR: 32,          // AF: "Multicolor"
        IMG_SEC_START: 42,  // AP (Fotos en AP, AQ, AR, AS)
        HAZARD: 94,         // CP: "No"
        NOM: 96             // CR: "Sí"
    }
};

// ============================================================================
// 2. FUNCIÓN PRINCIPAL DEL CONTROLADOR
// ============================================================================

async function generarCatalogos(req, res) {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Sube tu archivo maestro de nuevos productos.' });
        }

        const rutaMaestro = req.file.path;
        
        // Rutas estandarizadas a las plantillas limpias
        const rutaPlantillaCoppel = path.join(__dirname, '../../templates/plantilla_coppel_catalogo.xlsx');
        const rutaPlantillaWalmart = path.join(__dirname, '../../templates/plantilla_walmart_catalogo.xlsx');
        
        // Validamos físicamente que las plantillas existan antes de leerlas
        if (!fs.existsSync(rutaPlantillaCoppel) || !fs.existsSync(rutaPlantillaWalmart)) {
            fs.unlinkSync(rutaMaestro); // Limpiamos el temporal
            return res.status(500).json({ 
                error: 'Faltan archivos base. Asegúrate de tener plantilla_coppel_catalogo.xlsx y plantilla_walmart_catalogo.xlsx en la carpeta templates/' 
            });
        }
        

        // Cargamos el diccionario de categorías blindando posibles errores de formato
        const rutaDiccionario = path.join(__dirname, '../config/diccionario.json');
        let diccionario = [];
        if (fs.existsSync(rutaDiccionario)) {
            const contenidoJson = fs.readFileSync(rutaDiccionario, 'utf-8');
            // Solo intentamos parsear si el archivo no está completamente en blanco
            if (contenidoJson.trim() !== '') {
                diccionario = JSON.parse(contenidoJson);
            }
        }

        const libroMaestro = new ExcelJS.Workbook();
        const libroCoppel = new ExcelJS.Workbook();
        const libroWalmart = new ExcelJS.Workbook();

        console.log("➡️ 1. Empezando a leer el archivo Maestro...");
        await libroMaestro.xlsx.readFile(rutaMaestro);
        console.log("✅ Maestro leído correctamente.");

        console.log("➡️ 2. Empezando a leer plantilla Coppel...");
        await libroCoppel.xlsx.readFile(rutaPlantillaCoppel);
        console.log("✅ Plantilla Coppel leída.");

        console.log("➡️ 3. Empezando a leer plantilla Walmart...");
        await libroWalmart.xlsx.readFile(rutaPlantillaWalmart);
        console.log("✅ Plantilla Walmart leída.");

        const hojaMaestro = libroMaestro.worksheets[0];
        
        // CORRECCIÓN 1: Selección robusta de la hoja destino en Coppel (Siempre la última pestaña de datos)
        let hojaCoppel = libroCoppel.worksheets[libroCoppel.worksheets.length - 1];

        let procesados = 0;

        // ====================================================================
        // 3. RECORRIDO DEL ARCHIVO MAESTRO 
        // ====================================================================
        hojaMaestro.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { 
                const sku = row.getCell(MAESTRO_COL.SKU).text || row.getCell(MAESTRO_COL.SKU).value;
                const upc = row.getCell(MAESTRO_COL.UPC).text || row.getCell(MAESTRO_COL.UPC).value;
                const titulo = row.getCell(MAESTRO_COL.TITULO).value;

                if (sku && titulo) {
                    const skuLimpio = String(sku).trim();
                    const upcLimpio = upc ? String(upc).trim() : '';
                    const tituloLimpio = String(titulo).trim();
                    const desc = row.getCell(MAESTRO_COL.DESCRIPCION).value || '';
                    const precio = row.getCell(MAESTRO_COL.PRECIO).value || 0;
                    const medidas = row.getCell(MAESTRO_COL.MEDIDAS).value || '';
                    const ancho = row.getCell(MAESTRO_COL.ANCHO).value || '';
                    const alto = row.getCell(MAESTRO_COL.ALTO).value || '';
                    const largo = row.getCell(MAESTRO_COL.LARGO).value || '';
                    const peso = row.getCell(MAESTRO_COL.PESO).value || '';
                    const marca = row.getCell(MAESTRO_COL.MARCA).value || 'S/M';

                    const img1 = row.getCell(MAESTRO_COL.IMG1).text || '';
                    const img2 = row.getCell(MAESTRO_COL.IMG2).text || '';
                    const img3 = row.getCell(MAESTRO_COL.IMG3).text || '';
                    const img4 = row.getCell(MAESTRO_COL.IMG4).text || '';
                    const img5 = row.getCell(MAESTRO_COL.IMG5).text || '';

                    let catCoppel = 'Por Asignar';
                    let catWalmart = 'Por Asignar';
                    let pestanaWalmart = 'Instrumentos Musicales'; 
                    let codigoSAT = '01010101'; 

                    const tituloLower = tituloLimpio.toLowerCase();
                    for (const item of diccionario) {
                        if (item.palabras && Array.isArray(item.palabras)) {
                            const coincide = item.palabras.some(palabra => tituloLower.includes(palabra.toLowerCase()));
                            if (coincide) {
                                catCoppel = item.coppel || catCoppel;
                                catWalmart = item.walmart || catWalmart;
                                pestanaWalmart = item.pestana_walmart || pestanaWalmart;
                                codigoSAT = item.sat || codigoSAT;
                                break;
                            }
                        }
                    }

                    // ========================================================
                    // 4. INYECCIÓN EN COPPEL
                    // ========================================================
                    const nuevaFilaCoppel = hojaCoppel.addRow([]);
                    
                    const celdaSkuCpl = nuevaFilaCoppel.getCell(COPPEL_CAT.SKU);
                    celdaSkuCpl.value = skuLimpio;
                    celdaSkuCpl.numFmt = '@';

                    const celdaUpcCpl = nuevaFilaCoppel.getCell(COPPEL_CAT.UPC);
                    celdaUpcCpl.value = upcLimpio;
                    celdaUpcCpl.numFmt = '@';

                    const celdaModCpl = nuevaFilaCoppel.getCell(COPPEL_CAT.MODELO);
                    celdaModCpl.value = skuLimpio;
                    celdaModCpl.numFmt = '@';

                    const celdaOfSkuCpl = nuevaFilaCoppel.getCell(COPPEL_CAT.OFERTA_SKU);
                    celdaOfSkuCpl.value = skuLimpio;
                    celdaOfSkuCpl.numFmt = '@';

                    const celdaOfUpcCpl = nuevaFilaCoppel.getCell(COPPEL_CAT.OFERTA_UPC);
                    celdaOfUpcCpl.value = upcLimpio;
                    celdaOfUpcCpl.numFmt = '@';

                    nuevaFilaCoppel.getCell(COPPEL_CAT.CATEGORIA).value = catCoppel;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.TITULO).value = tituloLimpio;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.MARCA).value = marca;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.COLOR).value = 'Multicolor';
                    nuevaFilaCoppel.getCell(COPPEL_CAT.DESCRIPCION).value = desc;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.ORIGEN).value = 'China';
                    nuevaFilaCoppel.getCell(COPPEL_CAT.MEDIDAS).value = medidas;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.PESO).value = peso;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.PRECIO).value = precio;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.CANTIDAD).value = 1;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.PESO_PAQ).value = peso;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.ALTO_PAQ).value = alto;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.LARGO_PAQ).value = largo;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.ANCHO_PAQ).value = ancho;
                    nuevaFilaCoppel.getCell(COPPEL_CAT.SAT).value = codigoSAT;

                    if (img1) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START).value = img1;
                    if (img2) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 1).value = img2;
                    if (img3) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 2).value = img3;
                    if (img4) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 3).value = img4;
                    if (img5) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 4).value = img5;

                    // ========================================================
                    // 5. INYECCIÓN EN WALMART
                    // ========================================================
                    // CORRECCIÓN 2: Búsqueda flexible de pestañas ignorando mayúsculas y espacios
                    let hojaWalmart;
                    const pestanaBuscada = pestanaWalmart.trim().toLowerCase();
                    
                    libroWalmart.worksheets.forEach(ws => {
                        if (ws.name.trim().toLowerCase() === pestanaBuscada) {
                            hojaWalmart = ws;
                        }
                    });

                    // Respaldo por si hay un error tipográfico en el JSON
                    if (!hojaWalmart) {
                        hojaWalmart = libroWalmart.worksheets[libroWalmart.worksheets.length - 1];
                    }

                    const nuevaFilaWmt = hojaWalmart.addRow([]);

                    // Obtenemos el mapa de columnas de la pestaña, por defecto 'accesorios de instrumentos'
                    const conf = WALMART_CONFIGS[pestanaBuscada] || WALMART_CONFIGS['accesorios de instrumentos'];

                    // Blindaje SKU y UPC
                    if (conf.SKU) {
                        const celdaSkuWmt = nuevaFilaWmt.getCell(conf.SKU);
                        celdaSkuWmt.value = skuLimpio;
                        celdaSkuWmt.numFmt = '@';
                    }
                    if (conf.UPC) {
                        const celdaUpcWmt = nuevaFilaWmt.getCell(conf.UPC);
                        celdaUpcWmt.value = upcLimpio;
                        celdaUpcWmt.numFmt = '@';
                    }

                    // Datos generales (Idénticos para las 3 pestañas)
                    if (conf.TITULO) nuevaFilaWmt.getCell(conf.TITULO).value = tituloLimpio;
                    if (conf.MARCA) nuevaFilaWmt.getCell(conf.MARCA).value = marca;
                    if (conf.IMG_MAIN) nuevaFilaWmt.getCell(conf.IMG_MAIN).value = img1;
                    if (conf.FEATURES) nuevaFilaWmt.getCell(conf.FEATURES).value = desc;
                    if (conf.DESC) nuevaFilaWmt.getCell(conf.DESC).value = desc;
                    if (conf.PRECIO) nuevaFilaWmt.getCell(conf.PRECIO).value = precio;
                    if (conf.FABRICANTE) nuevaFilaWmt.getCell(conf.FABRICANTE).value = marca;
                    if (conf.SAT) nuevaFilaWmt.getCell(conf.SAT).value = codigoSAT;
                    if (conf.ORIGEN) nuevaFilaWmt.getCell(conf.ORIGEN).value = 'CN- China';
                    if (conf.CONTENIDO) nuevaFilaWmt.getCell(conf.CONTENIDO).value = tituloLimpio;
                    if (conf.HAZARD) nuevaFilaWmt.getCell(conf.HAZARD).value = 'No';
                    if (conf.NOM) nuevaFilaWmt.getCell(conf.NOM).value = 'Sí';

                    // Valores que varían según la pestaña
                    if (conf.COLOR) nuevaFilaWmt.getCell(conf.COLOR).value = 'Multicolor';
                    if (conf.GARANTIA) nuevaFilaWmt.getCell(conf.GARANTIA).value = '0';
                    if (conf.GARANTIA_AA) nuevaFilaWmt.getCell(conf.GARANTIA_AA).value = '1';
                    if (conf.CANT_AE) nuevaFilaWmt.getCell(conf.CANT_AE).value = '1';
                    if (conf.CANT_AN) nuevaFilaWmt.getCell(conf.CANT_AN).value = '1';

                    // 🤖 Condicional Inteligente para el Material (Solo en Instrumentos Musicales)
                    if (conf.MATERIAL) {
                        const regexMadera = /(batería|bateria|guitarra|bajo)/i;
                        if (regexMadera.test(tituloLimpio)) {
                            nuevaFilaWmt.getCell(conf.MATERIAL).value = 'Madera';
                        } else {
                            nuevaFilaWmt.getCell(conf.MATERIAL).value = 'Plástico';
                        }
                    }

                    // Inyección de Imágenes Secundarias desplazables
                    if (conf.IMG_SEC_START) {
                        if (img2) nuevaFilaWmt.getCell(conf.IMG_SEC_START).value = img2;
                        if (img3) nuevaFilaWmt.getCell(conf.IMG_SEC_START + 1).value = img3;
                        if (img4) nuevaFilaWmt.getCell(conf.IMG_SEC_START + 2).value = img4;
                        if (img5) nuevaFilaWmt.getCell(conf.IMG_SEC_START + 3).value = img5;
                    }

                    procesados++;
                }
            }
        });

        // ====================================================================
        // 6. GUARDAR ARCHIVOS
        // ====================================================================
        const timestamp = Date.now();
        const nombreArchivoCoppel = `Catalogo_Coppel_Listo_${timestamp}.xlsx`;
        const nombreArchivoWalmart = `Catalogo_Walmart_Listo_${timestamp}.xlsx`;

        const carpetaUploads = path.join(__dirname, '../../uploads/');
        if (!fs.existsSync(carpetaUploads)) {
            fs.mkdirSync(carpetaUploads, { recursive: true });
        }

        const rutaSalidaCoppel = path.join(carpetaUploads, nombreArchivoCoppel);
        const rutaSalidaWalmart = path.join(carpetaUploads, nombreArchivoWalmart);

        await libroCoppel.xlsx.writeFile(rutaSalidaCoppel);
        await libroWalmart.xlsx.writeFile(rutaSalidaWalmart);

        fs.unlinkSync(rutaMaestro);

        return res.json({
            mensaje: 'Catálogos generados con éxito',
            procesados: procesados,
            archivos: {
                coppel: nombreArchivoCoppel,
                walmart: nombreArchivoWalmart
            }
        });

    } catch (error) {
        console.error('Error crítico en el controlador de catálogos:', error);
        // Si falla en medio del proceso, intentamos borrar el archivo maestro residual
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        return res.status(500).json({ error: `Falla interna del servidor: ${error.message}` });
    }
}

module.exports = {
    generarCatalogos
};