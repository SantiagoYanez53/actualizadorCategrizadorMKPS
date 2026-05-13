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

const WALMART_COL = {
    SKU: 4, UPC: 6, TITULO: 7, MARCA: 8, IMG_MAIN: 9,
    FEATURES: 10, DESC: 11, PRECIO: 12, FABRICANTE: 13, SAT: 14,
    GARANTIA: 24, ORIGEN: 28, CONTENIDO: 29, COLOR: 32, IMG_SEC_START: 34,
    HAZARD: 88, NOM: 90
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
        
        // Rutas de plantillas limpias (Asegúrate de tenerlas en tu carpeta templates/)
        const rutaPlantillaCoppel = path.join(__dirname, '../../templates/plantilla_coppel_catalogo.xlsx');
        const rutaPlantillaWalmart = path.join(__dirname, '../../templates/plantilla_walmart_catalogo.xlsx');
        
        // Cargamos el diccionario de categorías (generado previamente de tu .xlsm)
        const rutaDiccionario = path.join(__dirname, '../config/diccionario.json');
        let diccionario = [];
        if (fs.existsSync(rutaDiccionario)) {
            diccionario = JSON.parse(fs.readFileSync(rutaDiccionario, 'utf-8'));
        }

        // Instanciamos los tres libros
        const libroMaestro = new ExcelJS.Workbook();
        const libroCoppel = new ExcelJS.Workbook();
        const libroWalmart = new ExcelJS.Workbook();

        await libroMaestro.xlsx.readFile(rutaMaestro);
        await libroCoppel.xlsx.readFile(rutaPlantillaCoppel);
        await libroWalmart.xlsx.readFile(rutaPlantillaWalmart);

        const hojaMaestro = libroMaestro.worksheets[0];
        const hojaCoppel = libroCoppel.worksheets[0]; // Coppel usa una sola hoja

        let procesados = 0;

        // ====================================================================
        // 3. RECORRIDO DEL ARCHIVO MAESTRO
        // ====================================================================
        hojaMaestro.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Saltamos encabezados
                // Leemos texto puro para proteger ceros iniciales
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

                    // Extraemos URLs de imágenes
                    const img1 = row.getCell(MAESTRO_COL.IMG1).text || '';
                    const img2 = row.getCell(MAESTRO_COL.IMG2).text || '';
                    const img3 = row.getCell(MAESTRO_COL.IMG3).text || '';
                    const img4 = row.getCell(MAESTRO_COL.IMG4).text || '';
                    const img5 = row.getCell(MAESTRO_COL.IMG5).text || '';

                    // --- MOTOR DE BÚSQUEDA DE CATEGORÍA ---
                    let catCoppel = 'Por Asignar';
                    let catWalmart = 'Por Asignar';
                    let pestanaWalmart = 'Instrumentos Musicales'; // Pestaña por defecto
                    let codigoSAT = '01010101'; // SAT por defecto

                    // Buscamos coincidencia de palabras clave en el título
                    const tituloLower = tituloLimpio.toLowerCase();
                    for (const item of diccionario) {
                        // Suponiendo que tu JSON tiene: { palabras: ["bateria", "acustica"], coppel: "...", walmart: "...", pestana_walmart: "...", sat: "..." }
                        const coincide = item.palabras.some(palabra => tituloLower.includes(palabra.toLowerCase()));
                        if (coincide) {
                            catCoppel = item.coppel || catCoppel;
                            catWalmart = item.walmart || catWalmart;
                            pestanaWalmart = item.pestana_walmart || pestanaWalmart;
                            codigoSAT = item.sat || codigoSAT;
                            break;
                        }
                    }

                    // ========================================================
                    // 4. INYECCIÓN EN COPPEL
                    // ========================================================
                    const nuevaFilaCoppel = hojaCoppel.addRow([]);
                    
                    // Asignar strings puros y forzar formato texto (@)
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

                    // Datos generales Coppel
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

                    // Imágenes Coppel
                    if (img1) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START).value = img1;
                    if (img2) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 1).value = img2;
                    if (img3) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 2).value = img3;
                    if (img4) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 3).value = img4;
                    if (img5) nuevaFilaCoppel.getCell(COPPEL_CAT.IMG_START + 4).value = img5;


                    // ========================================================
                    // 5. INYECCIÓN EN WALMART (Pestaña Dinámica)
                    // ========================================================
                    // Intentamos obtener la pestaña mapeada, si no existe usamos la primera
                    let hojaWalmart = libroWalmart.getWorksheet(pestanaWalmart);
                    if (!hojaWalmart) {
                        hojaWalmart = libroWalmart.worksheets[0];
                    }

                    const nuevaFilaWmt = hojaWalmart.addRow([]);

                    // Blindaje SKU y UPC Walmart
                    const celdaSkuWmt = nuevaFilaWmt.getCell(WALMART_COL.SKU);
                    celdaSkuWmt.value = skuLimpio;
                    celdaSkuWmt.numFmt = '@';

                    const celdaUpcWmt = nuevaFilaWmt.getCell(WALMART_COL.UPC);
                    celdaUpcWmt.value = upcLimpio;
                    celdaUpcWmt.numFmt = '@';

                    // Datos generales Walmart
                    nuevaFilaWmt.getCell(WALMART_COL.TITULO).value = tituloLimpio;
                    nuevaFilaWmt.getCell(WALMART_COL.MARCA).value = marca;
                    nuevaFilaWmt.getCell(WALMART_COL.IMG_MAIN).value = img1;
                    nuevaFilaWmt.getCell(WALMART_COL.FEATURES).value = desc;
                    nuevaFilaWmt.getCell(WALMART_COL.DESC).value = desc;
                    nuevaFilaWmt.getCell(WALMART_COL.PRECIO).value = precio;
                    nuevaFilaWmt.getCell(WALMART_COL.FABRICANTE).value = marca;
                    nuevaFilaWmt.getCell(WALMART_COL.SAT).value = codigoSAT;
                    
                    // Fijos y repetidos Walmart
                    nuevaFilaWmt.getCell(WALMART_COL.GARANTIA).value = '0';
                    nuevaFilaWmt.getCell(WALMART_COL.ORIGEN).value = 'CN- China';
                    nuevaFilaWmt.getCell(WALMART_COL.CONTENIDO).value = tituloLimpio;
                    nuevaFilaWmt.getCell(WALMART_COL.COLOR).value = 'Multicolor';
                    nuevaFilaWmt.getCell(WALMART_COL.HAZARD).value = 'No';
                    nuevaFilaWmt.getCell(WALMART_COL.NOM).value = 'Sí';

                    // Imágenes secundarias Walmart (AH a AK)
                    if (img2) nuevaFilaWmt.getCell(WALMART_COL.IMG_SEC_START).value = img2;
                    if (img3) nuevaFilaWmt.getCell(WALMART_COL.IMG_SEC_START + 1).value = img3;
                    if (img4) nuevaFilaWmt.getCell(WALMART_COL.IMG_SEC_START + 2).value = img4;
                    if (img5) nuevaFilaWmt.getCell(WALMART_COL.IMG_SEC_START + 3).value = img5;

                    procesados++;
                }
            }
        });

        // ====================================================================
        // 6. GUARDAR ARCHIVOS GENERADOS
        // ====================================================================
        const timestamp = Date.now();
        const nombreArchivoCoppel = `Catalogo_Coppel_Listo_${timestamp}.xlsx`;
        const nombreArchivoWalmart = `Catalogo_Walmart_Listo_${timestamp}.xlsx`;

        const rutaSalidaCoppel = path.join(__dirname, '../../uploads/', nombreArchivoCoppel);
        const rutaSalidaWalmart = path.join(__dirname, '../../uploads/', nombreArchivoWalmart);

        await libroCoppel.xlsx.writeFile(rutaSalidaCoppel);
        await libroWalmart.xlsx.writeFile(rutaSalidaWalmart);

        // Eliminamos el maestro temporal subido por el usuario para limpiar el servidor
        fs.unlinkSync(rutaMaestro);

        // Enviamos respuesta exitosa al frontend con las rutas de descarga
        return res.json({
            mensaje: 'Catálogos generados con éxito',
            procesados: procesados,
            archivos: {
                coppel: nombreArchivoCoppel,
                walmart: nombreArchivoWalmart
            }
        });

    } catch (error) {
        console.error('Error al generar catálogos:', error);
        return res.status(500).json({ error: 'Ocurrió un error interno al procesar los catálogos.' });
    }
}

module.exports = {
    generarCatalogos
};