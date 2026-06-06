const express = require('express');
const router = express.Router();
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const upload = multer({ dest: 'uploads/' });

router.post('/plantilla', upload.single('plantilla'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No se subió ningún archivo' });
        if (!req.body.tipo) return res.status(400).json({ error: 'Falta el tipo de plantilla' });

        const rutaTemporal = req.file.path;
        const tipoPlantilla = req.body.tipo; // Ej: 'coppel_stock'
        
        // Armamos el nombre oficial esperado por tus controladores
        const nombreOficial = `plantilla_${tipoPlantilla}.xlsx`;
        
        // Definimos la ruta de la carpeta templates/
        const rutaDestino = path.join(__dirname, '../../templates/', nombreOficial);

        // Asegurarnos de que la carpeta templates exista
        const carpetaTemplates = path.dirname(rutaDestino);
        if (!fs.existsSync(carpetaTemplates)) {
            fs.mkdirSync(carpetaTemplates, { recursive: true });
        }

        // Sobrescribimos el archivo (Mueve el temporal a la carpeta templates)
        fs.renameSync(rutaTemporal, rutaDestino);

        return res.json({ mensaje: 'Plantilla actualizada exitosamente' });

    } catch (error) {
        console.error('Error al actualizar plantilla:', error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        return res.status(500).json({ error: 'Error interno al guardar la plantilla' });
    }
});

module.exports = router;