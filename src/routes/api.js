const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const { procesarActualizacionStock } = require('../controllers/stockController');
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, path.join(__dirname, '../../uploads'))
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + '-' + file.originalName);
    }
});

const upload = multer({ storage: storage });

router.post('/subir-stock', upload.single('archivoStock'), procesarActualizacionStock);

router.get ('/descargar/:nombreArchivo', (req, res) => {
    const archivo = req.params.nombreArchivo;
    const rutaArchivo = path.join(__dirname, '../../uploads/', archivo);

    if (!fs.existsSync(rutaArchivo)) {
        return res.status(404).json({ mensaje: 'Archivo no encontrado' });
    }

    res.download(rutaArchivo, archivo, (err) => {
        if (err) {
            console.error('error al descargar)', err)
    } else {
        try {
            fs.unlinkSync(rutaArchivo);
            console.log("Archivo descargado y eliminando el storage");
        } catch (errorBorrado) {
            console.error('Error al eliminar el archivo después de la descarga:', errorBorrado);
        }
    };
})
});
    

module.exports = router;