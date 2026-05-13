const express = require('express');
const router = express.Router();
const multer = require('multer');
const catalogController = require('../controllers/catalogController');

// Usamos la carpeta uploads/ para guardar temporalmente el maestro crudo
const upload = multer({ dest: 'uploads/' });

// Definimos el endpoint POST que recibe el archivo con el campo 'maestro'
router.post('/generar', upload.single('maestro'), catalogController.generarCatalogos);

module.exports = router;