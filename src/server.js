const express = require('express');
const cors = require('cors');
const path = require('path');
const catalogRoutes = require('./routes/catalogRoutes');

const app = express();
const PORT = process.env.PORT || 3000;

//MIDDLEWARES
app.use(cors());
app.use(express.json())

//conexion con la interfaz web
app.use(express.static(path.join(__dirname, '../')));

const apiRoutes = require('./routes/api');
app.use ('/api', apiRoutes);

app.use('/api/catalogo', catalogRoutes);

app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

app.listen(PORT, () => {
    console.log(`servidor corriendo:${PORT}`);
});




