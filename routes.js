const express = require('express');
const router = express.Router();
const sheet = require('./spreadsheet');


// Rutas

router.get('/buscar_todos', async (req, res) => {
    let registros = {};

    registros.filas = await sheet.getRecords();
    registros.total = (registros.filas).length;

    res.send(registros);
});

router.get('/buscar_uno/:codigo', async (req, res) => {
    //console.log(req.params.codigo);

    let registro = await sheet.getRecord(req.params.codigo);

    res.send(registro);
});


router.post('/crear_nuevo', (req, res) => {
    let datos = req.body;

    let mensaje = sheet.createFile(datos);

    res.send(mensaje);
});

router.post('/cargar_uno', async (req, res) => {
    let datos = req.body;

    let mensaje = await sheet.addRow(datos);

    res.send(mensaje);
});

router.post('/cargar_varios', async (req, res) => {
    let datos = req.body;

    let mensaje = await sheet.addRows(datos);

    res.send(mensaje);
});

router.post('/actualizar_datos', async (req, res) => {
    let datos = req.body;

    let mensaje = await sheet.update(datos);

    res.send(mensaje);
});

module.exports = router;
