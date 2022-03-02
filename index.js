const express = require('express');
const router = require('./routes');

const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use('/excel-api', router);

app.listen(3000, () => {
    console.log('Servidor iniciado en puerto 3000');
})