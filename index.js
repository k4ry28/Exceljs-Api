const express = require('express');
const router = require('./routes');

const port = 4000;

const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use('/excel-api', router);

app.listen(port, () => {
    console.log(`Servidor iniciado en puerto ${port}`);
})