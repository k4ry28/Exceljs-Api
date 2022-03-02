const ExcelJS = require('exceljs');


function createFile(datos) {
    try {
        const workbook = new ExcelJS.Workbook();

        const worksheet = workbook.addWorksheet('My Sheet');

        worksheet.columns = [
            { header: 'alta', key: 'alta', width: 15 },
            { header: 'codigo_cliente', key: 'codigo_cliente', width: 15 },
            { header: 'nombre', key: 'nombre', width: 50 },
            { header: 'ip', key: 'ip', width: 20 },
            { header: 'etiquetas', key: 'etiquetas', width: 20 },
            { header: 'numTk', key: 'numTk', width: 15 },
            { header: 'comentario', key: 'comentario', width: 60 },
        ];

        for (let i = 0; i < datos.length; i++) {

            worksheet.addRow(datos[i]);

        }

        workbook.xlsx.writeFile('./prueba-excel.xlsx')
            .then(() => {
                console.log('Archivo creado')
            })
            .catch(err => {
                console.log(err.message);
            });

        return 'Datos guardados';

    } catch (error) {
        console.log(error);
        return 'Error al guardar';
    }
}

async function getRecords() {
    try {
        const workbook = new ExcelJS.Workbook();

        let registros = await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet')

                let n = 0, fila;
                let filas = [];

                for (let i = 2; i <= worksheet.actualRowCount; i++) {

                    fila = worksheet.getRow(i).values;
                    fila.shift();

                    filas[n] = {
                        alta: fila[0],
                        codigo_cliente: fila[1],
                        nombre: fila[2],
                        ip: fila[3],
                        etiquetas: fila[4],
                        numTk: fila[5],
                        comentario: fila[6]
                    }

                    n++;
                    //console.log('Dato ' + n);
                    //console.log(fila);
                }

                return filas;
            });

        return registros;

    } catch (error) {
        console.log(error);
        return 'Error al leer archivo';
    }
}

async function getRecord(dato) {

    try {

        const workbook = new ExcelJS.Workbook();

        let registro = await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet');

                // Trae los numeros de clientes mas dos valores al principio: el titulo de columna y un elemento vacío q debe ser el numero de columna
                cod_cliente = worksheet.getColumn(2).values;

                let fila;
                let registro;

                //console.log(cod_cliente);

                for (let i = 1; i < cod_cliente.length; i++) {

                    if (cod_cliente[i] == dato) {
                        fila = worksheet.getRow(i).values;
                        // Elimino el elemento vacío q hace referencia al numero de fila:
                        fila.shift();

                        registro = {
                            alta: fila[0],
                            codigo_cliente: fila[1],
                            nombre: fila[2],
                            ip: fila[3],
                            etiquetas: fila[4],
                            numTk: fila[5],
                            comentario: fila[6]
                        }

                        console.log('Registro encontrado: ');
                        console.log(fila);
                    }
                }

                if (registro == undefined) {
                    let mensaje = 'No se encontro ningun registro que coincida con la busqueda'
                    return mensaje;
                }

                return registro;
            }, dato)

        return registro;

    } catch (error) {
        console.log(error);
        return 'Error en la busqueda';
    }
}

async function addRow(dato) {

    try {

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet');

                let fila = [
                    dato[i].alta,
                    dato[i].codigo_cliente,
                    dato[i].nombre,
                    dato[i].ip,
                    dato[i].etiquetas,
                    dato[i].numTk,
                    dato[i].comentario
                ]
                worksheet.addRow(fila)

            })


        workbook.xlsx.writeFile('./prueba-excel.xlsx')
            .then(() => {
                console.log('Datos guardados')
            })
            .catch(err => {
                console.log(err.message);
            });

        return 'Dato guardado';

    } catch (error) {
        console.log(error);
        return 'Error al guardar';
    }
}

async function addRows(datos) {

    try {

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet');
                let fila = [];

                for (let i = 0; i < datos.length; i++) {
                    fila = [
                        datos[i].alta,
                        datos[i].codigo_cliente,
                        datos[i].nombre,
                        datos[i].ip,
                        datos[i].etiquetas,
                        datos[i].numTk,
                        datos[i].comentario
                    ];
                    worksheet.addRow(fila);
                }

            });

        workbook.xlsx.writeFile('./prueba-excel.xlsx')
            .then(() => {
                console.log('Datos guardados')
            })
            .catch(err => {
                console.log(err.message);
            });

        return 'Datos guardados';

    } catch (error) {
        console.log(error);
        return 'Error al guardar';
    }
}

async function update(datos) {
    try {

        const workbook = new ExcelJS.Workbook();
        let registros = {};

        // Agregar registros nuevos
        registros.creados = await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet');

                // Trae los numeros de clientes mas dos valores al principio: el titulo de columna y un elemento vacío q debe ser el numero de columna
                cod_cliente = worksheet.getColumn(2).values;
                // Elimino esos primeros dos valores q no me sirven:
                cod_cliente.shift();
                cod_cliente.shift();

                //console.log(cod_cliente);

                let n = 0;
                let fila = [];

                for (let i = 0; i < datos.length; i++) {
                    // Si el excel (viejo) no tiene un valor del json de Sin Internet (nuevo) se agrega la fila:                    
                    if (cod_cliente.includes(datos[i].codigo_cliente) == false) {
                        fila = [
                            datos[i].alta,
                            datos[i].codigo_cliente,
                            datos[i].nombre,
                            datos[i].ip,
                            datos[i].etiquetas,
                            datos[i].numTk,
                            datos[i].comentario
                        ];
                        worksheet.addRow(fila);
                        n++;
                    }
                }
                return n;

            }, datos);

        console.log('Registros creados: ' + registros.creados);


        // Guardar registros nuevos
        if (registros.creados > 0) {

            await workbook.xlsx.writeFile('./prueba-excel.xlsx')
                .then(() => {
                    console.log('Datos guardados')
                })
                .catch(err => {
                    console.log(err.message);
                });
        }

        // Buscar registros viejos resueltos: 
        registros.borrados = await workbook.xlsx.readFile('./prueba-excel.xlsx')
            .then(() => {
                let worksheet = workbook.getWorksheet('My Sheet');

                cod_cliente = worksheet.getColumn(2).values;
                cod_cliente.shift();
                cod_cliente.shift();

                //console.log(cod_cliente);

                let n = 0;

                // Borrar del excel filas q ya no esten en json de Sin Internet (casos resueltos)
                let presencia = false;

                for (let i = 0; i < cod_cliente.length; i++) {

                    for (let j = 0; j < datos.length; j++) {

                        if (cod_cliente[i] == datos[j].codigo_cliente) {
                            presencia = true;
                            //j = datos.length;
                            //console.log(`cliente: ${cod_cliente[i]} esta`);
                        }

                        if ((j == (datos.length - 1)) && (presencia == false)) {
                            console.log(`cliente: ${cod_cliente[i]} no esta. BORRADO`);
                            worksheet.spliceRows((i + 2), 1);
                            n++;
                        }
                    }

                    presencia = false;
                }

                return n;

            }, datos);

        console.log('Registros borrados: ' + registros.borrados);


        // Borrar registros viejos
        if (registros.borrados > 0) {

            await workbook.xlsx.writeFile('./prueba-excel.xlsx')
                .then(() => {
                    console.log('Datos guardados')
                })
                .catch(err => {
                    console.log(err.message);
                });
        }


        let mensaje = 'Archivo actualizado.\nRegistros agregados: ' + registros.creados + '\nRegistros eliminados: ' + registros.borrados;

        return mensaje;

    } catch (error) {
        console.log(error);
        return 'Error al actualizar';
    }
}


module.exports = {
    createFile: createFile,
    getRecords: getRecords,
    getRecord: getRecord,
    addRow: addRow,
    addRows: addRows,
    update: update,
}


