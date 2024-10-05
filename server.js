const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pdf = require('pdf-parse');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

// Configurar Multer para manejar la subida de archivos en memoria
const storage = multer.memoryStorage(); // Usar memoria en lugar de disco
const upload = multer({ storage });

// Ruta para mostrar el formulario de subida
app.get('/', (req, res) => {
    res.send(`
        <h1>Subir PDF para Convertir a Excel</h1>
        <form action="/convertir" method="POST" enctype="multipart/form-data">
            <input type="file" name="pdf" accept="application/pdf" required>
            <button type="submit">Convertir</button>
        </form>
    `);
});

// Ruta para manejar la subida y conversión
app.post('/convertir', upload.single('pdf'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No se ha subido ningún archivo.');
    }

    // Leer el archivo PDF desde el buffer en memoria
    let dataBuffer = req.file.buffer;

    pdf(dataBuffer).then(async (data) => {
        // Crear un nuevo libro de Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Datos');

        // Dividir el texto en líneas y agregarlo a la hoja de Excel
        const lines = data.text.split('\n');
        lines.forEach((line) => {
            worksheet.addRow([line]);
        });

        // Configurar la respuesta para descargar el archivo Excel
        res.setHeader('Content-Disposition', `attachment; filename="${req.file.originalname.replace('.pdf', '.xlsx')}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Escribir el archivo Excel en la respuesta
        await workbook.xlsx.write(res);
        res.end(); // Finalizar la respuesta
    }).catch(err => {
        console.error(err);
        res.status(500).send('Error al procesar el PDF.');
    });
});

app.listen(PORT, () => {
    console.log(`Servidor escuchando en http://localhost:${PORT}`);
});
