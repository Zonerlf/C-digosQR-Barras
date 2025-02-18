// Guardar y cargar el folio
function loadFolio() {
    const storedFolio = localStorage.getItem('folio');
    return storedFolio ? parseInt(storedFolio, 10) : 1;
}

function saveFolio(folio) {
    localStorage.setItem('folio', folio);
}

// Procesar archivo de entrada
document.querySelector('#file-input').addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const fileData = e.target.result;
        let data;
        const ext = file.name.split('.').pop().toLowerCase();

        if (ext === 'xlsx' || ext === 'xls') {
            data = XLSX.read(fileData, { type: 'binary' });
            const sheet = data.Sheets[data.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            populateTable(jsonData);
        }
        // Aquí puedes agregar más casos si soportas otros tipos de archivos
    };
    reader.readAsBinaryString(file);
}

// Población de la tabla con los datos
function populateTable(data) {
    const tableBody = document.querySelector('#data-table tbody');
    tableBody.innerHTML = '';  // Limpiar solo el contenido del tbody
    let folio = loadFolio();

    const rows = data.slice(1);  // Eliminar la primera fila (encabezado)

    rows.forEach((row) => {
        const qrData = {
            code: row[0],
            description: row[1],
            date: formatDate(row[2]),  // Aquí aplicamos la función formatDate
            invoice: row[3],
            oc: row[4],
            provider: 'comavime.com.mx',
            client: 'CONDUMEX S.A. de C.V.',
            folio: folio++
        };

        const qrCodeCanvas = document.createElement('canvas');
        QRCode.toCanvas(qrCodeCanvas, JSON.stringify(qrData), (error) => {
            if (error) console.error(error);
        });

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatDate(row[2])}</td>
            <td>${row[3]}</td>
            <td>${row[4]}</td>
            <td>MF500/56</td>
            <td>CONDUMEX S.A. de C.V.</td>
            <td>www.comavime.com.mx</td>
        `;

        const qrTd = document.createElement('td');
        qrTd.appendChild(qrCodeCanvas);
        tr.appendChild(qrTd);

        tableBody.appendChild(tr);
    });

    saveFolio(folio);
}

// Función para formatear la fecha en formato dd/mm/yyyy
function formatDate(date) {
    // Verificar si la fecha es un número, lo que indica que proviene de Excel
    if (typeof date === 'number') {
        const excelDate = new Date((date - 25569) * 86400 * 1000); // Convertir número de Excel a fecha
        return formatToDDMMYYYY(excelDate);
    }

    const parsedDate = new Date(date);
    if (isNaN(parsedDate.getTime())) return date;  // Si no es una fecha válida, retornamos tal cual

    // Extraemos el día, mes y año
    return formatToDDMMYYYY(parsedDate);
}

function formatToDDMMYYYY(parsedDate) {
    const day = parsedDate.getDate().toString().padStart(2, '0');
    const month = (parsedDate.getMonth() + 1).toString().padStart(2, '0');
    const year = parsedDate.getFullYear();
    return `${day}/${month}/${year}`;
}

// Generar todos los QR cuando se hace clic en el botón
document.querySelector('#generate-all-qr-btn').addEventListener('click', generateAllQRs);

function generateAllQRs() {
    const rows = document.querySelectorAll('#data-table tbody tr');
    rows.forEach((row, index) => {
        const qrData = {
            code: row.children[0].innerText,
            description: row.children[1].innerText,
            date: row.children[2].innerText,
            invoice: row.children[3].innerText,
            oc: row.children[4].innerText,
            regiter: 'MF500/56',
            provider: 'www.comavime.com.mx',
            client: 'CONDUMEX S.A. de C.V.',
            folio: loadFolio() + index + 1
        };

        const qrCodeCanvas = row.querySelector('canvas');
        if (qrCodeCanvas) {
            QRCode.toCanvas(qrCodeCanvas, JSON.stringify(qrData), (error) => {
                if (error) console.error(error);
            });
        }
    });
}

// Exportar tabla a archivo XLS
document.querySelector('#export-btn').addEventListener('click', exportToXLS);

function exportToXLS() {
    const wb = XLSX.utils.table_to_book(document.querySelector('#data-table'), { sheet: 'Sheet1' });
    XLSX.writeFile(wb, 'table_dataQR.xlsx');
}

// Imprimir códigos QR
document.querySelector('#print-qr-btn').addEventListener('click', printQRs);

function printQRs() {
    const rows = document.querySelectorAll('#data-table tbody tr');
    const qrData = []; // Almacenará las imágenes base64 de los QR y sus identificadores

    // Recopilar los canvas generados para los QR y sus identificadores
    rows.forEach((row) => {
        const qrCodeCanvas = row.querySelector('canvas');
        const code = row.children[0].innerText; // Obtener el valor de la primera columna (código)
        const date = row.children[2].innerText; // Obtener la fecha
        const url = row.children[6].innerText; // Obtener la URL
        if (qrCodeCanvas) {
            // Convertir el canvas a una imagen base64
            const qrImage = qrCodeCanvas.toDataURL('image/png');
            qrData.push({ image: qrImage, code: code, date: date, url: url }); // Guardar la imagen y los datos
        }
    });

    // Verificamos si hay QR para imprimir
    if (qrData.length > 0) {
        // Creamos una ventana para la impresión
        const printWindow = window.open('', '', 'width=800,height=600');
        printWindow.document.write('<html><head><title>Impresión de Códigos QR</title>');
        printWindow.document.write(`
            <style>
                @media print {
                    body { 
                        font-family: Arial, sans-serif; 
                        margin: 0; /* Eliminar márgenes predeterminados */
                    }
                    .page {
                        display: grid;
                        grid-template-columns: repeat(4, 1fr); /* 4 columnas */
                        grid-template-rows: repeat(6, 1fr); /* 10 filas */
                        gap: 2px; /* Espacio mínimo entre los QR */
                        width: 98%;
                        height: 100vh; /* Una página por cada cuadrícula */
                        page-break-after: always; /* Forzar salto de página después de cada cuadrícula */
                        padding: 0.5cm; /* Márgenes reducidos */
                        box-sizing: border-box;
                    }
                    .qr-container {
                        display: flex;
                        flex-direction: column;
                        justify-content: center;
                        align-items: center;
                        border: 1px solid #ccc; /* Borde opcional para los QR */
                        padding: 1px; /* Espacio interno reducido */
                    }
                    img {
                        width: 100%; /* El QR ocupa todo el ancho del contenedor */
                        height: auto; /* Altura automática para mantener la proporción */
                    }
                    .qr-code {
                        font-size: 5px; /* Tamaño pequeño para el identificador */
                        text-align: center;
                        margin-top: 2px; /* Espacio mínimo entre el QR y el texto */
                    }
                }
            </style>
        `);
        printWindow.document.write('</head><body>');

        // Dividir los QR en páginas de 40 (4x10)
        const qrPerPage = 20; // 4 columnas × 10 filas
        for (let i = 0; i < qrData.length; i += qrPerPage) {
            const pageQrData = qrData.slice(i, i + qrPerPage); // QR para esta página

            // Crear una página con una cuadrícula de 4x10
            printWindow.document.write('<div class="page">');
            pageQrData.forEach((qr) => {
                printWindow.document.write(`
                    <div class="qr-container">
                        <img src="${qr.image}" alt="Código QR">
                        <div class="qr-code">${qr.code}<br>${qr.date}<br>${qr.url}</div> <!-- Identificador debajo del QR -->
                    </div>
                `);
            });
            printWindow.document.write('</div>'); // Cerrar la página
        }

        printWindow.document.write('</body></html>');
        printWindow.document.close();

        // Esperamos que el contenido se cargue antes de imprimir
        printWindow.onload = function () {
            printWindow.print();
            printWindow.close();
        };
    } else {
        alert('No se encontraron códigos QR para imprimir.');
    }
}