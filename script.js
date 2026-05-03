// Función para exportar reporte de ausentes
function exportarReporteAusencia() {
    const asistencias = obtenerAsistencias();
    const presentes = new Set(asistencias.map(r => r.carnet));
    const ausentes = Object.entries(datosEstudiantes)
        .filter(([carnet]) => !presentes.has(carnet))
        .map(([carnet, datos]) => ({ carnet, nombre: datos.nombre, grado: datos.grado }));
    if (ausentes.length === 0) {
        mostrarMensaje('No hay ausentes para reportar.', 'info');
        return;
    }
    
    const fecha = new Date().toISOString().slice(0, 10);
    const win = window.open('', 'ReporteAusenciaPDF', 'width=800,height=1000');
    
    // Generar tabla HTML para Excel
    let tablaHTML = '<table border="1"><thead><tr><th>Carnet</th><th>Nombre</th><th>Grado</th></tr></thead><tbody>';
    ausentes.forEach(est => {
        tablaHTML += `<tr><td>${est.carnet}</td><td>${est.nombre}</td><td>${est.grado}</td></tr>`;
    });
    tablaHTML += '</tbody></table>';
    
    // Convertir a formato para Excel (HTML con formato de tabla)
    const excelContent = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Reporte de Ausencia</title></head><body>' + tablaHTML + '</body></html>';
    
    // Codificar para data URL
    const excelDataUrl = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelContent);
    
    win.document.write('<html><head><title>Reporte de Ausencia</title>');
    win.document.write('<style>body{font-family:sans-serif;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid #ccc;padding:8px;text-align:left;} th{background:#eee;} h2{margin-top:20px;} .accionesPDF{margin:20px 0;} .btnPDF{padding:8px 16px;border:none;border-radius:4px;background:#007bff;color:#fff;margin-right:8px;cursor:pointer;} .btnPDF:hover{background:#0056b3;}</style>');
    win.document.write('</head><body>');
    win.document.write('<div class="accionesPDF">');
    win.document.write('<button class="btnPDF" onclick="window.print()">Imprimir</button>');
    win.document.write('<button class="btnPDF" onclick="descargarPDF()">Descargar PDF</button>');
    win.document.write('<button class="btnPDF" onclick="descargarExcel()">Descargar Excel</button>');
    win.document.write('</div>');
    win.document.write('<h2>Reporte de Ausencia</h2>');
    win.document.write(`<p>Fecha: ${fecha}</p>`);
    win.document.write('<table><thead><tr><th>Carnet</th><th>Nombre</th><th>Grado</th></tr></thead><tbody>');
    ausentes.forEach(est => {
        win.document.write(`<tr><td>${est.carnet}</td><td>${est.nombre}</td><td>${est.grado}</td></tr>`);
    });
    win.document.write('</tbody></table>');
    win.document.write('<script>');
    win.document.write('function descargarPDF(){window.print();} ');
    win.document.write('function descargarExcel(){ ');
    win.document.write('const a = document.createElement("a"); ');
    win.document.write('a.href = "' + excelDataUrl + '"; ');
    win.document.write('a.download = "reporte_ausencia.xls"; ');
    win.document.write('a.click(); ');
    win.document.write('}');
    win.document.write('</' + 'script>');
    win.document.write('</body></html>');
    win.document.close();
    mostrarMensaje('Reporte de ausentes generado en PDF.', 'success');
}

// Función para exportar reporte de asistencia (PDF y Excel)
function exportarReporteAsistencia() {
    const asistencias = obtenerAsistencias();
    if (asistencias.length === 0) {
        mostrarMensaje('No hay registros para reportar.', 'error');
        return;
    }

    const fecha = new Date().toISOString().slice(0, 10);
    const win = window.open('', 'ReporteAsistenciaPDF', 'width=800,height=1000');
    
    // Generar tabla HTML para Excel
    let tablaHTML = '<table border="1"><thead><tr><th>Carnet</th><th>Nombre</th><th>Grado</th><th>Fecha</th><th>Hora</th></tr></thead><tbody>';
    asistencias.forEach(reg => {
        tablaHTML += `<tr><td>${reg.carnet}</td><td>${reg.nombre||'-'}</td><td>${reg.grado||'-'}</td><td>${reg.fecha}</td><td>${reg.hora}</td></tr>`;
    });
    tablaHTML += '</tbody></table>';
    
    // Convertir a formato para Excel (HTML con formato de tabla)
    const excelContent = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Reporte de Asistencia</title></head><body>' + tablaHTML + '</body></html>';
    
    win.document.write('<html><head><title>Reporte de Asistencia</title>');
    win.document.write('<style>body{font-family:sans-serif;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid #ccc;padding:8px;text-align:left;} th{background:#eee;} h2{margin-top:20px;} .accionesPDF{margin:20px 0;} .btnPDF{padding:8px 16px;border:none;border-radius:4px;background:#007bff;color:#fff;margin-right:8px;cursor:pointer;} .btnPDF:hover{background:#0056b3;}</style>');
    win.document.write('</head><body>');
    win.document.write('<div class="accionesPDF">');
    win.document.write('<button class="btnPDF" onclick="window.print()">Imprimir</button>');
    win.document.write('<button class="btnPDF" onclick="descargarPDF()">Descargar PDF</button>');
    win.document.write('<button class="btnPDF" onclick="descargarExcel()">Descargar Excel</button>');
    win.document.write('</div>');
    win.document.write('<h2>Reporte de Asistencia</h2>');
    win.document.write(`<p>Fecha: ${fecha}</p>`);
    win.document.write('<table><thead><tr><th>Carnet</th><th>Nombre</th><th>Grado</th><th>Fecha</th><th>Hora</th></tr></thead><tbody>');
    asistencias.forEach(reg => {
        win.document.write(`<tr><td>${reg.carnet}</td><td>${reg.nombre || '-'}</td><td>${reg.grado || '-'}</td><td>${reg.fecha}</td><td>${reg.hora}</td></tr>`);
    });
    win.document.write('</tbody></table>');
    win.document.write('<script>');
    win.document.write('function descargarPDF(){window.print();} ');
    win.document.write('function descargarExcel(){ ');
    // Usar data URL para evitar problemas de escape
    const excelDataUrl = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelContent);
    win.document.write('const a = document.createElement("a"); ');
    win.document.write('a.href = "' + excelDataUrl + '"; ');
    win.document.write('a.download = "reporte_asistencia.xls"; ');
    win.document.write('a.click(); ');
    win.document.write('}');
    win.document.write('</' + 'script>');
    win.document.write('</body></html>');
    win.document.close();
    mostrarMensaje('Reporte de asistencia generado en PDF.', 'success');
}
const CLAVE_STORAGE = 'asistencias';
const CARPETA_FOTOS = 'fotos';
const RUTAS_FOTOS = ['', 'FOTOS DE CUARTOS 2026', 'FOTOS DE QUINTOS', 'FOTOS DE SEXTO'];
const EXTENSIONES = ['jpg', 'jpeg', 'png', 'JPG', 'JPEG'];

let datosEstudiantes = {};

async function cargarEstudiantes() {
    try {
        let resJson = await fetch('datos/estudiantes_actualizado.json');
        if (!resJson.ok) {
            resJson = await fetch('datos/estudiantes.json');
        }
        if (resJson.ok) {
            const parsed = await resJson.json();
            if (parsed && Object.keys(parsed).length > 0) { datosEstudiantes = parsed; return; }
        }
    } catch (error) {
        console.error('Error cargando datos de estudiantes:', error);
    }
    datosEstudiantes = {};
}

function obtenerDatosEstudiante(carnet) {
    const d = datosEstudiantes[carnet];
    return d ? {
        nombre: d.nombre || '-',
        grado: d.grado || '-',
        imagen: d.imagen || null
    } : { nombre: '-', grado: '-', imagen: null };
}

function obtenerAsistencias() {
    try {
        const data = localStorage.getItem(CLAVE_STORAGE);
        return data ? JSON.parse(data) : [];
    } catch {
        return [];
    }
}

function guardarAsistencias(asistencias) {
    localStorage.setItem(CLAVE_STORAGE, JSON.stringify(asistencias));
}

function registrarAsistencia(carnet) {
    const ahora = new Date();
    const { nombre, grado } = obtenerDatosEstudiante(carnet);
    const registro = {
        carnet: String(carnet).trim(),
        nombre,
        grado,
        fecha: ahora.toISOString().slice(0, 10),
        hora: ahora.toTimeString().slice(0, 8)
    };
    const lista = obtenerAsistencias();
    lista.unshift(registro);
    guardarAsistencias(lista);
    return registro;
}

function buscarFotoEstudiante(carnet, callback) {
    const img = document.createElement('img');
    let rutaIndex = 0;
    let extIndex = 0;

    function siguienteIntento() {
        if (rutaIndex >= RUTAS_FOTOS.length) {
            callback(null);
            return;
        }

        const ruta = RUTAS_FOTOS[rutaIndex];
        const subcarpeta = ruta ? `${CARPETA_FOTOS}/${ruta}` : CARPETA_FOTOS;
        const src = `${subcarpeta}/${carnet}.${EXTENSIONES[extIndex]}`;

        img.onload = () => callback(src);
        img.onerror = () => {
            extIndex += 1;
            if (extIndex >= EXTENSIONES.length) {
                extIndex = 0;
                rutaIndex += 1;
            }
            siguienteIntento();
        };
        img.src = src;
    }

    siguienteIntento();
}

function mostrarFotoEstudiante(carnet) {
    const img = document.getElementById('fotoImg');
    const placeholder = document.getElementById('fotoPlaceholder');
    const nombreEl = document.getElementById('fotoNombre');
    const gradoEl = document.getElementById('fotoGrado');
    const carnetEl = document.getElementById('fotoCarnet');
    if (!img || !placeholder) return;

    const { nombre, grado, imagen } = obtenerDatosEstudiante(carnet);
    if (nombreEl) nombreEl.textContent = nombre;
    if (gradoEl) gradoEl.textContent = grado;
    if (carnetEl) carnetEl.textContent = `Carnet: ${carnet}`;

    if (imagen) {
        const src = `${CARPETA_FOTOS}/${imagen}`;
        img.onload = () => {
            img.classList.remove('hidden');
            placeholder.classList.add('hidden');
        };
        img.onerror = () => {
            img.classList.add('hidden');
            placeholder.classList.remove('hidden');
        };
        img.src = src;
    } else {
        buscarFotoEstudiante(carnet, (src) => {
            if (src) {
                img.src = src;
                img.classList.remove('hidden');
                placeholder.classList.add('hidden');
            } else {
                img.classList.add('hidden');
                placeholder.classList.remove('hidden');
            }
        });
    }
}

function mostrarMensaje(texto, tipo) {
    const el = document.getElementById('mensajeFlash');
    if (!el) return;
    el.textContent = texto;
    el.className = `mensaje-flash ${tipo}`;
    el.classList.remove('hidden');
    setTimeout(() => el.classList.add('hidden'), 2500);
}

function renderizarTabla() {
    const tbody = document.getElementById('tablaBody');
    const sinRegistros = document.getElementById('sinRegistros');
    if (!tbody || !sinRegistros) return;

    const asistencias = obtenerAsistencias();

    if (asistencias.length === 0) {
        tbody.innerHTML = '';
        sinRegistros.classList.remove('hidden');
        return;
    }

    sinRegistros.classList.add('hidden');
    tbody.innerHTML = asistencias
        .slice(0, 100)
        .map(r => {
            const nombre = r.nombre != null ? escapeHtml(r.nombre) : '-';
            const grado = r.grado != null ? escapeHtml(r.grado) : '-';
            return `<tr><td>${escapeHtml(r.carnet)}</td><td>${nombre}</td><td>${grado}</td><td>${r.fecha}</td><td>${r.hora}</td></tr>`;
        })
        .join('');
}

function escapeHtml(texto) {
    const div = document.createElement('div');
    div.textContent = texto;
    return div.innerHTML;
}

function exportarCSV() {
    const asistencias = obtenerAsistencias();
    if (asistencias.length === 0) {
        mostrarMensaje('No hay registros para exportar.', 'error');
        return;
    }

    const filas = [
        ['Carnet', 'Nombre', 'Grado', 'Fecha', 'Hora'],
        ...asistencias.map(r => [
            r.carnet,
            r.nombre != null ? r.nombre : '-',
            r.grado != null ? r.grado : '-',
            r.fecha,
            r.hora
        ])
    ];

    const csv = filas
        .map(row => row.map(celda => {
            const s = String(celda);
            if (s.includes(',') || s.includes('"') || s.includes('\n')) {
                return '"' + s.replace(/"/g, '""') + '"';
            }
            return s;
        }).join(','))
        .join('\r\n');

    const bom = '\uFEFF';
    const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `asistencia_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    mostrarMensaje('Archivo exportado correctamente.', 'success');
}

function limpiarTodo() {
    if (!confirm('¿Eliminar el historial de asistencia?')) return;
    guardarAsistencias([]);
    renderizarTabla();
    mostrarMensaje('Historial borrado. Los datos de estudiantes permanecen intactos.', 'success');
}

const form = document.getElementById('scanForm');
const input = document.getElementById('carnetInput');

form?.addEventListener('submit', (e) => {
    e.preventDefault();
    const carnet = input?.value?.trim();
    if (!carnet) return;

    registrarAsistencia(carnet);
    mostrarFotoEstudiante(carnet);
    mostrarMensaje(`Asistencia registrada: ${carnet}`, 'success');
    input.value = '';
    input.focus();
    renderizarTabla();
});

// Lógica de importación eliminada
document.getElementById('btnReporteAusencia')?.addEventListener('click', exportarReporteAusencia);
document.getElementById('btnReporteAsistencia')?.addEventListener('click', exportarReporteAsistencia);

// Buscador de estudiante
document.getElementById('btnBuscar')?.addEventListener('click', () => {
    const valor = document.getElementById('buscadorInput')?.value?.trim().toLowerCase();
    if (!valor) return;
    const encontrado = Object.entries(datosEstudiantes).find(([carnet, datos]) =>
        carnet.toLowerCase() === valor ||
        (datos.nombre && datos.nombre.toLowerCase().includes(valor))
    );
    if (encontrado) {
        const [carnet] = encontrado;
        mostrarFotoEstudiante(carnet);
        document.getElementById('carnetInput').value = carnet;
        mostrarMensaje('Estudiante encontrado.', 'success');
    } else {
        mostrarMensaje('No se encontró el estudiante.', 'error');
    }
});
document.getElementById('btnLimpiar')?.addEventListener('click', limpiarTodo);

// Solo enfocar carnetInput si el click no es en el buscador
document.addEventListener('click', (e) => {
    const target = e.target;
    if (target && (target.id === 'buscadorInput' || target.id === 'btnBuscar')) return;
    input?.focus();
});

(async () => {
    await cargarEstudiantes();
    renderizarTabla();
    input?.focus();
})();
