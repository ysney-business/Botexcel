// Módulo de control asíncrono para la interfaz de usuario (Evita congelamiento visual)
const { exec } = require('child_process');

function dispararProcesoBackend(tareaNombre) {
    console.log(`[UI] Solicitando ejecución de: ${tareaNombre}...`);
    
    exec(`python3 app.py --task="${tareaNombre}"`, (error, stdout, stderr) => {
        if (error) {
            console.error(`[ERROR CRÍTICO] El proceso falló: ${error.message}`);
            return;
        }
        if (stderr) {
            console.warn(`[ADVERTENCIA] ${stderr}`);
        }
        console.log(`[RESPUESTA DEL MOTOR]:\n${stdout}`);
    });
}

dispararProcesoBackend("Inventario");
