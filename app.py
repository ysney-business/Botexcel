import os
import io
import pandas as pd
from flask import Flask, jsonify, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

app = Flask(__name__)

# Configuración de seguridad para limitar el tamaño de archivos subidos (Ej: máximo 16MB)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

@app.route("/", methods=["GET"])
def health_check():
    """Ruta de verificación de estado y salud del sistema."""
    return jsonify({
        "status": "success",
        "message": "Servidor activo, seguro y operando correctamente en produccion.",
        "environment": os.environ.get("RENDER", "local")
    }), 200

@app.route("/api/procesar-excel", methods=["POST"])
def procesar_excel():
    """
    Endpoint blindado para la recepcion, analisis y procesamiento 
    de archivos de Excel aplicables a cualquier modelo de negocio.
    """
    try:
        # Validacion de existencia de archivos en la peticion HTTP
        if 'file' not in request.files:
            return jsonify({
                "error": "Peticion incompleta",
                "mensaje": "No se encontro ningun archivo adjunto con la clave 'file'."
            }), 400

        file = request.files['file']
        
        if file.filename == '':
            return jsonify({
                "error": "Archivo no seleccionado",
                "mensaje": "El nombre del archivo esta vacio."
            }), 400

        # Validacion de extension de archivo permitida
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({
                "error": "Formato no soportado",
                "mensaje": "El archivo debe ser una planilla de Excel valida (.xlsx o .xls)."
            }), 400

        # Lectura segura del archivo mediante Pandas
        df = pd.read_excel(file)

        # Lógica de negocio automatizada para inventarios, cierres o cálculos
        if 'Cantidad' in df.columns and 'Precio' in df.columns:
            df['Total'] = pd.to_numeric(df['Cantidad'], errors='coerce') * pd.to_numeric(df['Precio'], errors='coerce')

        # Almacenamiento temporal seguro en memoria RAM
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Procesado_Optimo')
        output.seek(0)

        # Retorno seguro del archivo procesado para descarga directa
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='resultado_procesado_optimo.xlsx'
        )

    except RequestEntityTooLarge:
        return jsonify({
            "error": "Archivo demasiado grande",
            "mensaje": "El archivo excede el limite maximo permitido de 16MB."
        }), 413

    except Exception as e:
        return jsonify({
            "error": "Falla critica en el procesamiento del flujo de datos",
            "detalle": str(e)
        }), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
