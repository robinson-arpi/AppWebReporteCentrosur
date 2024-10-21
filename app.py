from flask import Flask, request, jsonify
import utils.data_preprocessing as d_p
from utils.json_creation import get_df_as_json
from utils.report_generation import process_data_for_report
import os

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Obtención del path del archivo desde el JSON enviado en la solicitud
        data = request.json
        file_path = data.get("path")

        # Separar nombre de archivo y ruta
        directory, filename = os.path.split(file_path)

        # Quitar la extensión del nombre del archivo
        filename = os.path.splitext(filename)[0]

        if not file_path:
            return jsonify({"msg": "No se proporcionó el path del archivo"}), 400

        # Lectura de df desde la ruta estática
        df = d_p.process_data(file_path)
        
        # Guardar el DataFrame procesado como JSON
        json_data = get_df_as_json(df)
        
        # Generación de reporte
        report_path = process_data_for_report(df, f"{filename}_mem.xlsx")

        return jsonify({"ok" : True, "msg": "","data":json_data,"report_path": f"{report_path}" }), 200

    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False)
