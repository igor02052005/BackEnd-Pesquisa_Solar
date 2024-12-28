from flask import Flask, render_template, request, jsonify
import subprocess

app = Flask(__name__)

# Caminho para o script
SCRIPT_PATH = r"C:\Users\igort\OneDrive\Desktop\VSCODE\PESQUISA SOLAR\CONCAT-GERAÇÃO CLOUD BETA OFC.py"

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/executar", methods=["POST"])
def executar_script():
    try:
        # Executa o script Python
        result = subprocess.run(["python", SCRIPT_PATH], capture_output=True, text=True)
        output = result.stdout  # Captura a saída do script
        return jsonify({"status": "success", "output": output})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
