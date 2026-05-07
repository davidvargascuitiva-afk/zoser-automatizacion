import threading
import webbrowser
import time
from app.web.server import app


def main():
    def abrir_navegador():
        time.sleep(1.2)
        webbrowser.open('http://127.0.0.1:5000')

    threading.Thread(target=abrir_navegador, daemon=True).start()
    print("Automatización — Zoser SAS corriendo en http://127.0.0.1:5000")
    print("Presiona Ctrl+C para cerrar la aplicación.")
    app.run(debug=False, port=5000, threaded=True)


if __name__ == "__main__":
    main()
