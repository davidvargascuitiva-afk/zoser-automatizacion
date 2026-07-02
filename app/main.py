import os
import sys
import socket
import threading
import webbrowser
import time
from app.web.server import app

_TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'web', 'templates')
_TEMPLATE_REQUERIDO = os.path.join(_TEMPLATES_DIR, 'index.html')


def _verificar_templates():
    """
    Garantiza que index.html existe antes de arrancar.
    Si fue renombrado accidentalmente, lo restaura solo y registra el aviso.
    Si no hay ningún HTML, aborta con mensaje claro.
    """
    if os.path.isfile(_TEMPLATE_REQUERIDO):
        return  # todo bien

    htmls = [f for f in os.listdir(_TEMPLATES_DIR) if f.lower().endswith('.html')] \
            if os.path.isdir(_TEMPLATES_DIR) else []

    if len(htmls) == 1:
        ruta_encontrada = os.path.join(_TEMPLATES_DIR, htmls[0])
        os.rename(ruta_encontrada, _TEMPLATE_REQUERIDO)
        print(f"[AVISO] Template renombrado automaticamente: '{htmls[0]}' -> 'index.html'")
        return

    # Sin recuperación posible — fallar con mensaje claro
    print("\n" + "=" * 60)
    print("ERROR CRÍTICO — La aplicación no puede iniciar.")
    print(f"  No se encontró el template 'index.html' en:")
    print(f"  {_TEMPLATES_DIR}")
    if htmls:
        print(f"  Archivos HTML presentes: {', '.join(htmls)}")
        print("  Renombra uno de ellos a 'index.html' y vuelve a intentar.")
    else:
        print("  La carpeta 'templates' está vacía o no existe.")
    print("=" * 60 + "\n")
    sys.exit(1)


def _obtener_ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


def main():
    _verificar_templates()

    ip_local = _obtener_ip_local()
    url_local = f'http://127.0.0.1:5000'
    url_red   = f'http://{ip_local}:5000'

    def abrir_navegador():
        time.sleep(1.2)
        webbrowser.open(url_local)

    threading.Thread(target=abrir_navegador, daemon=True).start()

    print("")
    print("=" * 55)
    print("  Automatizacion Zoser SAS — ACTIVA")
    print("=" * 55)
    print(f"  Este equipo  : {url_local}")
    print(f"  Red local    : {url_red}   <-- comparte esta")
    print("=" * 55)
    print("  Otros equipos de la red pueden entrar desde")
    print("  su navegador usando la direccion de Red local.")
    print("  Presiona Ctrl+C para cerrar la aplicacion.")
    print("")

    app.run(host='0.0.0.0', debug=False, port=5000, threaded=True)


if __name__ == "__main__":
    main()
