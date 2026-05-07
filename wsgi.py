"""
Punto de entrada WSGI para servidores de producción (gunicorn).
No modificar el código de la aplicación — solo expone el objeto Flask.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.web.server import app  # noqa: E402

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
