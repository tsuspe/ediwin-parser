📦 EDIWIN Parser — Eurofiel & El Corte Inglés
Automatización del procesamiento de pedidos desde PDFs EDIWIN

Creado por Aitor Susperregui (@elvasco.x)

🚀 ¿Qué hace este proyecto?

Esta herramienta convierte automáticamente los PDFs de pedidos descargados desde EDIWIN en un Excel limpio, coloreado y listo para trabajar.

Compatible con:

Eurofiel

El Corte Inglés

Incluye:

✔️ Lectura avanzada de PDFs (pdfplumber)
✔️ Identificación automática de pedidos
✔️ Extracción de modelos, color, cantidades, fechas, sucursales, precio…
✔️ Web App en Streamlit
✔️ Excel con:

colores por modelo

bordes finos

cabeceras amarillo corporativo

filas TOTAL automáticas
✔️ CSV export
✔️ Cero errores manuales, cero horas perdidas

🧬 Estructura del proyecto
/src
   app.py                 # Web Streamlit
   eurofiel_parser.py     # Parser Eurofiel
   eci_parser.py          # Parser Corte Inglés

/input                    # PDFs
/output                   # Informes Excel y CSV
/docs                     # Capturas y documentación

requirements.txt
README.md
.gitignore

🛠 Instalación

Clona el repositorio:

git clone https://github.com/tsuspe/ediwin-parser.git
cd ediwin-parser


Instala dependencias:

pip install -r requirements.txt


Ejecuta la app:

streamlit run src/app.py

📥 Cómo usarlo

Coloca tus PDFs en /input

En Streamlit selecciona:

“Eurofiel”

o “El Corte Inglés”

Sube el archivo

Descarga:

Excel con colores y totales

CSV

Resúmenes por modelo, modelo+color…

🧠 Tecnología usada

Python

Streamlit

Pandas

pdfplumber

openpyxl

RegEx avanzado

Arquitectura modular

❤️ Creado con mucho amor por:

Aitor Susperregui

@elvasco.x
