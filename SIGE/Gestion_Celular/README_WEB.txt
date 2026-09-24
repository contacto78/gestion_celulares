GESTION CORPORATIVA DE CELULARES - VERSION WEB

Esta version web convive con la aplicacion Tkinter original. Usa los mismos archivos:
- Base_Datos_Celulares.xlsx
- Acta de Entrega.docx
- carpeta Actas Generadas

INSTALACION
1. Abra una terminal en la carpeta Gestion_Celular.
2. Instale las dependencias:
   python -m pip install -r requirements_web.txt
3. Ejecute:
   python app_web.py
4. Abra en el navegador:
   http://127.0.0.1:5000

NOTAS
- Los usuarios y claves son los mismos de la hoja Usuarios del Excel.
- La generacion de actas PDF sigue requiriendo Windows con Microsoft Word instalado.
- Para produccion, cambie SIGE_SECRET_KEY por una clave segura del entorno.
