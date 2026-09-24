GESTIÓN CORPORATIVA DE CELULARES - V20.2

NUEVO EN V20.2
- Nueva pestaña "Actas de Entrega".
- Columnas: Fecha, Usuario, Nombre, DNI, Fecha de Entrega, Nombre del documento generado.
- El usuario que genera el acta queda registrado.
- Doble clic sobre un registro para abrir el PDF generado.
- El historial de actas se guarda también en la hoja Excel "Actas de Entrega".
- Los registros históricos de la hoja anterior "Actas" sin usuario se migran como "No registrado"; no se inventa información histórica.

REQUISITOS
- Windows 10/11.
- Python 3.13 (o versión compatible con las dependencias instaladas).
- Microsoft Word instalado y activado para generar los PDF desde la plantilla.
- openpyxl.

ARCHIVOS PRINCIPALES
- app_celulares.py
- Base_Datos_Celulares.xlsx
- Acta de Entrega.docx

USO
1. Verifique que Base_Datos_Celulares.xlsx y Acta de Entrega.docx estén junto a app_celulares.py.
2. Ejecute app_celulares.py.
3. Genere un acta desde una asignación activa.
4. Abra la pestaña "Actas de Entrega".
5. Haga doble clic sobre el acta para abrir el PDF.

NOTA
La plantilla Word original no se modifica. El sistema crea una copia temporal, reemplaza los marcadores y genera el PDF.
