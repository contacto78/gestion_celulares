PORTAL DE GESTIÓN - V20.4

1. Pantalla inicial
Al ejecutar Launcher.py se muestra una pantalla previa con cuatro opciones:
- Gestion Celular
- Gestión Inmobiliario
- Gestión de Taxi
- Gestión de Estacionamiento

Cada opción tiene un ícono conceptual sobre su botón.

2. Gestión Celular
El botón Gestion Celular ejecuta:
Gestion_Celular\app_celulares.py

La ruta NO está fijada a una unidad ni a una carpeta absoluta. El launcher determina
la carpeta donde está ubicado el propio Launcher.py (o el ejecutable si se compila)
y desde allí construye la ruta relativa Gestion_Celular\app_celulares.py.

Para ejecutar el .py, el launcher utiliza el intérprete Python disponible en el entorno.
Si el launcher se distribuye como ejecutable, busca primero Python/runtime junto al
launcher y luego Python disponible en PATH.

3. Estructura recomendada
Portal/
  Launcher.py
  Gestion_Celular/
    app_celulares.py
    Base_Datos_Celulares.xlsx
    Acta de Entrega.docx
    README.txt

La carpeta Portal puede ubicarse en cualquier unidad o ruta (C:, D:, red, etc.).
No se requiere modificar una ruta dentro del código.

4. Otras opciones
Los botones Gestión Inmobiliario, Gestión de Taxi y Gestión de Estacionamiento
quedan preparados como accesos para futuras aplicaciones.
