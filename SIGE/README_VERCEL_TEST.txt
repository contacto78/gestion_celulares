AMBIENTE DE PRUEBAS EN VERCEL

Esta configuracion permite desplegar la version web Flask en vercel.app como
ambiente de pruebas.

Archivos agregados:
- api/index.py: entrypoint serverless para Vercel.
- vercel.json: rutas hacia la funcion Python.
- requirements.txt: dependencias detectables por Vercel.

IMPORTANTE
- En Vercel, la base Excel se copia a /tmp/sige en tiempo de ejecucion.
- Los cambios realizados desde la web son temporales y pueden perderse cuando
  Vercel recicle la funcion.
- La generacion de PDF de la version web usa ReportLab y funciona en Vercel.
  La plantilla Word queda solo para la aplicacion de escritorio original.
- Este ambiente sirve para validar navegacion, login, vistas, formularios,
  generacion de PDF y comportamiento general. Para produccion se debe migrar
  Excel/PDFs a servicios persistentes.

DESPLIEGUE
1. Instale Vercel CLI si no lo tiene:
   npm i -g vercel

2. Desde la carpeta raiz del proyecto, ejecute:
   vercel

3. Para publicar en produccion:
   vercel --prod

VARIABLES RECOMENDADAS
- SIGE_SECRET_KEY: clave secreta de Flask para sesiones.
- SIGE_DATA_DIR: opcional. En Vercel se usa /tmp/sige por defecto.
