Siniestralidad: Dashboard de Concentración de Riesgos

Herramienta analítica construida con Python y JavaScript para la evaluación espacial de carteras de seguros. El proyecto procesa bases de datos de pólizas, estandariza variables geográficas y genera un reporte HTML interactivo con lógica de en el cliente.

Aplicación
Identificación de zonas críticas (Riesgo Alto) mediante mapas coropléticos.
Monitoreo del cúmulo de sumas aseguradas por estado y ramo (Autos, Gastos Médicos, Vida, Daños) para mitigar el impacto de eventos catastróficos.
Transformación de datos tabulares crudos en un panel visual, responsivo e interactivo.

Arquitectura Técnica
El proyecto utiliza un enfoque híbrido de inyección de datos:
1. Python & Pandas: Se encarga de la lectura, limpieza y procesamiento del dataset, transformando la estructura a formato JSON.
2. JS, Plotly & DataTables: Renderiza los gráficos e implementa un sistema de filtros cruzados. Al hacer clic en cualquier elemento (un estado en el mapa, un nivel de riesgo o un ramo de seguro), toda la interfaz y la tabla de datos subyacente se recalculan en tiempo real.

Instalación y Uso

1. Clonar el repositorio:
   ```bash
   git clone [https://github.com/TuUsuario/NombreDelRepo.git](https://github.com/TuUsuario/NombreDelRepo.git)
   cd NombreDelRepo

Instalar dependencias de Python:
pip install pandas plotly openpyxl

Ejecutar el pipeline:
El repositorio incluye un archivo bd_siniestralidad.xlsx con registros ficticios.

python generador_dashboard.py
Visualización:
El script compilará el archivo Dashboard_Siniestralidad.html. Ábrelo en cualquier navegador moderno para interactuar con los datos.