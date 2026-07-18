# Visor IFC Pro 🏢🔍

![Profesión](https://img.shields.io/badge/Profesi%C3%B3n-Arquitectos%20T%C3%A9cnicos-2e7d32?logo=micro%3Abit&logoColor=white&style=flat-square)
![Role](https://img.shields.io/badge/Role-BIM%20%26%20ConTech-007ACC?logo=bim360&style=flat-square)
![Location](https://img.shields.io/badge/Location-A%20Coru%C3%B1a%20%F0%9F%8C%8A-005B94?logo=lighthouse&logoColor=white&style=flat-square)
![Sector](https://img.shields.io/badge/Sector-ConTech%20%7C%20AECO-E65100?logo=construct3&style=flat-square)
![BIM](https://img.shields.io/badge/BIM-IFC%20%2F%20openBIM-009688?style=flat-square)
![Maker](https://img.shields.io/badge/Maker-Software-red?logo=makerbot&style=flat-square)
![Hardware](https://img.shields.io/badge/Hardware---grey?style=flat-square)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&style=flat-square)
![Language](https://img.shields.io/badge/Language-Python-3776AB?logo=python&logoColor=white&style=flat-square)
![Stars](https://img.shields.io/github/stars/jmcaamanog/Visor-IFC-Pro?style=flat-square&color=yellow&logo=github)
![License](https://img.shields.io/github/license/jmcaamanog/Visor-IFC-Pro?style=flat-square&color=green)

(Arquitecto Técnico_JMC) Herramienta de escritorio avanzada para la auditoría, visualización 3D y edición bidireccional de modelos BIM en formato IFC. Diseñada desde la trinchera para agilizar el trabajo de perfiles técnicos en el sector AECO.

## 🚀 Características Principales

* **Árbol de Proyecto Estructurado:** Navegación jerárquica fluida por todos los elementos constructivos del modelo IFC.
* **Visor 3D Integrado:** Renderizado del modelo completo o de elementos aislados, incluyendo modos de aislamiento y transparencia.
* **Edición Bidireccional:** Capacidad de lectura y modificación de parámetros (Property Sets) con guardado directo en un nuevo archivo IFC modificado.
* **Dashboards Analíticos:** Cálculo automático de áreas, volúmenes, recuento de elementos y estimación de costes, visualizados mediante gráficos integrados.
* **Exportación Profesional:** Generación de resúmenes de datos en formato CSV y volcado de informes completos tabulados a Excel (`.xlsx`).
* **Búsqueda Avanzada:** Filtrado de elementos por nombre, ID, tipo IFC o propiedades específicas.

## 🛠️ Stack Tecnológico

El proyecto está construido con herramientas potentes y optimizadas para el manejo de datos BIM y renderizado 3D:

* **IfcOpenShell:** Motor principal para la lectura de la estructura IFC, extracción de Psets y edición directa de atributos.
* **Vedo:** Librería encargada de la renderización y visualización 3D de la geometría de los elementos.
* **ttkbootstrap:** Chasis de la interfaz gráfica, proporcionando un aspecto moderno y profesional.
* **Matplotlib & Pandas:** Cerebro analítico para el procesamiento de recuentos y generación de gráficos (barras y anillos).
* **Openpyxl:** Pasarela para empaquetar los datos estructurados e imágenes gráficas en informes Excel formateados.

## ⚙️ Requisitos e Instalación

1. Clona este repositorio:
   ```bash
   git clone [https://github.com/TU_USUARIO/visor-ifc-pro.git](https://github.com/TU_USUARIO/visor-ifc-pro.git)
   
2. Navega al directorio del proyecto:
   ```bash
   pip install ifcopenshell vedo matplotlib pandas openpyxl ttkbootstrap Pillow numpy

3. Ejecuta la aplicación:
    ```bash
    python visor_ifc_final.py


## 👨‍💻 Autor

Jose Manuel Caamaño González | Arquitecto Técnico & BIM Manager.
Digital Product Lead | ConTech & Digital Twin SaaS | BIM, Energy Modeling & Sustainability | Data Analytics (SQL, Power BI)

Hecho con código y café desde A Coruña. ☕

Jose Manuel Caamaño González | [LinkedIn](https://www.linkedin.com/in/jmcaamanog/)

Acuérdate de cambiar `TU_USUARIO` en el bloque de instalación por el tuyo de GitHub. Pégalo en tu VS Code, haz el "Commit" y "Sync Changes" que te comenté antes, y lo tendrás brillando en la web en cuestión de segundos. ¡Dale caña!
