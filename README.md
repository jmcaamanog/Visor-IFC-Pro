# Visor IFC Pro 🏢🔍

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![BIM](https://img.shields.io/badge/BIM-IFC-success.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

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
