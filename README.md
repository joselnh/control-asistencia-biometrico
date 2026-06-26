# Pipeline ETL Automatizado y Control de Asistencia Biométrico

## Descripción del Proyecto
Este software resuelve el problema de los formatos propietarios y datos en la gestión de recursos humanos. Actúa como una capa de extracción y transformación (ETL) que toma los reportes planos `.xls` exportados por un reloj biométrico físico, analiza dinámicamente sus datos y consolida la información en un modelo de datos horizontal, limpio y protegido en Excel.

El sistema cuenta con una interfaz gráfica (GUI) para facilitar su uso por personal no técnico (Recursos Humanos), automatizando tareas que antes requerían horas de procesamiento manual y que era propenso a errores.

## Características Clave
- **Algoritmo de Búsqueda:** Localiza de forma dinámica el periodo del reporte y las matrices de los empleados mediante reconocimiento de patrones, haciendo al software inmune a cambios de cabeceras o filas en blanco en el archivo de origen.
- **Motor de Reglas de Negocio:** Calcula con precisión (usando deltas de tiempo) las horas extras y horas faltantes, aplicando ventanas de tolerancia de acuerdo a reglamento de la empresa y discriminando los horarios de lunes a viernes, sábados y domingos.
- **Seguridad y Control de Cambios:** Mediante la API de `openpyxl`, el script inyecta fórmulas nativas de Excel y aplica una protección selectiva: bloquea las celdas automáticas estables y desbloquea/pinta de rojo únicamente las celdas vacías para permitir la justificación manual de incidencias.

## Tecnologías Utilizadas
- **Python 3.11**
- **Pandas** (Procesamiento y reestructuración de matrices de datos)
- **OpenPyXL** (Estilos, inyección de fórmulas nativas y seguridad de hojas)
- **Tkinter** (Interfaz gráfica de usuario de escritorio)

## Roadmap (Futuras Mejoras)
- [ ] Implementar un sistema de logs de auditoría (`logging`) para soporte técnico y control de excepciones.
- [ ] Desarrollar un módulo de persistencia local ligera (`SQLite`) para almacenar el histórico anual.
- [ ] Conectar la base de datos a un Dashboard ejecutivo en Power BI para la analítica avanzada de ausentismo y costos de horas extra en la planta.

## Requisitos del Archivo de Entrada (Input)

Para que el pipeline procese la información correctamente, el archivo que se ingrese debe cumplir con las siguientes especificaciones técnicas basadas en las exportaciones nativas del hardware biométrico:

1. **Formato del archivo:** Debe ser obligatoriamente un archivo de Excel con extensión `.xls` (formato heredado compatible con el motor `xlrd`) o `.xlsx`.
2. **Nombre de la pestaña:** La hoja de cálculo donde están los datos debe llamarse exactamente `Reporte de Asistencia`.
3. **Estructura Interna Estricta (Estructura de Matrices):**
   - **Fila de Período:** El script busca de forma heurística la etiqueta `Periodo:` dentro de las primeras 10 filas para extraer dinámicamente el año y mes en formato `AAAA-MM` (ej. `2026-05`).
   - **Fila de Días:** Se localiza mediante reconocimiento de patrones buscando la secuencia numérica horizontal `[1, 2, 3, 4, 5]` que representa los días del mes.
   - **Bloques de Empleados:** Cada trabajador debe estar estructurado en un bloque vertical de dos filas:
     - **Fila 1:** Contiene las celdas identificadoras `ID:` y `Nombre:` seguidas de sus respectivos valores.
     - **Fila 2 (Marcaciones):** Contiene los strings de tiempo correspondientes a cada día del mes (ej. `08:00 17:30` en formato `HH:MM`).

> **Nota:** Si deseas probar el funcionamiento del pipeline de inmediato sin procesar datos reales, puedes utilizar el archivo limpio de estructura de pruebas incluido en la documentación del proyecto.