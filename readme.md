# Sistema de Gestión de Contratos (SGC)

Este proyecto es una Single Page Application (SPA) construida sobre Google Apps Script (GAS) implementando un patrón MVC (Model-View-Controller) modificado para GAS. Utiliza múltiples hojas de cálculo de Google como base de datos relacional y Tailwind CSS (vía CDN) / Bootstrap para la estilización e interfaz gráfica dinámica.

## 🏛 Arquitectura y Reglas del Proyecto

Para mantener el proyecto escalable y organizado, nos adherimos a una **Arquitectura Estricta de Archivos** y responsabilidades:

### 1. Lógica de Servidor (Backend en `.gs`)
Todo el código de backend en Google Apps Script (entorno V8).
*   **`code.gs`**: Controles y enrutamiento principal. Sirve la peticiones iniciales y coordina el MVC.
*   **`db_utils.gs`**: Operaciones CRUD y conexión con Google Sheets. Aquí reside la lógica de persistencia. **Regla de Eficiencia:** Todas las llamadas a la hoja de cálculo deben optimizarse usando lecturas y escrituras en bloque (procesar arrays gigantes en memoria) en lugar de bucles individuales para evitar límites de tiempo de ejecución de GAS.
*   **`agente_ia.gs`**: Lógica y endpoints del agente de IA interno utilizado para la extracción, validación y formateo de datos (e.g., Catálogo de Conceptos, CAF, Programas).
*   **`fix.gs`**: Funciones auxiliares o scripts de corrección/migración de datos de un solo uso o mantenimiento temporal.

### 2. Vistas (Frontend en `.html`)
Interfaces de usuario y HTML parciales que componen la SPA.
*   Archivos principales: `index.html`, `dashboard.vw.html`, `contratos.vw.html`, `contrato_detalle.vw.html`, `importacion_maestra.vw.html`.
*   Las vistas son inyectadas a demanda por el router de frontend.

### 3. Recursos del Cliente
*   **`main.js.html`**: Lógica centralizada del navegador (Vanilla JS / jQuery si aplica). Contiene el enrutador de front-end y controladores de vista.
*   **`styles.css.html`**: Diseño universal, utilidades, y overrides (como personalizaciones de Bootstrap, colores institucionales -ej. "vino"-, márgenes, etc.).

### 4. Reglas de Desarrollo Obligatorias
1.  **Inyección de Código:** Para mantener las vistas limpias, el CSS (`styles.css.html`) y JS (`main.js.html`) deben inyectarse incondicionalmente en el HTML principal (`index.html` u otros) utilizando templating de GAS: `<?!= include('nombre_del_archivo'); ?>`.
2.  **Comunicación Asíncrona Seguro:** Toda interacción entre el cliente (DOM/JavaScript) y el backend (`.gs`) se hace **exclusivamente a través de `google.script.run`**. Es imperativo implementar `withSuccessHandler` y `withFailureHandler` en *cada* llamada.
3.  **Flujo de Trabajo Controlado:** Antes de refactorizaciones que abarquen múltiples archivos, se requiere generar un plan de acción detallado para aprobación.

---

## 💼 Lógica de Negocios y Reglas Específicas

El sistema está diseñado para la administración integral del ciclo de vida de contratos de obras públicas/servicios. Sus directrices operativas fundamentales incluyen:

### A. Integridad Referencial y Relaciones de BD (`ESQUEMA_BD`)
*   La base de datos relacional simulada en Sheets depende estricta y obligatoriamente de sus llaves foráneas (`FK`), principalmente **`ID_Contrato`**.
*   Todas las operaciones de Inserción/Actualización a tablas hijas (Catálogo, Matriz de Insumos, Programas, Estimaciones) deben llevar consigo y preservar el `ID_Contrato` válido del registro padre.

### B. Flujo de Importación Inteligente (AI-Driven Imports)
*   **Orquestación Secuencial:** En `main.js.html`, la importación de documentos complejos vía IA (Contratos -> CAF -> Programas -> APUs) debe ser síncrona/secuencial usando `async/await`. 
*   **Paso de Contexto:** El `ID_Contrato` resultante de la creación de un nuevo contrato debe fluir y pasarse mediante un `parentContext` real hacia la IA para garantizar que todas las extracciones subsecuentes se vinculen correctamente y no generen "registros huérfanos".
*   Las tablas de **`Matriz_Insumos`** deben asociarse invariablemente al concepto correcto dentro de `Catalogo_Conceptos` por medio de su `ID_Contrato` e `ID_Concepto`.

### C. Módulo de Estimaciones y Porcentajes
*   **Cálculo de Avance Financiero:** Los avances acumulados y actuales en las estimaciones (`Detalle_Estimacion`) se deben capturar como fracciones/porcentajes limpios. El cálculo debe ser siempre en relación con el "Monto Total del Contrato" de la jerarquía superior.
*   **Actualizaciones Dinámicas:** Cualquier modificación a porcentajes en un "Programa de Ejecución" debe detonar métodos de re-cálculo automáticos en todo el flujo, soportado visualmente por spinners o bloqueos para conservar la congruencia de los totales antes de guardar en Base de Datos.
*   Los campos textuales definidos genéricamente (ej. "Periodo" en tablas de programa) deben preservarse estrictamente como cadenas de texto (`string`) para evitar que el motor de Sheets o Javascript aplique conversiones erróneas a "ISO Timestamp".

---

## 🗄 Diccionario de Datos Relacional Completo

A continuación el esquema normativo establecido para la estructura en Sheets (Las columnas deben respetarse en el estricto orden y nombre).

1.  **Usuarios_Sistema**: 
    - ID_Usuario (PK), Username, Nombre_Full, Rol, Email, Activo
2.  **Log_Actividad**: 
    - ID_Log (PK), ID_Usuario (FK), Accion, Tabla_Afectada, Timestamp, Detalles
3.  **Parametros_Sistema**: 
    - Clave_Parametro (PK), Valor_Parametro, Descripcion
4.  **Estadisticas_Financieras**: 
    - ID_Periodo (PK), Año, Mes, Monto_Ejecutado, Monto_Proyectado
5.  **Conversaciones_IA**: 
    - ID_Conversacion (PK), Fecha_Hora, Usuario, Prompt, Respuesta
6.  **Convenios_Recurso**: 
    - ID_Convenio (PK), Numero_Acuerdo, Nombre_Fondo, Monto_Apoyo, Fecha_Firma, Vigencia_Fin, Objeto_Programa, Estado, Link_Sharepoint
7.  **Contratistas**: 
    - ID_Contratista (PK), Razon_Social, RFC, Domicilio_Fiscal, Representante_Legal, Telefono, Banco, Cuenta_Bancaria, Cuenta_CLABE
8.  **Contratos**: 
    - ID_Contrato (PK), Numero_Contrato, ID_Convenio_Vinculado (FK), ID_Contratista (FK), Objeto_Contrato, Tipo_Contrato, Area_Responsable, No_Concurso, Modalidad_Adjudicacion, Fecha_Adjudicacion, Monto_Total_Sin_IVA, Monto_Total_Con_IVA, Fecha_Firma, Fecha_Inicio_Obra, Fecha_Fin_Obra, Plazo_Ejecucion_Dias, Porcentaje_Amortizacion_Anticipo, Porcentaje_Penas_Convencionales, No_Fianza_Cumplimiento, Monto_Fianza_Cumplimiento, No_Fianza_Anticipo, Monto_Fianza_Anticipo, No_Fianza_Garantia, Monto_Fianza_Garantia, No_Fianza_Vicios_Ocultos, Monto_Fianza_Vicios_Ocultos, Estado, Retencion_Vigilancia_Pct, Retencion_Garantia_Pct, Otras_Retenciones_Pct, Nombre_Residente_Dependencia, Link_Sharepoint, Pct_Indirectos_Oficina, Pct_Indirectos_Campo, Pct_Indirectos_Totales, Pct_Financiamiento, Pct_Utilidad, Pct_Cargos_SFP, Pct_ISN
9.  **Convenios_Modificatorios**: 
    - ID_Convenio_Mod (PK), ID_Contrato (FK), Numero_Convenio_Mod, Tipo_Modificacion, Nuevo_Monto_Con_IVA, Nueva_Fecha_Fin, Motivo, Link_Sharepoint
10. **Anticipos**: 
    - ID_Anticipo (PK), ID_Contrato (FK), Porcentaje_Otorgado, Monto_Anticipo, Fecha_Pago, Monto_Amortizado_Acumulado, Saldo_Por_Amortizar
11. **Catalogo_Conceptos**: 
    - ID_Concepto (PK), ID_Contrato (FK), Clave, Descripcion, Unidad, Cantidad_Contratada, Precio_Unitario, Importe_Total_Sin_IVA, Orden, Costo_Directo
12. **Programa**: 
    - ID_Numero_Programa (PK), ID_Contrato (FK), Tipo_Programa, Programa, Fecha_Inicio, Fecha_Termino
13. **Programa_Periodo**: 
    - ID_Programa_Periodo (PK), ID_Numero_Programa (FK), Numero_Periodo (Contador secuencial), Periodo (Descripción del mes), Fecha_Inicio, Fecha_Termino
14. **Programa_Ejecucion**: 
    - ID_Programa (FK), ID_Concepto (FK), ID_Programa_Periodo (FK), Fecha_Inicio, Fecha_Fin, Monto_Programado, Avance_Programado_Pct, Link_Sharepoint
15. **Estimaciones**: 
    - ID_Estimacion (PK), ID_Contrato (FK), No_Estimacion, Tipo_Estimacion, Periodo_Inicio, Periodo_Fin, Monto_Bruto_Estimado, Deduccion_Surv_05_Monto, Subtotal, IVA, Monto_Neto_A_Pagar, Avance_Acumulado_Anterior, Avance_Actual, Estado_Validacion, Link_Sharepoint, Avance_Anterior_Porcentaje, Avance_Actual_Porcentaje, Avance_Acumulado_Porcentaje
16. **Detalle_Estimacion**: 
    - ID_Detalle (PK), ID_Estimacion (FK), ID_Concepto (FK), Cantidad_Estimada_Periodo, Precio_Unitario_Contrato, Importe_Este_Periodo, Avance_Acumulado_Porcentaje, Importe_Acumulado, Avance_Periodo_Porcentaje
17. **Deducciones_Retenciones**: 
    - ID_Deduccion (PK), ID_Estimacion (FK), Tipo_Deduccion, Monto_Deducido, Concepto_Deduccion
18. **Facturas**: 
    - ID_Factura (PK), ID_Estimacion (FK), Folio_Fiscal_UUID, No_Factura, Fecha_Emision, Monto_Facturado, Estatus_SAT, Link_Sharepoint
19. **Pagos_Emitidos**: 
    - ID_Pago (PK), ID_Estimacion (FK), Fecha_Pago, Monto_Pagado, Referencia_Bancaria, Estatus_Pago
20. **Matriz_Insumos**:
    - ID_Matriz (PK), ID_Concepto (FK), Tipo_Insumo, Clave_Insumo, Descripcion, Unidad, Costo_Unitario, Rendimiento_Cantidad, Importe, Porcentaje_Incidencia
21. **Validacion_Archivos**:
    - ID_Validacion (PK), ID_Estimacion (FK), Tipo_Archivo, Fecha_Carga, Estado_Validacion, Checklist_JSON, Observaciones_Resumen

---

## 🔗 Dependencias y Relaciones Jerárquicas (`RELACIONES_BD`)

El sistema asume la siguiente dependencia estricta entre tablas para la propagación de datos y las importaciones:

*   **Nivel 1 (Hijos directos de Contratos):**
    *   `Catalogo_Conceptos` -> FK: `ID_Contrato`
    *   `Programa` -> FK: `ID_Contrato`
    *   `Estimaciones` -> FK: `ID_Contrato`
    *   `Anticipos` -> FK: `ID_Contrato`
    *   `Convenios_Modificatorios` -> FK: `ID_Contrato`
*   **Nivel 2:**
    *   `Matriz_Insumos` -> FK: `ID_Concepto` (Depende de Catalogo_Conceptos)
    *   `Programa_Periodo` -> FK: `ID_Numero_Programa` (Depende de Programa)
    *   `Detalle_Estimacion`, `Deducciones_Retenciones`, `Facturas`, `Pagos_Emitidos`, `Validacion_Archivos` -> FK: `ID_Estimacion` (Dependen de Estimaciones)
*   **Nivel 3:**
    *   `Programa_Ejecucion` -> FK: `ID_Programa_Periodo` (Depende de Programa_Periodo) y FK: `ID_Concepto` (Depende de Catalogo_Conceptos)
