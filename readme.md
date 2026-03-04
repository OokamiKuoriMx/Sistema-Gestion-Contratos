# Sistema de Gestión de Contratos (SGC)

Este proyecto es una Single Page Application (SPA) construida sobre Google Apps Script (GAS) implementando un patrón MVC (Model-View-Controller) modificado para GAS. Utiliza múltiples hojas de cálculo de Google como base de datos relacional y Tailwind CSS (vía CDN) para la estilización e interfaz gráfica dinámica.

## Arquitectura

*   **Backend (`code.gs`)**: Contiene los controladores y modelos. Interactúa con las hojas del spreadsheet de Google Apps, sirviendo a las peticiones usando `doGet` para iniciar la aplicación mediante una plantilla HTML, y expone funciones globales (`google.script.run`) con encapsulamiento en try-catch robusto, enviando siempre respuestas estructuradas al cliente JSON-Like: `{ success: true/false, data, error, functionName }`.
*   **Frontend**:
    *   `index.html`: Layout envoltorio de la aplicación. Contiene un Overlay Spinner Central (`#global-spinner`) para bloquear interacciones del usuario durante de solicitudes asincrónicas, y un `<div id="app-shell">` para montaje dinámico.
    *   `main.js.html`: Implementa el router de frontend (Lazy Loading) capturando Vistas parciales desde el servidor a demanda. Manejador de error global con impresiones de consola estructuradas.
    *   `styles.css.html`: Reglas complementarias para scrollbars, overlays y z-indexing superior.
    *   **Vistas (`*.vw.html`)**: Plantillas inyectables para tableros y módulos de operaciones.

## Diccionario de Datos Relacional Completo

A continuación el esquema normativo establecido para la estructura en Sheets (Las columnas deben respetarse en el estricto orden y nombre).

1.  **Usuarios_Sistema**: 
    - ID_Usuario (PK)
    - Username
    - Nombre_Full
    - Rol
    - Email
    - Activo
2.  **Log_Actividad**: 
    - ID_Log (PK)
    - ID_Usuario (FK)
    - Accion
    - Tabla_Afectada
    - Timestamp
    - Detalles
3.  **Parametros_Sistema**: 
    - Clave_Parametro (PK)
    - Valor_Parametro
    - Descripcion
4.  **Estadisticas_Financieras**: 
    - ID_Periodo (PK)
    - Año
    - Mes
    - Monto_Ejecutado
    - Monto_Proyectado
5.  **Conversaciones_IA**: 
    - ID_Conversacion (PK)
    - Fecha_Hora
    - Usuario
    - Prompt
    - Respuesta
6.  **Convenios_Recurso**: 
    - ID_Convenio (PK)
    - Numero_Acuerdo
    - Nombre_Fondo
    - Monto_Apoyo
    - Fecha_Firma
    - Vigencia_Fin
    - Objeto_Programa
    - Estado
    - Link_Sharepoint
7.  **Contratistas**: 
    - ID_Contratista (PK)
    - Razon_Social
    - RFC
    - Domicilio_Fiscal
    - Representante_Legal
    - Telefono
    - Banco
    - Cuenta_Bancaria
    - Cuenta_CLABE
8.  **Contratos**: 
    - ID_Contrato (PK)
    - Numero_Contrato
    - ID_Convenio_Vinculado (FK)
    - ID_Contratista (FK)
    - Objeto_Contrato
    - Monto_Total_Sin_IVA
    - Monto_Total_Con_IVA
    - Fecha_Firma
    - Fecha_Inicio_Obra
    - Fecha_Fin_Obra
    - Estado
    - Retencion_Vigilancia_Pct
    - Retencion_Garantia_Pct
    - Otras_Retenciones_Pct
    - Link_Sharepoint
9.  **Convenios_Modificatorios**: 
    - ID_Convenio_Mod (PK)
    - ID_Contrato (FK)
    - Numero_Convenio_Mod
    - Tipo_Modificacion
    - Nuevo_Monto_Con_IVA
    - Nueva_Fecha_Fin
    - Motivo
    - Link_Sharepoint
10. **Anticipos**: 
    - ID_Anticipo (PK)
    - ID_Contrato (FK)
    - Porcentaje_Otorgado
    - Monto_Anticipo
    - Fecha_Pago
    - Monto_Amortizado_Acumulado
    - Saldo_Por_Amortizar
11. **Catalogo_Conceptos**: 
    - ID_Concepto (PK)
    - ID_Contrato (FK)
    - Clave
    - Descripcion
    - Unidad
    - Cantidad_Contratada
    - Precio_Unitario
    - Importe_Total_Sin_IVA
12. **Programa**: 
    - ID_Numero_Programa (PK)
    - ID_Contrato (FK)
    - Tipo_Programa
    - Fecha_Inicio
    - Fecha_Termino
13. **Programa_Periodo**: 
    - ID_Programa_Periodo (PK)
    - ID_Numero_Programa (FK)
    - Numero_Periodo (Contador secuencial 1, 2, 3...)
    - Periodo (Descripción/Nombre del mes)
    - Fecha_Inicio
    - Fecha_Termino
14. **Programa_Ejecucion**: 
    - ID_Programa (FK - ID_Numero_Programa)
    - ID_Concepto (FK)
    - ID_Programa_Periodo (FK)
    - Fecha_Inicio
    - Fecha_Fin
    - Monto_Programado
    - Avance_Programado_Pct
    - Link_Sharepoint
15. **Estimaciones**: 
    - ID_Estimacion (PK)
    - ID_Contrato (FK)
    - No_Estimacion
    - Tipo_Estimacion
    - Periodo_Inicio
    - Periodo_Fin
    - Monto_Bruto_Estimado
    - Deduccion_Surv_05_Monto
    - Subtotal
    - IVA
    - Monto_Neto_A_Pagar
    - Avance_Acumulado_Anterior
    - Avance_Actual
    - Estado_Validacion
    - Link_Sharepoint
16. **Detalle_Estimacion**: 
    - ID_Detalle (PK)
    - ID_Estimacion (FK)
    - ID_Concepto (FK)
    - Cantidad_Estimada_Periodo
    - Precio_Unitario_Contrato
    - Importe_Este_Periodo
    - Avance_Acumulado_Porcentaje
    - Importe_Acumulado
17. **Deducciones_Retenciones**: 
    - ID_Deduccion (PK)
    - ID_Estimacion (FK)
    - Tipo_Deduccion
    - Monto_Deducido
    - Concepto_Deduccion
18. **Facturas**: 
    - ID_Factura (PK)
    - ID_Estimacion (FK)
    - Folio_Fiscal_UUID
    - No_Factura
    - Fecha_Emision
    - Monto_Facturado
    - Estatus_SAT
    - Link_Sharepoint
19. **Pagos_Emitidos**: 
    - ID_Pago (PK)
    - ID_Estimacion (FK)
    - Fecha_Pago
    - Monto_Pagado
    - Referencia_Bancaria
    - Estatus_Pago
