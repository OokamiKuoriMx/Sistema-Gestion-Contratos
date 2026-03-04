// ---------------------------------------------------------
// CONFIGURACIÓN DEL AGENTE IA SGC (GOOGLE APPS SCRIPT)
// ---------------------------------------------------------
// Reemplazar por tu clave de API de Google Gemini obtenida en Google AI Studio
var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

if (!GEMINI_API_KEY) {
    console.error("ERROR: La propiedad GEMINI_API_KEY no está configurada en el Proyecto.");
}

/**
 * Normaliza un texto para comparaciones (quita acentos, puntuación, espacios extra y convierte a minúsculas)
 */
function normalizeText(text) {
    if (!text) return "";
    return text.toString()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // Quitar acentos
        .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "") // Quitar puntuación
        .replace(/\s+/g, " ") // Colapsar espacios
        .trim();
}

const PROMPT_IMPORTACION_CONTRATOS = `
Rol y Objetivo:
Eres un Agente de Inteligencia IA experto en el Sistema de Gestión de Contratos (SGC). Tu misión es procesar documentos (PDF, Imágenes, Excel) para extraer datos estructurados y vincularlos con la base de datos de Google Sheets.

Estructura de la Base de Datos (Esquema Completo):
1. **Convenios_Recurso**: ID_Convenio, Numero_Acuerdo, Nombre_Fondo, Monto_Apoyo, Fecha_Firma, Vigencia_Fin, Objeto_Programa, Estado, Link_Sharepoint.
2. **Contratistas**: ID_Contratista, Razon_Social, RFC, Domicilio_Fiscal, Representante_Legal, Telefono, Banco, Cuenta_Bancaria, Cuenta_CLABE.
3. **Contratos**: ID_Contrato, Numero_Contrato, ID_Convenio_Vinculado, ID_Contratista, Objeto_Contrato, Tipo_Contrato, Area_Responsable, No_Concurso, Modalidad_Adjudicacion, Fecha_Adjudicacion, Monto_Total_Sin_IVA, Monto_Total_Con_IVA, Fecha_Firma, Fecha_Inicio_Obra, Fecha_Fin_Obra, Plazo_Ejecucion_Dias, Porcentaje_Amortizacion_Anticipo, Porcentaje_Penas_Convencionales, No_Fianza_Cumplimiento, Monto_Fianza_Cumplimiento, No_Fianza_Anticipo, Monto_Fianza_Anticipo, No_Fianza_Garantia, Monto_Fianza_Garantia, No_Fianza_Vicios_Ocultos, Monto_Fianza_Vicios_Ocultos, Estado, Retencion_Vigilancia_Pct, Retencion_Garantia_Pct, Otras_Retenciones_Pct, Nombre_Residente_Dependencia, Link_Sharepoint.
4. **Convenios_Modificatorios**: ID_Convenio_Mod, ID_Contrato, Numero_Convenio_Mod, Tipo_Modificacion, Nuevo_Monto_Con_IVA, Nueva_Fecha_Fin, Motivo, Link_Sharepoint.
5. **Anticipos**: ID_Anticipo, ID_Contrato, Porcentaje_Otorgado, Monto_Anticipo, Fecha_Pago, Monto_Amortizado_Acumulado, Saldo_Por_Amortizar.
6. **Catalogo_Conceptos**: ID_Concepto, ID_Contrato, Clave, Descripcion, Unidad, Cantidad_Contratada, Precio_Unitario, Importe_Total_Sin_IVA, Orden.
7. **Programa**: ID_Numero_Programa, ID_Contrato, Tipo_Programa, Fecha_Inicio, Fecha_Termino.
8. **Programa_Periodo**: ID_Programa_Periodo, ID_Numero_Programa, Numero_Periodo, Periodo, Fecha_Inicio, Fecha_Termino.
9. **Programa_Ejecucion**: ID_Programa, ID_Concepto, ID_Programa_Periodo, Fecha_Inicio, Fecha_Fin, Monto_Programado, Avance_Programado_Pct, Link_Sharepoint.
10. **Estimaciones**: ID_Estimacion, ID_Contrato, No_Estimacion, Tipo_Estimacion, Periodo_Inicio, Periodo_Fin, Monto_Bruto_Estimado, Deduccion_Surv_05_Monto, Subtotal, IVA, Monto_Neto_A_Pagar, Avance_Acumulado_Anterior, Avance_Actual, Estado_Validacion, Link_Sharepoint.
11. **Detalle_Estimacion**: ID_Detalle, ID_Estimacion, ID_Concepto, Cantidad_Estimada_Periodo, Precio_Unitario_Contrato, Importe_Este_Periodo, Avance_Acumulado_Porcentaje, Importe_Acumulado.
12. **Deducciones_Retenciones**: ID_Deduccion, ID_Estimacion, Tipo_Deduccion, Monto_Deducido, Concepto_Deduccion.
13. **Facturas**: ID_Factura, ID_Estimacion, Folio_Fiscal_UUID, No_Factura, Fecha_Emision, Monto_Facturado, Estatus_SAT, Link_Sharepoint.
14. **Pagos_Emitidos**: ID_Pago, ID_Estimacion, Fecha_Pago, Monto_Pagado, Referencia_Bancaria, Estatus_Pago.

Funciones de Negocio Disponibles:
- 'getContratosData()', 'upsertConcepto(obj)', 'getProgramaEjecucion(idContrato)', 'actualizarPeriodoPrograma(...)'.

Reglas Críticas:
- **Datos de Origen Legal**: Captura No. de Concurso, Modalidad de Adjudicación y Área Responsable (Primera página).
- **Control Financiero**: Extrae Porcentajes de Amortización, Penas Convencionales y Números/Montos de Fianza (Cumplimiento, Anticipo, Garantía, Vicios Ocultos).
- **Anticipos**: Si el documento NO menciona anticipo o es 0, establece 'Porcentaje_Amortizacion_Anticipo' como 0 o null. No inventes datos.
- **Plazos**: Identifica el Plazo de Ejecución en Días Naturales.
- **Vínculos**: Usa IDs temporales (ej. "temp_1") para relacionar registros nuevos.
- **Formato**: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

const PROMPT_IMPORTACION_MATRICES = `
Rol y Objetivo:
Eres un Experto en Ingeniería de Costos para el Sistema de Gestión de Contratos (SGC). Tu misión es procesar documentos de "Análisis de Precios Unitarios (Matrices)" y extraer la estructura técnica y económica completa.

Estructura de Extracción (Tablas):
1. **Contratos**: Extraer porcentajes encontrados: 'Pct_Indirectos_Oficina', 'Pct_Indirectos_Campo', 'Pct_Financiamiento', 'Pct_Utilidad', 'Pct_Cargos_SFP' (5 al millar), 'Pct_ISN'.
2. **Catalogo_Conceptos**: Para CADA MATRIZ leída, extraer: 'Clave', 'Descripcion', 'Unidad', 'Costo_Directo' (Suma de insumos), 'Precio_Unitario' (Precio final después de sobrecostos). 
   - **Orden**: Genera un número secuencial (1, 2, 3...) según el orden de aparición en el documento.
   - **Cantidad_Contratada**: SIEMPRE establece este valor en 1 para todas las importaciones de matrices.
3. **Matriz_Insumos**: Extraer el desglose de CADA concepto: 
   - **Tipo_Insumo**: Clasificación numérica OBLIGATORIA:
     - **1 (MATERIALES)**: Cemento, varilla, cables, pintura, consumibles, etc.
     - **2 (MANO DE OBRA)**: Gerentes, Especialistas, Técnicos, Cuadrillas, Ayudantes, JORNALES, MO.
     - **3 (EQUIPO Y HERRAMIENTA)**: Vehículos, Herramienta menor, Maquinaria pesada, Computadoras, Equipos de medición, EQ.
   - 'Clave_Insumo', 'Descripcion', 'Unidad', 'Costo_Unitario', 'Rendimiento_Cantidad', 'Importe', 'Porcentaje_Incidencia'.

Reglas Críticas de Clasificación y Clasificación:
- **Identificación por Contexto**: Si los insumos están agrupados bajo un título (ej: "MANO DE OBRA" o "SUBTOTAL DE MATERIALES"), asigna el tipo correspondiente a todos los registros del bloque aunque no se mencione por fila.
- **Identificación por Palabra Clave**: Analiza la descripción o clave. Ej: "MO..." o "Cabo" o "Cuadrilla" -> Tipo 2. "Maquinaria" o "EQ..." o "Hora-Máquina" -> Tipo 3.
- **Matching de Catálogo**: Se te proporcionará un "Catálogo Actual" en el prompt adicional. SI la descripción del concepto que estás leyendo coincide contextualmente (aunque varíen espacios o acentos) con uno del catálogo, DEBES usar su 'Clave' original.
- **Hierarquía**: Cada 'Matriz_Insumo' debe estar vinculada al concepto correspondiente mediante IDs temporales.
- **Sobrecostos**: Si el documento muestra un porcentaje de utilidad o indirectos, regístralo en la tabla 'Contratos'.
- **Precisión Numérica**: Captura los rendimientos y costos con todos sus decimales.
- **Formato**: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

/**
 * Función llamada desde el frontend para leer el archivo con Gemini y guardar
 */
function processDocumentWithAI(base64Data, mimeType, targetTable = null, parentContext = null) {
    try {
        if (!base64Data || !mimeType) {
            throw new Error("Datos del archivo o MimeType faltantes.");
        }

        // 1. Preparar payload para Gemini
        let b64 = base64Data;
        if (b64.includes('base64,')) {
            b64 = b64.split('base64,')[1];
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

        // Seleccionar el System Instruction basado en el objetivo
        let systemPrompt = PROMPT_IMPORTACION_CONTRATOS;
        if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') {
            systemPrompt = PROMPT_IMPORTACION_MATRICES;
        }

        let promptAdicional = `Lee este archivo oficial. Debes generar estrictamente un objeto JSON estructurado con arrays para las tablas correspondientes basándote en la información encontrada.`;
        if (targetTable) {
            promptAdicional += `\n\nATENCIÓN: El usuario subió este documento desde el contexto de la tabla '${targetTable}'.`;
            if (parentContext) {
                promptAdicional += `\n\nCONTEXTO ADICIONAL: El registro padre tiene esta información: ${JSON.stringify(parentContext)}. Asegúrate de usar estos IDs para vincular los registros hijo que encuentres.`;

                // Si es matrices, pasar el catálogo actual para matching por descripción
                if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') {
                    const idContrato = parentContext.ID_Contrato || parentContext.idContrato;
                    if (idContrato) {
                        const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: idContrato });
                        if (catalogo && catalogo.length > 0) {
                            promptAdicional += `\n\nIMPORTANTE - CATÁLOGO EXISTENTE: A continuación se listan los conceptos que YA ESTÁN en la base de datos para este contrato. 
                            SI la descripción del concepto en el documento coincide contextualmente con alguno de estos, USA SU ID_Concepto Y CLAVE ORIGINAL. NO inventes nuevas claves si puedes hacer match.
                            Catálogo actual: ${JSON.stringify(catalogo.map(c => ({ ID_Concepto: c.ID_Concepto, Clave: c.Clave, Descripcion: c.Descripcion })))}`;
                        }
                    }
                }
            }
        }

        const payload = {
            system_instruction: {
                parts: [{ text: systemPrompt }]
            },
            contents: [
                {
                    parts: [
                        {
                            inlineData: {
                                mimeType: mimeType,
                                data: b64
                            }
                        },
                        {
                            text: promptAdicional
                        }
                    ]
                }
            ],
            generationConfig: {
                temperature: 0.0,
                responseMimeType: "application/json"
            }
        };

        const options = {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        // 2. Enviar a Gemini API usando el servicio REST de Apps Script
        const response = UrlFetchApp.fetch(url, options);
        const resultText = response.getContentText();
        const resultJson = JSON.parse(resultText);

        if (response.getResponseCode() !== 200) {
            throw new Error("HTTP " + response.getResponseCode() + " - " + (resultJson.error?.message || resultText));
        }

        // 3. Extraer respuesta (Es JSON texto, así que hay que parsearlo)
        const llmJsonString = resultJson.candidates[0].content.parts[0].text;
        const datosExtraidos = JSON.parse(llmJsonString);

        if (datosExtraidos.accion === "importar_datos" && datosExtraidos.datos) {
            return generarRespuesta(true, datosExtraidos.datos);
        } else {
            throw new Error("El formato JSON devuelto por la IA no corresponde al solicitado.");
        }

    } catch (e) {
        console.error(e);
        return generarRespuesta(false, e.toString(), 'processDocumentWithAI');
    }
}

/**
 * Función auxiliar para insertar y actualizar los datos en las pestañas como hace el webhook.
 * Realiza un chequeo de existencia (UPSERT) para evitar records duplicados desde la IA.
 */
function guardarDatosIA(datos) {
    let resultados = { insertados: 0, actualizados: 0 };

    // Jerarquía de inserción (Padres primero, Hijos después) para resolver llaves foráneas
    const jerarquia = [
        'Convenios_Recurso',
        'Contratistas',
        'Contratos',
        'Convenios_Modificatorios',
        'Anticipos',
        'Catalogo_Conceptos',
        'Programa',
        'Programa_Periodo',
        'Programa_Ejecucion',
        'Estimaciones',
        'Facturas',
        'Deducciones_Retenciones',
        'Pagos_Emitidos',
        'Detalle_Estimacion',
        'Matriz_Insumos'
    ];

    // Mapa de traducción de IDs: Si la IA generó un "ID_Contrato": "temp_1", y dbInsert le asignó el entero 15, guardamos { "temp_1": 15 }
    const mapaIds = {};

    jerarquia.forEach(tabla => {
        if (datos[tabla] && Array.isArray(datos[tabla])) {
            const headers = ESQUEMA_BD[tabla];
            const pkName = headers[0];

            // --- LÓGICA DE PURGA PARA MATRICES (SUSTITUCIÓN) ---
            if (tabla === 'Matriz_Insumos' && typeof dbDelete === 'function') {
                const conIds = [...new Set(datos[tabla].map(ins => ins.ID_Concepto).filter(id => id && !String(id).startsWith('temp_')))];
                conIds.forEach(idC => {
                    dbDelete('Matriz_Insumos', { ID_Concepto: idC });
                });
            }

            let skName = null;
            if (tabla === 'Contratistas') skName = 'RFC';
            else if (tabla === 'Contratos') skName = 'Numero_Contrato';
            else if (tabla === 'Facturas') skName = 'Folio_Fiscal_UUID';
            else if (tabla === 'Estimaciones') skName = 'No_Estimacion';
            else if (tabla === 'Convenios_Recurso') skName = 'Numero_Acuerdo';

            datos[tabla].forEach(registro => {
                // --- SANITIZACIÓN DE CAMPOS (MAPPING FALLBACK) ---
                if (tabla === 'Catalogo_Conceptos') {
                    if (registro.Clave_Concepto && !registro.Clave) registro.Clave = registro.Clave_Concepto;
                    if (registro.Descripcion_Concepto && !registro.Descripcion) registro.Descripcion = registro.Descripcion_Concepto;
                    if (registro.Unidad_Medida && !registro.Unidad) registro.Unidad = registro.Unidad_Medida;
                }

                // 1. Reemplazar llaves foráneas con literales reales usando nuestro mapa (ej. Contrato "temp_1" -> 15)
                Object.keys(registro).forEach(k => {
                    if (k.startsWith('ID_') && k !== pkName && registro[k]) {
                        // Si es un temporal, buscar en mapa
                        if (mapaIds[registro[k]]) {
                            registro[k] = mapaIds[registro[k]];
                        }
                    }
                });

                // --- INYECCIÓN DE CONTEXTO (SI FALTA ID_CONTRATO) ---
                if (!registro.ID_Contrato && datos.ID_Contrato_Contexto) {
                    registro.ID_Contrato = datos.ID_Contrato_Contexto;
                }

                if (tabla === 'Estimaciones' && registro.ID_Contrato) {
                    // Fetch contract to get its retention parameters
                    const contrato = dbSelect('Contratos', { ID_Contrato: registro.ID_Contrato });
                    if (contrato && contrato.length > 0) {
                        const c = contrato[0];
                        const bruto = parseFloat(registro.Monto_Bruto_Estimado) || 0;

                        if (bruto > 0) {
                            // Surveillance (Vigilancia)
                            const pctVig = parseFloat(c.Retencion_Vigilancia_Pct) || 0.005; // Default 0.5% if not set
                            if (!registro.Deduccion_Surv_05_Monto || registro.Deduccion_Surv_05_Monto == 0) {
                                registro.Deduccion_Surv_05_Monto = (bruto * (pctVig / 100)).toFixed(2);
                            }

                            // You could add other retentions here as separate fields or deducciones records
                        }
                    }

                    // Auto-calculate Subtotal/Net if possible
                    const bruto = parseFloat(registro.Monto_Bruto_Estimado) || 0;
                    if (!registro.Subtotal && bruto > 0) {
                        const ded = parseFloat(registro.Deduccion_Surv_05_Monto) || 0;
                        registro.Subtotal = (bruto - ded).toFixed(2);
                    }
                }

                let match = null;
                const idOriginalDadoPorIA = registro[pkName];

                if (registro[pkName] && typeof registro[pkName] !== 'string' && typeof registro[pkName] !== 'number') {
                    delete registro[pkName]; // Limpiar basuras
                }

                if (registro[pkName]) {
                    const cond = {}; cond[pkName] = registro[pkName];
                    const res = dbSelect(tabla, cond);
                    if (res && res.length > 0) match = res[0];
                }

                // --- SMART MATCHING PARA CONCEPTOS (POR DESCRIPCIÓN) ---
                if (tabla === 'Catalogo_Conceptos' && !match && registro.ID_Contrato) {
                    const catalogoExistente = dbSelect('Catalogo_Conceptos', { ID_Contrato: registro.ID_Contrato });
                    const descNorm = normalizeText(registro.Descripcion || registro.Concepto || "");
                    if (descNorm) {
                        match = catalogoExistente.find(c => normalizeText(c.Descripcion) === descNorm);
                    }
                }

                // --- MATCH POR CLAVE (FALLBACK SI NO HAY MATCH POR DESCRIPCIÓN) ---
                if (tabla === 'Catalogo_Conceptos' && !match && registro.Clave && registro.ID_Contrato) {
                    const matchClave = dbSelect('Catalogo_Conceptos', { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato });
                    if (matchClave && matchClave.length > 0) match = matchClave[0];
                }

                // --- REGLAS DE NEGOCIO PARA MATRICES (TIPO NUMÉRICO) ---
                if (tabla === 'Matriz_Insumos') {
                    // Mapeo de Tipo de Insumo a Numérico: 1:Material, 2:ManoObra, 3:Equipo
                    const t = String(registro.Tipo_Insumo || registro.Tipo || "").toLowerCase();
                    if (t.includes('material')) registro.Tipo_Insumo = 1;
                    else if (t.includes('mano') || t.includes('jornal') || t.includes('cuadrilla') || t.includes('ayudante') || t.includes('especialista') || t.includes('gerente') || t.includes('tecnico')) registro.Tipo_Insumo = 2;
                    else if (t.includes('equipo') || t.includes('herra') || t.includes('maquin') || t.includes('vehiculo') || t.includes('computadora')) registro.Tipo_Insumo = 3;
                    else if (!isNaN(parseInt(t))) registro.Tipo_Insumo = parseInt(t);
                    else registro.Tipo_Insumo = 1; // Default Materials

                    // Asegurar que solo exista Tipo_Insumo según esquema
                    delete registro.Tipo;
                }

                if (!match && skName && registro[skName]) {
                    const cond = {}; cond[skName] = registro[skName];
                    const res = dbSelect(tabla, cond);
                    if (res && res.length > 0) {
                        match = res[0];
                        registro[pkName] = match[pkName];
                    }
                }

                // --- COMPOSITE UNIQUE CHECKS FOR MASS IMPORTERS ---
                if (!match) {
                    if (tabla === 'Catalogo_Conceptos' && registro.Clave && registro.ID_Contrato) {
                        const res = dbSelect(tabla, { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato });
                        if (res && res.length > 0) {
                            match = res[0];
                            registro[pkName] = match[pkName];
                        }
                    }
                    else if (tabla === 'Estimaciones' && registro.No_Estimacion && registro.ID_Contrato) {
                        const res = dbSelect(tabla, { No_Estimacion: registro.No_Estimacion, ID_Contrato: registro.ID_Contrato });
                        if (res && res.length > 0) {
                            match = res[0];
                            registro[pkName] = match[pkName];
                        }
                    }
                    else if (tabla === 'Detalle_Estimacion' && registro.ID_Estimacion && registro.ID_Concepto) {
                        const res = dbSelect(tabla, { ID_Estimacion: registro.ID_Estimacion, ID_Concepto: registro.ID_Concepto });
                        if (res && res.length > 0) {
                            match = res[0];
                            registro[pkName] = match[pkName];
                        }
                    }
                    else if (tabla === 'Matriz_Insumos' && registro.ID_Concepto && registro.Clave_Insumo) {
                        const res = dbSelect(tabla, { ID_Concepto: registro.ID_Concepto, Clave_Insumo: registro.Clave_Insumo });
                        if (res && res.length > 0) {
                            match = res[0];
                            registro[pkName] = match[pkName];
                        }
                    }
                }

                if (match) {
                    const cond = {}; cond[pkName] = match[pkName];
                    const dataMerged = { ...match };

                    // --- REGLA DE PERSISTENCIA (Priorizar DB sobre IA) ---
                    // Si es un concepto existente, evitamos que la IA sobrescriba campos base 
                    // a menos que estén vacíos en la base de datos.
                    const camposProtegidos = (tabla === 'Catalogo_Conceptos')
                        ? ['Clave', 'Descripcion', 'Unidad']
                        : [];

                    Object.keys(registro).forEach(k => {
                        if (registro[k] !== undefined && registro[k] !== null && registro[k] !== '') {
                            // Solo sobreescribir si el campo no está protegido o si el valor en DB es nulo/vacío
                            if (!camposProtegidos.includes(k) || !match[k]) {
                                dataMerged[k] = registro[k];
                            }
                        }
                    });

                    // En contexto de Matrices, forzar reglas incluso en actualización
                    if (tabla === 'Catalogo_Conceptos' && (datos.Matriz_Insumos || datos.Análisis_P_U)) {
                        dataMerged.Cantidad_Contratada = 1;
                        const pu = parseFloat(dataMerged.Precio_Unitario || 0);
                        dataMerged.Importe_Total_Sin_IVA = pu.toFixed(2);
                    }

                    dbUpdate(tabla, dataMerged, { [pkName]: match[pkName] });
                    resultados.actualizados++;

                    // Guardar mapa por si había un ID temporal
                    if (idOriginalDadoPorIA) mapaIds[idOriginalDadoPorIA] = match[pkName];

                } else {
                    // Al insertar conceptos nuevos vía Matrices, aplicar reglas de Orden y Cantidad
                    if (tabla === 'Catalogo_Conceptos') {
                        registro.Cantidad_Contratada = 1;
                        const pu = parseFloat(registro.Precio_Unitario || 0);
                        registro.Importe_Total_Sin_IVA = pu.toFixed(2);

                        // Generar Orden si no existe
                        if (!registro.Orden) {
                            const last = dbSelect('Catalogo_Conceptos', { ID_Contrato: registro.ID_Contrato });
                            registro.Orden = (last ? last.length : 0) + 1;
                        }
                    }

                    // Limpiar el PK temporal para que Code.gs asigne el consecutivo numérico real
                    if (idOriginalDadoPorIA) {
                        delete registro[pkName];
                    }

                    const resInsert = dbInsert(tabla, registro);
                    resultados.insertados++;

                    // Registrar el nuevo ID real para actualizar a los hijos que iterarán después
                    if (idOriginalDadoPorIA && resInsert.success && resInsert.insertId) {
                        mapaIds[idOriginalDadoPorIA] = resInsert.insertId;
                    }
                }
            });
        }
    });

    return {
        success: true,
        message: "Procesamiento IA completado con UPSERT Inteligente",
        detalles: resultados
    };
}

/**
 * Compara los datos extraídos por la IA con la base de datos actual para detectar cambios.
 */
function analizarCambiosIA(datos) {
    if (!datos) return generarRespuesta(false, "No hay datos para analizar.");

    const cambios = [];
    const tablasAAnalizar = ['Contratos', 'Contratistas', 'Convenios_Recurso', 'Catalogo_Conceptos', 'Matriz_Insumos'];

    tablasAAnalizar.forEach(tabla => {
        if (datos[tabla] && Array.isArray(datos[tabla])) {
            const headers = ESQUEMA_BD[tabla];
            const pkName = headers[0];

            let skName = null;
            if (tabla === 'Contratistas') skName = 'RFC';
            else if (tabla === 'Contratos') skName = 'Numero_Contrato';
            else if (tabla === 'Convenios_Recurso') skName = 'Numero_Acuerdo';

            datos[tabla].forEach(registro => {
                let match = null;
                // Buscar por PK o SK
                if (registro[pkName]) {
                    const res = dbSelect(tabla, { [pkName]: registro[pkName] });
                    if (res && res.length > 0) match = res[0];
                }
                if (!match && skName && registro[skName]) {
                    const res = dbSelect(tabla, { [skName]: registro[skName] });
                    if (res && res.length > 0) match = res[0];
                }

                // Caso especial Catalogo (por Clave + Contrato)
                if (!match && tabla === 'Catalogo_Conceptos' && registro.Clave && registro.ID_Contrato) {
                    const res = dbSelect(tabla, { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato });
                    if (res && res.length > 0) match = res[0];
                }

                // Caso especial Matriz (por ID_Concepto + Clave_Insumo)
                if (!match && tabla === 'Matriz_Insumos' && registro.ID_Concepto && registro.Clave_Insumo) {
                    const res = dbSelect(tabla, { ID_Concepto: registro.ID_Concepto, Clave_Insumo: registro.Clave_Insumo });
                    if (res && res.length > 0) match = res[0];
                }

                if (match) {
                    const camposProtegidos = (tabla === 'Catalogo_Conceptos')
                        ? ['Clave', 'Descripcion', 'Unidad']
                        : [];

                    headers.forEach(h => {
                        if (h.startsWith('ID_') || h === 'Link_Sharepoint') return;

                        // Si el campo está protegido y ya tiene dato en DB, ignorar sugerencia de la IA
                        if (camposProtegidos.includes(h) && (match[h] !== undefined && match[h] !== null && match[h] !== '')) {
                            return;
                        }

                        let valIA = registro[h];
                        let valDB = match[h];

                        if (valIA === undefined || valIA === null) return;

                        let nIA = utilNormalizarValor(valIA);
                        let nDB = utilNormalizarValor(valDB);

                        if (nIA !== "" && nIA !== nDB) {
                            cambios.push({
                                tabla: tabla,
                                numContrato: match['Numero_Contrato'] || match['Numero_Acuerdo'] || match['Clave'] || match['Clave_Insumo'] || 'N/A',
                                campo: h,
                                valorActual: (valDB instanceof Date) ? Utilities.formatDate(valDB, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(valDB || "(Vacío)"),
                                valorIA: String(valIA),
                                rawIA: nIA,
                                rawDB: nDB
                            });
                        }
                    });
                }
            });
        }
    });

    // --- VALIDACIÓN CRUZADA (CROSS-CHECK) PARA MATRICES ---
    if (datos.Catalogo_Conceptos && (datos.Matriz_Insumos || datos.Análisis_P_U)) {
        const insumos = datos.Matriz_Insumos || datos.Análisis_P_U;
        datos.Catalogo_Conceptos.forEach(concepto => {
            const tempIdMapping = concepto.ID_Concepto || concepto.Clave;
            const relacionados = insumos.filter(i => i.ID_Concepto === tempIdMapping);

            if (relacionados.length > 0) {
                const suma = relacionados.reduce((acc, i) => acc + (parseFloat(i.Importe) || 0), 0);
                const cdIA = parseFloat(concepto.Costo_Directo) || 0;

                if (Math.abs(suma - cdIA) > 0.05) { // Tolerancia un poco mayor por redondeos acumulados
                    cambios.push({
                        tipo: "VALIDACION",
                        tabla: "Catalogo_Conceptos",
                        numContrato: concepto.Clave || "N/A",
                        campo: "Costo_Directo (Validación)",
                        valorActual: "Suma Insumos: " + suma.toFixed(2),
                        valorIA: "CD Extraído: " + cdIA.toFixed(2),
                        nota: "La suma de los insumos no coincide con el Costo Directo extraído."
                    });
                }
            }
        });
    }

    return generarRespuesta(true, cambios);
}

/**
 * Normaliza valores para una comparación lógica justa (fechas, números, vacíos).
 */
function utilNormalizarValor(val) {
    if (val === null || val === undefined || val === "") return "";

    // 1. Manejo de Fechas
    if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // 2. Manejo Numérico (Limpieza de símbolos y redondeo)
    if (typeof val === 'number') {
        return Math.round(val * 100) / 100;
    }

    if (typeof val === 'string') {
        // Si parece un número/moneda (ej: "$1,234.56")
        if (/^[\$\s\d,\.-]+$/.test(val) && /[0-9]/.test(val)) {
            let n = parseFloat(val.replace(/[\$\s,]/g, ''));
            if (!isNaN(n)) return Math.round(n * 100) / 100;
        }

        // Si es un string con formato ISO o fecha simple
        if (/^\d{4}-\d{2}-\d{2}/.test(val)) {
            return val.substring(0, 10);
        }

        return normalizeText(val); // Normalización de texto base
    }

    return normalizeText(String(val));
}
