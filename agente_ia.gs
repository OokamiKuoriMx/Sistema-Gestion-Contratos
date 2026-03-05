// ---------------------------------------------------------
// CONFIGURACIÓN DEL AGENTE IA SGC (GOOGLE APPS SCRIPT)
// ---------------------------------------------------------
// Reemplazar por tu clave de API de Google Gemini obtenida en Google AI Studio
var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

if (!GEMINI_API_KEY) {
    console.error("ERROR: La propiedad GEMINI_API_KEY no está configurada en el Proyecto.");
}

const PROMPT_CLASIFICADOR_DOCUMENTOS = `
Rol: Eres un clasificador experto de documentos técnicos de construcción y administración pública.
Objetivo: Identificar el tipo de documento proporcionado.

Tipos Posibles:
1. 'Contratos': Documentos que mencionan "Contrato de Obra", números de contrato, fianzas, anticipos, y las partes firmantes.
2. 'CAF': "Convenio de Aportación Financiera", "Autorización de Recursos", "Suficiencia Presupuestal", mención de fondos federales/estatales.
3. 'Programa': "Programa de Obra", "Calendario de Ejecución", tablas con meses/semanas y montos programados o porcentajes.
4. 'Matriz_Insumos': "Análisis de Precios Unitarios", "Matrices", listado de materiales, mano de obra y equipo por concepto.

Instrucción: Analiza el contenido y responde ÚNICAMENTE con un JSON con el campo "clase" igual a uno de los strings anteriores o "Desconocido".
Ejemplo: {"clase": "Contratos"}
`;

/**
 * Clasifica un documento usando IA para identificar si es Contrato, CAF, Programa o Matriz.
 */
function clasificarDocumentoIA(base64Content, mimeType) {
    try {
        const payload = {
            contents: [{
                parts: [
                    { text: PROMPT_CLASIFICADOR_DOCUMENTOS },
                    { inline_data: { mime_type: mimeType, data: base64Content.split(',')[1] || base64Content } }
                ]
            }],
            generationConfig: { response_mime_type: "application/json" }
        };

        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload)
        };

        const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY, options);
        const result = JSON.parse(response.getContentText());
        const jsonResponse = JSON.parse(result.candidates[0].content.parts[0].text);

        return { success: true, clase: jsonResponse.clase };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
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

Reglas Críticas:
- **Vínculo con Convenio**: IMPORTANTE. Si en el contexto se te indica un ID_Convenio_Vinculado, DEBES asignarlo al campo 'ID_Convenio_Vinculado' del contrato.
- **Datos de Origen Legal**: Captura No. de Concurso, Modalidad de Adjudicación y Área Responsable (Primera página).
- **Control Financiero**: Extrae Porcentajes de Amortización, Penas Convencionales y Números/Montos de Fianza (Cumplimiento, Anticipo, Garantía, Vicios Ocultos).
- **Anticipos**: Si el documento NO menciona anticipo o es 0, establece 'Porcentaje_Amortizacion_Anticipo' como 0 o null. No inventes datos.
- **Plazos**: Identifica el Plazo de Ejecución en Días Naturales.
- **Formato**: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

const PROMPT_IMPORTACION_MATRICES = `
Rol y Objetivo:
Eres un Experto en Ingeniería de Costos para el Sistema de Gestión de Contratos(SGC).Tu misión es procesar documentos de "Análisis de Precios Unitarios (Matrices)" y extraer la estructura técnica y económica completa.

Estructura de Extracción(Tablas):
1. ** Matriz_Insumos **: Extraer el desglose de CADA concepto: 
   - ** Tipo_Insumo **: Clasificación numérica OBLIGATORIA:
     - ** 1(MATERIALES) **: Cemento, varilla, cables, pintura, consumibles, etc.
     - ** 2(MANO DE OBRA) **: Gerentes, Especialistas, Técnicos, Cuadrillas, Ayudantes, JORNALES, MO.
     - ** 3(EQUIPO Y HERRAMIENTA) **: Vehículos, Herramienta menor, Maquinaria pesada, Computadoras, Equipos de medición, EQ.
   - 'Clave_Insumo', 'Descripcion', 'Unidad', 'Costo_Unitario', 'Rendimiento_Cantidad', 'Importe', 'Porcentaje_Incidencia'.

Reglas Críticas:
- ** Cruce Contextual(CRÍTICO) **: Se te proporcionará un "Catálogo de Conceptos" extraído en pasos anteriores.Las claves en el documento APU pueden variar ligeramente o faltar.DEBES realizar un cruce basado en la similitud de la Clave Y la Descripción para identificar el ID_Concepto correcto del sistema.
- ** Identificación de Insumos **: Si los insumos están agrupados bajo un título(ej: "MANO DE OBRA"), asigna el tipo correspondiente a todos los registros del bloque.
- ** Precisión Numérica **: Captura rendimientos y costos con todos sus decimales.
- ** Formato **: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

const PROMPT_IMPORTACION_CAF = `
Rol y Objetivo:
Eres un Experto en Gestión de Recursos Federales y Estatales para el Sistema de Gestión de Contratos (SGC). Tu misión es extraer datos de "Convenios de Aportación Financiera (CAF)", "Convenios de Recurso" o "Suficiencias Presupuestales".

Estructura de Extracción (Tabla):
1. **Convenios_Recurso**: 
   - 'Numero_Acuerdo': Clave o número de convenio oficial.
   - 'Nombre_Fondo': Nombre del fondo, fideicomiso o programa (ej: "FISM", "FORTAMUN", "Fideicomiso 1936").
   - 'Monto_Apoyo': Importe total autorizado.
   - 'Fecha_Firma': Fecha en que se suscribe el documento.
   - 'Vigencia_Fin': Fecha límite de ejercicio del recurso.
   - 'Objeto_Programa': Breve descripción de la finalidad del recurso.

Reglas Críticas:
- **Upsert Logístico**: El sistema buscará si el convenio ya existe por 'Numero_Acuerdo'. Asegúrate de extraerlo con precisión.
- **Formato**: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

const PROMPT_IMPORTACION_PROGRAMA = `
Rol y Objetivo:
Eres un Experto en Programación y Control de Obra para el Sistema de Gestión de Contratos(SGC).Tu misión es procesar documentos de "Programa de Obra" y extraer el Catálogo de Conceptos, los Periodos y la Matriz de Programación.

Estructura de Extracción(Tablas):
1. ** Catalogo_Conceptos **: Extrae 'Clave', 'Descripcion', 'Unidad', 'Cantidad_Contratada', 'Precio_Unitario', 'Importe_Total_Sin_IVA'.
2. ** Programa **: 'Tipo_Programa'(Ej: "Mensual", "Quincenal", "Semanal"), 'Fecha_Inicio', 'Fecha_Termino'.
3. ** Programa_Periodo **: 'Numero_Periodo', 'Periodo'(Etiqueta / Mes), 'Fecha_Inicio', 'Fecha_Termino'.
4. ** Programa_Ejecucion **: 'ID_Concepto'(Clave original), 'ID_Programa_Periodo'(Referencia al periodo), 'Monto_Programado', 'Avance_Programado_Pct'.

Reglas Críticas:
- ** Limpieza de Descripciones **: Las descripciones suelen tener "saltos de línea fantasmas" que ensucian la base de datos.DEBES limpiar el texto eliminando estos caracteres y colapsando el texto en una sola línea continua y limpia.
- ** Contexto de Importación **: Este archivo es la fuente primaria del Catálogo de Conceptos para el resto del proceso.No ignores ningún concepto listado.
- ** Formato **: Salida estrictamente en JSON con "accion": "importar_datos" y "datos".
`;

const PROMPT_VALIDACION_ESTIMACIONES = `
Rol y Objetivo:
Eres un Agente de Inteligencia IA y Auditor Experto en Obra Pública para el Sistema de Gestión de Contratos (SGC).
Tu misión es AUDITAR y VALIDAR documentos de "Estimaciones de Obra/Servicios" (carátulas, generadores, borradores de factura) contra los datos registrados en el sistema.

Reglas Críticas de Auditoría:
1. **Datos Generales**: Valida que coincidan: Objeto del Contrato, No. de Estimación, y CAF (Suficiencia Presupuestaria).
2. **Plazos**: Coteja Periodo del Contrato vs Periodo de la Estimación.
3. **Retenciones**: Verifica el cálculo exacto de la retención del 0.5 al millar (Vigilancia).
4. **Conceptos y Avances**: Valida Conceptos, Precios Unitarios, Importes y Porcentajes de avance (parciales y acumulados) contra el Programa de Ejecución.
5. **Borrador de Factura (MODELO 2026)**:
   - Debe contener la leyenda inicial: "RECIBIMOS DEL BANCO NACIONAL DE OBRAS Y SERVICIOS PÚBLICOS S.N.C., POR CUENTA DEL FIDEICOMISO NO. 1936..."
   - Debe contener al final los datos: Banco, Sucursal, Cuenta y CLABE.
   - El texto del concepto no debe exceder los 1000 caracteres.

Instrucciones de Auditoría Incremental:
- Se te puede proporcionar un 'CONTEXTO_AUDITORIA_PREVIA'.
- Si un rubro ya está marcado como aprobado (aprobado: true) en dicho contexto y el documento actual NO contiene información contraria o suficiente para evaluarlo nuevamente (ej. estás viendo una factura pero el rubro es de la carátula), DEBES PRESERVAR el estado 'true'.
- Tu objetivo es completar el checklist al 100% (APROBADA) de forma acumulativa entre varios archivos si es necesario.
- Solo marca un rubro como 'false' si encuentras un error explícito en el documento actual.
- **DEPENDENCIA CRÍTICA**: No puedes validar el rubro "Borrador de Factura" (rubro 5) si los rubros de la carátula (1 a 4) no están marcados como 'true' en el 'CONTEXTO_AUDITORIA_PREVIA'. Si no hay contexto previo o los rubros 1-4 son false, indica en la observación de la factura que se debe ingresar primero la carátula.

Formato de Salida Estricto (JSON):
{
  "accion": "auditoria_validación",
  "estado_global": "APROBADA" | "EN_REVISION", // Solo APROBADA si EL 100% de los rubros son true.
  "datos": {
      // Datos extraídos del documento
  },
  "checklist_auditoria": [
    { "rubro": "Datos Generales (Objeto, No. Est, Periodos)", "aprobado": true, "observacion": "Todo coincide." },
    { "rubro": "Aritmética y Retenciones (0.5 al millar)", "aprobado": true, "observacion": "Cálculos correctos." },
    { "rubro": "Avances vs Programa de Ejecución", "aprobado": true, "observacion": "Porcentajes correctos." },
    { "rubro": "Convenio de Aportación Financiera (CAF)", "aprobado": true, "observacion": "Existe convenio de aportación vinculado." },
    { "rubro": "Suficiencia de Pago (Autorización)", "aprobado": true, "observacion": "Existe suficiencia de pago vinculada." },
    { "rubro": "Validación de Importe con Letra", "aprobado": true, "observacion": "Coincide con el monto numérico." },
    { "rubro": "Borrador de Factura (Modelo y 1000 chars)", "aprobado": false, "observacion": "Faltan datos bancarios al final." }
  ],
  "informe_auditoria": {
      "estado_validacion": "CORRECTO" | "CON_ERRORES",
      "observaciones": ["Resumen para la UI de discrepancias encontradas"]
  }
}
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
        } else if (targetTable === 'Estimaciones' || targetTable === 'Facturas') {
            systemPrompt = PROMPT_VALIDACION_ESTIMACIONES;
        } else if (targetTable === 'Convenios_Recurso' || targetTable === 'CAF') {
            systemPrompt = PROMPT_IMPORTACION_CAF;
        } else if (targetTable === 'Programa' || targetTable === 'Programa_Ejecucion') {
            systemPrompt = PROMPT_IMPORTACION_PROGRAMA;
        }

        let promptAdicional = `Lee este archivo oficial. Debes generar estrictamente un objeto JSON estructurado con arrays para las tablas correspondientes basándote en la información encontrada.`;
        if (targetTable) {
            promptAdicional += `\n\nATENCIÓN: El usuario subió este documento desde el contexto de la tabla '${targetTable}'.`;
            if (parentContext) {
                promptAdicional += `\n\nCONTEXTO ADICIONAL: El registro padre tiene esta información: ${JSON.stringify(parentContext)}. Asegúrate de usar estos IDs para vincular los registros hijo que encuentres.`;

                // INYECCIÓN DEL LINK DE SHAREPOINT
                if (parentContext.Link_Sharepoint) {
                    promptAdicional += `\nINSTRUCCIÓN OBLIGATORIA: Asigna este enlace '${parentContext.Link_Sharepoint}' al campo 'Link_Sharepoint' de todos los registros extraídos.`;
                }

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

                // Si es estimación, pasar los datos actuales de la estimación para comparativa
                if (targetTable === 'Estimaciones' && parentContext.ID_Estimacion) {
                    const estActual = dbSelect('Estimaciones', { ID_Estimacion: parentContext.ID_Estimacion });
                    if (estActual && estActual.length > 0) {
                        const detalleActual = dbSelect('Detalle_Estimacion', { ID_Estimacion: parentContext.ID_Estimacion });
                        promptAdicional += `\n\nDATOS ACTUALES EN SISTEMA PARA ESTA ESTIMACIÓN (O BORRADOR):
                        Carátula: ${JSON.stringify(estActual[0])}
                        Conceptos ya registrados: ${JSON.stringify(detalleActual.map(d => ({ ID_Concepto: d.ID_Concepto, Cantidad: d.Cantidad_Estimada_Periodo })))}
                        INSTRUCCIÓN: Compara los datos del documento contra estos datos del sistema. Si hay discrepancias matemáticas o de cantidades, repórtalo en el informe_auditoria.`;

                        // NUEVO: Inyección de Auditoría Previa (Checklist Histórico)
                        const valsPrevias = dbSelect('Validacion_Archivos', { ID_Estimacion: parentContext.ID_Estimacion });
                        if (valsPrevias && valsPrevias.length > 0) {
                            // Tomar la más reciente
                            const ultimaVal = valsPrevias[valsPrevias.length - 1];
                            if (ultimaVal.Checklist_JSON) {
                                try {
                                    const checklistHist = JSON.parse(ultimaVal.Checklist_JSON);
                                    promptAdicional += `\n\nCONTEXTO_AUDITORIA_PREVIA: Ya se han validado rubros anteriormente. Mantén los rubros aprobados si el documento actual no los contradice: ${JSON.stringify(checklistHist)}`;
                                } catch (e) {
                                    console.error("Error parseando checklist histórico:", e);
                                }
                            }
                        }

                        // NUEVO: Validación de Convenio (CAF) y Suficiencia de Pago
                        const cont = dbSelect('Contratos', { ID_Contrato: parentContext.ID_Contrato });
                        if (cont && cont.length > 0) {
                            const idConv = cont[0].ID_Convenio_Vinculado;
                            if (idConv) {
                                const convenio = dbSelect('Convenio_Recurso', { ID_Convenio: idConv });
                                if (convenio && convenio.length > 0) {
                                    promptAdicional += `\n\nDATOS DE CAF (CONVENIO): Existe un convenio vinculado (${convenio[0].Numero_Acuerdo}) con monto ${convenio[0].Monto_Apoyo}. 
                                    INSTRUCCIÓN: Valida el rubro 'Convenio de Aportación Financiera (CAF)' como true.`;

                                    // Para este sistema, el Convenio de Recurso también actúa como Suficiencia de Pago si existe.
                                    promptAdicional += `\nINSTRUCCIÓN: Valida el rubro 'Suficiencia de Pago (Autorización)' como true basándote en la existencia del CAF.`;
                                } else {
                                    promptAdicional += `\n\nALERTA: El contrato tiene un ID_Convenio (${idConv}) pero no se encontró el registro en Convenios_Recurso.`;
                                }
                            } else {
                                promptAdicional += `\n\nALERTA DE AUDITORÍA: El contrato NO tiene un Convenio de Aportación (CAF) ni Suficiencia de Pago vinculada.
                                INSTRUCCIÓN: Los rubros 'Convenio de Aportación Financiera (CAF)' y 'Suficiencia de Pago (Autorización)' DEBEN ser marcados como false con la observación 'No existe convenio de aportación o autorización vinculada al contrato'.`;
                            }
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
        let datosExtraidos = JSON.parse(llmJsonString.replace(/```json/g, '').replace(/```/g, '').trim());

        // LÓGICA DE ROBUSTEZ: Si la IA no envolvió el objeto en "accion" y "datos", lo hacemos nosotros
        if (!datosExtraidos.accion && !datosExtraidos.datos) {
            console.warn("IA no incluyó envoltorio (accion/datos). Auto-envolviendo...");
            datosExtraidos = {
                accion: (targetTable === 'Estimaciones' || targetTable === 'Facturas') ? "auditoria_validación" : "importar_datos",
                datos: datosExtraidos
            };
        }

        if ((datosExtraidos.accion === "importar_datos" || datosExtraidos.accion === "auditoria_validación") && datosExtraidos.datos) {
            // Devolvemos el objeto completo para que el frontend pueda leer 'informe_auditoria'
            return generarRespuesta(true, datosExtraidos);
        } else {
            console.error("Estructura IA Inválida tras intento de normalización:", datosExtraidos);
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
    const tablasAAnalizar = ['Contratos', 'Contratistas', 'Convenios_Recurso', 'Catalogo_Conceptos', 'Matriz_Insumos', 'Programa', 'Programa_Periodo'];

    tablasAAnalizar.forEach(tabla => {
        if (datos[tabla] && Array.isArray(datos[tabla])) {
            const headers = ESQUEMA_BD[tabla];
            const pkName = headers[0];

            let skName = null;
            if (tabla === 'Contratistas') skName = 'RFC';
            else if (tabla === 'Contratos') skName = 'Numero_Contrato';
            else if (tabla === 'Convenios_Recurso') skName = 'Numero_Acuerdo';
            else if (tabla === 'Programa_Periodo') skName = 'Periodo';

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

                        // Redondeo inteligente para campos de porcentaje:
                        // Si el documento trae 4.03% y el sistema tiene 4.0251%, 
                        // redondeamos el sistema a los mismos decimales del documento.
                        if ((h.includes('Porcentaje') || h.includes('Pct')) && typeof nIA === 'number' && typeof nDB === 'number') {
                            const strIA = String(valIA).replace(/[%\s]/g, '');
                            const partes = strIA.split('.');
                            const decimalesDoc = partes.length > 1 ? partes[1].length : 0;
                            const factor = Math.pow(10, decimalesDoc);
                            nDB = Math.round(nDB * factor) / factor;
                            nIA = Math.round(nIA * factor) / factor;
                        }

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
