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
Objetivo: Identificar el tipo de documento proporcionado aplicando la "Regla de Oro".

CONTEXTO ADICIONAL:
- Filename: {{FILENAME}}

REGLA DE ORO (Criterios de Clasificación):

1. **CAF (Convenio de Apoyo Financiero)**: 
   - Palabras clave: "BANOBRAS", "FONADIN", "FIDUCIARIO", "Secretaría de Hacienda", "Suficiencia Presupuestal". 
   - Contenido: Oficio legal corto (1-5 págs) sobre asignación de recursos. No tiene listados de materiales ni calendarios.

2. **APU / MATRIZ (Análisis de Precios Unitarios)**: 
   - Palabras clave: "Costo Directo", "Indirectos", "Utilidad", "Relación de Insumos". 
   - **ESTRUCTURA (CRÍTICO)**: Contiene tablas extensas de desgloses de INSUNMOS (MANO DE OBRA, EQUIPO, MATERIALES) con columnas de Unidad (JOR, Hr, kg), Rendimientos y Costos Unitarios.
   - **NOTA**: Si el documento está lleno de tablas de desgloses técnicos de precios, su clase es 'Matriz_Insumos', incluso si tiene un encabezado de un contrato.

3. **CONTRATO**: 
   - Palabras clave: "TIPO DE CONTRATO", "SERVICIOS RELACIONADOS CON LA OBRA PÚBLICA", "FECHA DE ADJUDICACIÓN", "RFC", "Fianzas", "Cláusulas".
   - **ESTRUCTURA (CRÍTICO)**: Es un documento legal articulado en CLÁUSULAS (Primera, Segunda...). No contiene desgloses de materiales por concepto.

4. **PROGRAMA (Programa de Erogaciones / Trabajo)**: 
   - Palabras clave: "PROGRAMA DE EROGACIONES", "EJECUCIÓN GENERAL", "CALENDARIO". 
   - ESTRUCTURA: Matriz temporal con columnas para meses/quincenas/semanas y celdas con barras, porcentajes (%) o montos programados.

INSTRUCCIÓN: Analiza el contenido y responde ÚNICAMENTE con un JSON con el campo "clase".
Clases permitidas: 'Contratos', 'CAF', 'Programa', 'Matriz_Insumos' o 'Desconocido'.
Ejemplo: {"clase": "Contratos"}
`;

/**
 * Clasifica un documento usando IA para identificar si es Contrato, CAF, Programa o Matriz.
 */
function clasificarDocumentoIA(base64Content, mimeType, filename = "documento.pdf") {
    try {
        const finalPrompt = PROMPT_CLASIFICADOR_DOCUMENTOS.replace('{{FILENAME}}', filename);
        const payload = {
            contents: [{
                parts: [
                    { text: finalPrompt },
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

        const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + GEMINI_API_KEY, options);
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
Eres un Agente de Inteligencia IA experto en el Sistema de Gestión de Contratos (SGC). Tu misión es procesar documentos de CONTRATO.

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido (sin formato Markdown, sin texto adicional).

Estructura requerida:
{
  "tipo_documento": "CONTRATO",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "Numero_Contrato": "Clave del contrato (ej. 25-E-CF-...)",
    "Objeto_Contrato": "Descripción del servicio u obra",
    "Fecha_Firma": "Fecha de formalización (YYYY-MM-DD)",
    "Fecha_Inicio_Obra": "Fecha de inicio pactada (YYYY-MM-DD)",
    "Fecha_Fin_Obra": "Fecha de término pactada (YYYY-MM-DD)",
    "Montos": {
      "Monto_Total_Sin_IVA": 1499787.55,
      "Monto_Total_Con_IVA": 1739753.56
    },
    "Contratista": {
      "Razon_Social": "Nombre de la empresa",
      "RFC": "RFC con homoclave",
      "Representante_Legal": "Nombre del representante"
    },
    "ID_Convenio_Vinculado": "ASIGNAR EL VALOR PROPORCIONADO EN EL CONTEXTO SI EXISTE"
  }
}
`;

const PROMPT_IMPORTACION_MATRICES = `
Rol y Objetivo:
Eres un Experto en Ingeniería de Costos. Tu misión es procesar documentos de APU (Analísis de Precios Unitarios).
UN DOCUMENTO PUEDE CONTENER MÚLTIPLES CONCEPTOS. Extrae TODOS los conceptos que encuentres.

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "APU",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "conceptos": [
      {
        "Clave_Concepto": "Clave principal del análisis (ej. N3-25-1)",
        "Descripcion_Concepto": "Descripción de la actividad principal",
        "Unidad": "Unidad del concepto (ej. ESTUDIO, MES)",
        "Costo_Directo": 200000.00,
        "Precio_Unitario_Total": 255011.33,
        "insumos": [
          {
            "Tipo_Insumo": "1(MATERIALES) | 2(MANO DE OBRA) | 3(EQUIPO)",
            "Clave_Insumo": "Clave del insumo (ej. GERPROY, EQ01)",
            "Descripcion": "Descripción del insumo",
            "Unidad": "JOR, Hr, Pza, etc.",
            "Costo_Unitario": 7032.75,
            "Rendimiento_Cantidad": 5.0,
            "Importe": 35163.75,
            "Porcentaje_Incidencia": 31.27
          }
        ]
      }
    ],
    "porcentajes_contrato": {
      "Pct_Indirectos_Oficina": 27.62,
      "Pct_Indirectos_Campo": 0.45,
      "Pct_Indirectos_Totales": 28.07,
      "Pct_Financiamiento": 1.60,
      "Pct_Utilidad": 8.00,
      "Pct_Cargos_SFP": 0.50,
      "Pct_ISN": 0.00
    }
  }
}

NOTAS CRÍTICAS:
- Extrae TODOS los conceptos del documento, no solo el primero.
- El "Costo_Directo" de cada concepto es la SUMA de todos los importes de sus insumos (materiales + mano de obra + equipo). NO es el Precio Unitario Total.
- El "Precio_Unitario_Total" ya incluye el Costo Directo + Indirectos + Utilidad + Financiamiento.
- Cada concepto tiene su propio array de "insumos" asociados.
- Los porcentajes de Indirectos, Financiamiento, Utilidad, Cargos SFP e ISN generalmente aparecen en la hoja resumen del análisis, en una tabla de desglose porcentual. Extráelos como decimales (ej. 27.62, no 0.2762).
- Si algún porcentaje no aparece en el documento, usar 0.
`;

const PROMPT_IMPORTACION_CAF = `
Rol y Objetivo:
Eres un Experto en Gestión de Recursos. Tu misión es extraer datos de CAF (Convenio de Apoyo Financiero).

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "CAF",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "Numero_Acuerdo": "Extraer el folio o número del convenio",
    "Nombre_Fondo": "Nombre del fideicomiso o fondo (ej. FONADIN)",
    "Monto_Apoyo": 0.0,
    "Fecha_Firma": "YYYY-MM-DD",
    "Entidades_Involucradas": ["BANOBRAS", "SHCP", "..."]
  }
}
`;

const PROMPT_IMPORTACION_PROGRAMA = `
Rol y Objetivo:
Eres un Experto en Programación de Obra. Tu misión es extraer datos del PROGRAMA de Erogaciones.

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "PROGRAMA",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "Fechas_Generales": {
      "Fecha_Inicio": "YYYY-MM-DD",
      "Fecha_Fin": "YYYY-MM-DD",
      "Plazo_Dias": 90
    },
    "Periodos_Identificados": [
      {
        "Nombre": "DICIEMBRE 2025",
        "Fecha_Inicio": "YYYY-MM-DD",
        "Fecha_Termino": "YYYY-MM-DD"
      },
      {
        "Nombre": "ENERO 2026",
        "Fecha_Inicio": "YYYY-MM-DD",
        "Fecha_Termino": "YYYY-MM-DD"
      }
    ],
    "conceptos_programados": [
      {
        "Clave_Concepto": "Clave que se está programando (ej. N3-25-5)",
        "Descripcion_Limpiada": "Descripción sin saltos de línea",
        "Unidad": "Unidad original",
        "Cantidad": 1.00,
        "Precio_Unitario": 138975.44,
        "Importe_Total": 138975.44,
        "avances_por_periodo": [
          {
            "Nombre_Periodo": "DICIEMBRE 2025",
            "Porcentaje_Avance": 10.06,
            "Monto_Periodo": 13980.92
          }
        ]
      }
    ]
  }
}

REGLAS CRÍTICAS DE LECTURA MATRICIAL (¡NO LAS IGNORES!):
1. Identifica los Encabezados (Columnas): Primero, busca los meses o quincenas que forman las columnas (ej. DICIEMBRE, ENERO, FEBRERO). Estos son tus periodos.
2. Lectura de Intersecciones (Filas vs Columnas): Lee cada fila de concepto. El documento distribuye los avances (%) y montos ($) a lo largo de las columnas de los meses.
3. PROHIBICIÓN DE AGRUPAMIENTO: NO asocies todos los importes o porcentajes al último periodo, ni los asumas todos como el "Total". Debes distribuir los valores exactamente en el mes al que pertenecen según el espaciado del documento.
4. SOLO DELTAS: Solo incluye en avances_por_periodo los periodos donde el concepto tiene actividad real (Porcentaje > 0 o Monto > 0). NO incluyas periodos con valor 0.

NOTAS SOBRE FECHAS DE PERIODOS:
- Para cada periodo, extrae el rango de fechas EXACTO que aparece en el documento.
- Si el documento solo muestra meses (ej. "DICIEMBRE 2025"), usa el primer día del mes como Fecha_Inicio y el último día como Fecha_Termino.
- Si el primer periodo coincide con Fecha_Inicio del contrato, usa esa fecha exacta como inicio del periodo.
- Si el último periodo coincide con Fecha_Fin del contrato, usa esa fecha exacta como fin del periodo.
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
        if (parentContext && typeof parentContext === 'string') {
            try {
                parentContext = JSON.parse(parentContext);
            } catch (e) {
                console.error("Error parseando parentContext:", e);
            }
        }

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
        let extraction = JSON.parse(llmJsonString.replace(/```json/g, '').replace(/```/g, '').trim());

        // LÓGICA REGLA DE ORO: Extraer datos_extraidos si existe el wrapper
        let datosFinales = extraction.datos_extraidos || extraction.datos || extraction;

        // Formatear para el orquestador frontend
        const respuestaFormateada = {
            accion: (targetTable === 'Estimaciones' || targetTable === 'Facturas') ? "auditoria_validación" : "importar_datos",
            datos: datosFinales,
            tipo_documento: extraction.tipo_documento || "DESCONOCIDO", // De la Regla de Oro
            nivel_confianza: extraction.nivel_confianza || "Bajo",
            parentContext: parentContext // Preservar para el guardado
        };

        // Si es auditoría, preservar el checklist
        if (extraction.checklist_auditoria) {
            respuestaFormateada.checklist_auditoria = extraction.checklist_auditoria;
            respuestaFormateada.informe_auditoria = extraction.informe_auditoria;
            respuestaFormateada.estado_global = extraction.estado_global;
        }

        return generarRespuesta(true, respuestaFormateada);

    } catch (e) {
        console.error(e);
        return generarRespuesta(false, e.toString(), 'processDocumentWithAI');
    }
}

/**
 * Función auxiliar para insertar y actualizar los datos en las pestañas como hace el webhook.
 * Realiza un chequeo de existencia (UPSERT) para evitar records duplicados desde la IA.
 * @param {Object} respuestaIA - El objeto devuelto por processDocumentWithAI.
 * @param {string} tablaDestino - La tabla objetivo (opcional).
 * @param {string} idConvenioVinculado - ID de convenio para vincular (opcional).
 */
function guardarDatosIA(respuestaIA, tablaDestino = null, idConvenioVinculado = null) {
    if (!respuestaIA) return generarRespuesta(false, "No se recibieron datos para guardar.");

    // El objeto puede venir directo o envuelto por processDocumentWithAI
    const datos = respuestaIA.datos || respuestaIA;
    const tipo = respuestaIA.tipo_documento || "DESCONOCIDO";
    let resultados = { insertados: 0, actualizados: 0, ids: {} };

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

    // --- NORMALIZACIÓN REGLA DE ORO ---
    // Si los datos vienen en el formato anidado de la Regla de Oro, los "aplanamos" o movemos a las llaves que espera guardarDatosIA
    if (tipo === 'CONTRATO') {
        const c = datos; // datos ya es datos_extraidos
        const contratista = {
            Razon_Social: c.Contratista?.Razon_Social,
            RFC: c.Contratista?.RFC,
            Representante_Legal: c.Contratista?.Representante_Legal
        };
        const contrato = {
            Numero_Contrato: c.Numero_Contrato,
            Objeto_Contrato: c.Objeto_Contrato,
            Fecha_Firma: c.Fecha_Firma,
            Fecha_Inicio_Obra: c.Fecha_Inicio_Obra,
            Fecha_Fin_Obra: c.Fecha_Fin_Obra,
            Monto_Total_Sin_IVA: c.Montos?.Monto_Total_Sin_IVA,
            Monto_Total_Con_IVA: c.Montos?.Monto_Total_Con_IVA,
            ID_Convenio_Vinculado: idConvenioVinculado || c.ID_Convenio_Vinculado
        };
        datos.Contratistas = [contratista];
        datos.Contratos = [contrato];
    } else if (tipo === 'CAF') {
        datos.Convenios_Recurso = [datos];
    } else if (tipo === 'PROGRAMA') {
        // Mapear Programa
        datos.Programa = [{
            Programa: datos.Nombre_Programa || datos.Programa || "Programa General de Obra",
            Tipo_Programa: datos.Tipo_Programa || "MENSUAL",
            Fecha_Inicio: datos.Fechas_Generales?.Fecha_Inicio || null,
            Fecha_Termino: datos.Fechas_Generales?.Fecha_Fin || null,
            Plazo_Ejecucion_Dias: datos.Fechas_Generales?.Plazo_Dias || null
        }];
        // Períodos (ahora vienen como objetos con Nombre, Fecha_Inicio, Fecha_Termino)
        if (datos.Periodos_Identificados) {
            datos.Programa_Periodo = datos.Periodos_Identificados.map((p, i) => {
                // Soportar formato viejo (string) y nuevo (objeto)
                const esObjeto = typeof p === 'object';
                return {
                    Numero_Periodo: i + 1,
                    Periodo: esObjeto ? p.Nombre : p,
                    Fecha_Inicio: esObjeto ? p.Fecha_Inicio : null,
                    Fecha_Termino: esObjeto ? p.Fecha_Termino : null
                };
            });
        }
        // Conceptos y Ejecución
        if (datos.conceptos_programados) {
            datos.Catalogo_Conceptos = datos.conceptos_programados.map(cp => ({
                Clave: cp.Clave_Concepto,
                Descripcion: cp.Descripcion_Limpiada,
                Unidad: cp.Unidad,
                Cantidad_Contratada: cp.Cantidad || 1,
                Precio_Unitario: cp.Precio_Unitario || null,
                Importe_Total_Sin_IVA: cp.Importe_Total
            }));
            // La ejecución requiere IDs de período, se manejará en el loop de jerarquía
            datos.Programa_Ejecucion = [];
            datos.conceptos_programados.forEach(cp => {
                cp.avances_por_periodo?.forEach(av => {
                    // SOLO DELTAS: no registrar periodos sin actividad
                    const monto = parseFloat(av.Monto_Periodo) || 0;
                    const pct = parseFloat(av.Porcentaje_Avance) || 0;
                    if (monto === 0 && pct === 0) return;

                    datos.Programa_Ejecucion.push({
                        Clave_Concepto_Temp: cp.Clave_Concepto,
                        Periodo_Temp: av.Nombre_Periodo,
                        Monto_Programado: av.Monto_Periodo,
                        Avance_Programado_Pct: av.Porcentaje_Avance
                    });
                });
            });
        }
    } else if (tipo === 'APU') {
        // MULTI-CONCEPTO: Soportar array de conceptos o formato legado concepto_padre
        const listaConceptos = datos.conceptos || (datos.concepto_padre ? [datos.concepto_padre] : []);
        if (listaConceptos.length > 0) {
            datos.Catalogo_Conceptos = listaConceptos.map(c => ({
                Clave: c.Clave_Concepto,
                Descripcion: c.Descripcion_Concepto,
                Unidad: c.Unidad,
                Costo_Directo: c.Costo_Directo,
                Precio_Unitario: c.Precio_Unitario_Total
            }));
            // Aplanar todos los insumos de todos los conceptos
            datos.Matriz_Insumos = [];
            listaConceptos.forEach(c => {
                const insumosDelConcepto = c.insumos || [];
                insumosDelConcepto.forEach(ins => {
                    datos.Matriz_Insumos.push({
                        Clave_Concepto_Padre: c.Clave_Concepto,
                        Tipo_Insumo: ins.Tipo_Insumo,
                        Clave_Insumo: ins.Clave_Insumo,
                        Descripcion: ins.Descripcion,
                        Unidad: ins.Unidad,
                        Costo_Unitario: ins.Costo_Unitario,
                        Rendimiento_Cantidad: ins.Rendimiento_Cantidad,
                        Importe: ins.Importe,
                        Porcentaje_Incidencia: ins.Porcentaje_Incidencia
                    });
                });
            });
            // Compatibilidad: también aplanar insumos sueltos del formato legado
            if (datos.insumos && datos.Matriz_Insumos.length === 0) {
                datos.Matriz_Insumos = datos.insumos.map(ins => ({
                    Clave_Concepto_Padre: ins.Clave_Concepto_Padre,
                    Tipo_Insumo: ins.Tipo_Insumo,
                    Clave_Insumo: ins.Clave_Insumo,
                    Descripcion: ins.Descripcion,
                    Unidad: ins.Unidad,
                    Costo_Unitario: ins.Costo_Unitario,
                    Rendimiento_Cantidad: ins.Rendimiento_Cantidad,
                    Importe: ins.Importe,
                    Porcentaje_Incidencia: ins.Porcentaje_Incidencia
                }));
            }
        }
        // Porcentajes del contrato extraídos del APU
        if (datos.porcentajes_contrato) {
            const pct = datos.porcentajes_contrato;
            let pCtx = respuestaIA.parentContext;
            if (typeof pCtx === 'string') { try { pCtx = JSON.parse(pCtx); } catch (e) { pCtx = null; } }
            const idContratoCtx = respuestaIA.ID_Contrato_Contexto || (pCtx ? pCtx.ID_Contrato : null);
            if (idContratoCtx) {
                datos.Contratos = [{
                    ID_Contrato: idContratoCtx,
                    Pct_Indirectos_Oficina: pct.Pct_Indirectos_Oficina || 0,
                    Pct_Indirectos_Campo: pct.Pct_Indirectos_Campo || 0,
                    Pct_Indirectos_Totales: pct.Pct_Indirectos_Totales || 0,
                    Pct_Financiamiento: pct.Pct_Financiamiento || 0,
                    Pct_Utilidad: pct.Pct_Utilidad || 0,
                    Pct_Cargos_SFP: pct.Pct_Cargos_SFP || 0,
                    Pct_ISN: pct.Pct_ISN || 0
                }];
            }
        }
    }

    // --- RESOLUCIÓN GLOBAL DE CONTRATO DESDE LA RAÍZ ---
    let pContext = respuestaIA.parentContext;
    if (typeof pContext === 'string') {
        try { pContext = JSON.parse(pContext); } catch (e) { pContext = null; }
    }

    let globalIdContrato = respuestaIA.ID_Contrato_Contexto || (pContext ? pContext.ID_Contrato : null);

    // Si aún no hay ID global pero la IA extrajo un Numero_Contrato o ID_Contrato, buscarlo o crearlo
    const needleText = datos.Numero_Contrato || datos.ID_Contrato || (datos.Contrato && typeof datos.Contrato === 'string' ? datos.Contrato : null);
    if (!globalIdContrato && needleText) {
        if (!isNaN(parseInt(needleText)) && String(needleText).length < 6) {
            // Es un ID numérico directo
            globalIdContrato = parseInt(needleText);
        } else {
            // Es una clave de contrato (texto), buscarlo en la DB
            const contDb = dbSelect('Contratos', { Numero_Contrato: needleText });
            if (contDb && contDb.length > 0) {
                globalIdContrato = contDb[0].ID_Contrato;
            } else {
                // Auto-create minimal contract to hold the relationship
                const newCont = dbInsert('Contratos', { Numero_Contrato: needleText, Objeto_Contrato: "Importado automáticamente por IA" });
                if (newCont && newCont.ID_Contrato) {
                    globalIdContrato = newCont.ID_Contrato;
                }
            }
        }
    }

    // Mapa de traducción de IDs: Si la IA generó un "ID_Contrato": "temp_1", y dbInsert le asignó el entero 15, guardamos { "temp_1": 15 }
    const mapaIds = {};
    const ultimosIdsInsertados = {}; // Para linking automático de hijos sin ID temporal
    if (globalIdContrato) {
        ultimosIdsInsertados['Contratos'] = globalIdContrato;
    }

    jerarquia.forEach(tabla => {
        if (datos[tabla] && Array.isArray(datos[tabla])) {
            const headers = ESQUEMA_BD[tabla];
            const pkName = headers[0];

            // --- LÓGICA DE PURGA PARA MATRICES (SUSTITUCIÓN) ---
            if (tabla === 'Matriz_Insumos' && typeof dbDelete === 'function') {
                const listadoInsumos = datos[tabla];
                const idContratoParaResolve = resultados.ids['Contratos'] || globalIdContrato;
                // PRE-RESOLVER: Convertir Clave_Concepto_Padre → ID_Concepto ANTES de purgar
                if (idContratoParaResolve) {
                    listadoInsumos.forEach(ins => {
                        if (!ins.ID_Concepto && (ins.Clave_Concepto_Padre || ins.Clave_Padre)) {
                            const clave = ins.Clave_Concepto_Padre || ins.Clave_Padre;
                            const con = dbSelect('Catalogo_Conceptos', { Clave: clave, ID_Contrato: idContratoParaResolve });
                            if (con && con.length > 0) ins.ID_Concepto = con[0].ID_Concepto;
                        }
                        // Fallback: último concepto insertado en esta sesión
                        if (!ins.ID_Concepto && ultimosIdsInsertados['Catalogo_Conceptos']) {
                            ins.ID_Concepto = ultimosIdsInsertados['Catalogo_Conceptos'];
                        }
                    });
                }
                const conIds = [...new Set(listadoInsumos.map(ins => ins.ID_Concepto).filter(id => id && !String(id).startsWith('temp_')))];
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
                try {
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

                    // INYECCIÓN FORZOSA DE LLAVES FORÁNEAS (Basado en RELACIONES_BD)
                    const relacion = typeof RELACIONES_BD !== 'undefined' ? RELACIONES_BD[tabla] : null;
                    const idContratoContextoActual = resultados.ids['Contratos'] || globalIdContrato;

                    if (relacion) {
                        let tablaPadre = Array.isArray(relacion) ? relacion[0].padre : relacion.padre;
                        let llaveForanea = Array.isArray(relacion) ? relacion[0].fk : relacion.fk;

                        // Si la tabla es hija de Contratos, SIEMPRE forzar el ID_Contrato correcto
                        if (tablaPadre === 'Contratos' && idContratoContextoActual) {
                            registro[llaveForanea] = idContratoContextoActual;
                        } else if (tablaPadre !== 'Contratos' && ultimosIdsInsertados[tablaPadre]) {
                            if (!registro[llaveForanea]) {
                                registro[llaveForanea] = ultimosIdsInsertados[tablaPadre];
                            }
                        }
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
                            }
                        }

                        // Auto-calculate Subtotal/Net if possible
                        const bruto = parseFloat(registro.Monto_Bruto_Estimado) || 0;
                        if (!registro.Subtotal && bruto > 0) {
                            const ded = parseFloat(registro.Deduccion_Surv_05_Monto) || 0;
                            registro.Subtotal = (bruto - ded).toFixed(2);
                        }
                    }

                    // --- RESOLUCIÓN DE TEMPORALES Y LINKS JERÁRQUICOS ---
                    if (tabla === 'Matriz_Insumos') {
                        // Cruzar para obtener el ID_Concepto real desde la base de datos usando la Clave
                        if (!registro.ID_Concepto && (registro.Clave_Concepto_Padre || registro.Clave_Padre)) {
                            const clavePadre = registro.Clave_Concepto_Padre || registro.Clave_Padre;
                            if (idContratoContextoActual) {
                                const con = dbSelect('Catalogo_Conceptos', { Clave: clavePadre, ID_Contrato: idContratoContextoActual });
                                if (con && con.length > 0) {
                                    registro.ID_Concepto = con[0].ID_Concepto;
                                }
                            }
                        }
                        // Fallback: usar el último concepto insertado/actualizado en esta sesión
                        if (!registro.ID_Concepto && ultimosIdsInsertados['Catalogo_Conceptos']) {
                            registro.ID_Concepto = ultimosIdsInsertados['Catalogo_Conceptos'];
                        }
                        // Limpiar campos temporales que no existen en el esquema
                        delete registro.Clave_Concepto_Padre;
                        delete registro.Clave_Padre;
                        delete registro.ID_Contrato; // No existe en esquema Matriz_Insumos
                    }
                    if (tabla === 'Programa_Ejecucion') {
                        // 1. Resolver ID_Numero_Programa
                        if (!registro.ID_Numero_Programa) {
                            registro.ID_Numero_Programa = ultimosIdsInsertados['Programa'];
                        }

                        // 2. Resolver ID_Concepto (por Clave_Temp o por búsqueda inteligente)
                        if (!registro.ID_Concepto && (registro.Clave_Concepto_Temp || registro.Clave)) {
                            const clave = registro.Clave_Concepto_Temp || registro.Clave;
                            if (idContratoContextoActual) {
                                const con = dbSelect('Catalogo_Conceptos', { Clave: clave, ID_Contrato: idContratoContextoActual });
                                if (con && con.length > 0) registro.ID_Concepto = con[0].ID_Concepto;
                            }
                        }

                        // 3. Resolver ID_Programa_Periodo (por Periodo_Temp o búsqueda)
                        if (!registro.ID_Programa_Periodo && (registro.Periodo_Temp || registro.Periodo)) {
                            const perLabel = registro.Periodo_Temp || registro.Periodo;
                            const idProg = registro.ID_Numero_Programa || ultimosIdsInsertados['Programa'];
                            if (idProg) {
                                const per = dbSelect('Programa_Periodo', { Periodo: perLabel, ID_Numero_Programa: idProg });
                                if (per && per.length > 0) {
                                    registro.ID_Programa_Periodo = per[0].ID_Programa_Periodo;
                                    // Propagar fechas del periodo a la ejecución
                                    if (!registro.Fecha_Inicio && per[0].Fecha_Inicio) {
                                        registro.Fecha_Inicio = per[0].Fecha_Inicio;
                                    }
                                    if (!registro.Fecha_Fin && per[0].Fecha_Termino) {
                                        registro.Fecha_Fin = per[0].Fecha_Termino;
                                    }
                                }
                            }
                        }

                        // Limpiar temporales
                        delete registro.Clave_Concepto_Temp;
                        delete registro.Periodo_Temp;
                    }

                    let match = null;
                    const idOriginalDadoPorIA = registro[pkName];

                    if (registro[pkName] && typeof registro[pkName] !== 'string' && typeof registro[pkName] !== 'number') {
                        delete registro[pkName];
                    }

                    if (registro[pkName]) {
                        const cond = {}; cond[pkName] = registro[pkName];
                        const res = dbSelect(tabla, cond);
                        if (res && res.length > 0) match = res[0];
                    }

                    // --- BÚSQUEDA POR LLAVE SECUNDARIA (skName) ---
                    if (!match && skName && registro[skName]) {
                        const condSk = {}; condSk[skName] = registro[skName];
                        const resSk = dbSelect(tabla, condSk);
                        if (resSk && resSk.length > 0) {
                            match = resSk[0];
                        }
                    }

                    // --- SMART MATCHING Y DE-DUPLICACIÓN (Hijos) ---
                    if (!match) {
                        if (tabla === 'Catalogo_Conceptos' && registro.ID_Contrato) {
                            // Match por Clave primero
                            if (registro.Clave) {
                                const resClave = dbSelect('Catalogo_Conceptos', { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato });
                                if (resClave && resClave.length > 0) match = resClave[0];
                            }
                            // Match por Descripción como fallback
                            if (!match) {
                                const descNorm = normalizeText(registro.Descripcion || registro.Concepto || "");
                                if (descNorm) {
                                    const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: registro.ID_Contrato });
                                    match = catalogo.find(c => normalizeText(c.Descripcion) === descNorm);
                                }
                            }
                        } else if (tabla === 'Programa' && registro.ID_Contrato) {
                            const condProg = { ID_Contrato: registro.ID_Contrato };
                            if (registro.Tipo_Programa) condProg.Tipo_Programa = registro.Tipo_Programa;
                            const resProg = dbSelect('Programa', condProg);
                            if (resProg && resProg.length > 0) match = resProg[0];
                        } else if (tabla === 'Programa_Periodo' && registro.ID_Numero_Programa && registro.Periodo) {
                            const resPer = dbSelect('Programa_Periodo', { ID_Numero_Programa: registro.ID_Numero_Programa, Periodo: registro.Periodo });
                            if (resPer && resPer.length > 0) match = resPer[0];
                        } else if (tabla === 'Matriz_Insumos' && registro.ID_Concepto && registro.Clave_Insumo) {
                            const resMat = dbSelect('Matriz_Insumos', { ID_Concepto: registro.ID_Concepto, Clave_Insumo: registro.Clave_Insumo });
                            if (resMat && resMat.length > 0) match = resMat[0];
                        } else if (tabla === 'Programa_Ejecucion' && registro.ID_Numero_Programa && registro.ID_Concepto && registro.ID_Programa_Periodo) {
                            const resEjec = dbSelect('Programa_Ejecucion', {
                                ID_Numero_Programa: registro.ID_Numero_Programa,
                                ID_Concepto: registro.ID_Concepto,
                                ID_Programa_Periodo: registro.ID_Programa_Periodo
                            });
                            if (resEjec && resEjec.length > 0) match = resEjec[0];
                        }
                    }

                    // --- REGLAS DE NEGOCIO PARA MATRICES ---
                    if (tabla === 'Matriz_Insumos') {
                        const t = String(registro.Tipo_Insumo || registro.Tipo || "").toLowerCase();
                        if (t.includes('material')) registro.Tipo_Insumo = 1;
                        else if (t.includes('mano') || t.includes('jornal') || t.includes('cuadrilla')) registro.Tipo_Insumo = 2;
                        else if (t.includes('equipo') || t.includes('herra') || t.includes('maquin')) registro.Tipo_Insumo = 3;
                        else if (!isNaN(parseInt(t))) registro.Tipo_Insumo = parseInt(t);
                        else registro.Tipo_Insumo = 1;
                        delete registro.Tipo;
                    }

                    if (match) {
                        const dataMerged = { ...match };
                        const camposProtegidos = (tabla === 'Catalogo_Conceptos') ? ['Clave', 'Descripcion', 'Unidad'] : [];

                        Object.keys(registro).forEach(k => {
                            if (registro[k] !== undefined && registro[k] !== null && registro[k] !== '') {
                                if (!camposProtegidos.includes(k) || !match[k]) {
                                    dataMerged[k] = registro[k];
                                }
                            }
                        });

                        dbUpdate(tabla, dataMerged, { [pkName]: match[pkName] });
                        resultados.actualizados++;
                        resultados.ids[tabla] = match[pkName];
                        if (idOriginalDadoPorIA) mapaIds[idOriginalDadoPorIA] = match[pkName];
                        ultimosIdsInsertados[tabla] = match[pkName];

                    } else {
                        if (tabla === 'Catalogo_Conceptos') {
                            registro.Cantidad_Contratada = registro.Cantidad_Contratada || 1;
                            if (!registro.Orden) {
                                const last = dbSelect('Catalogo_Conceptos', { ID_Contrato: registro.ID_Contrato });
                                registro.Orden = (last ? last.length : 0) + 1;
                            }
                        }

                        if (idOriginalDadoPorIA) delete registro[pkName];

                        const resInsert = dbInsert(tabla, registro);
                        resultados.insertados++;
                        const newId = resInsert[pkName] || registro[pkName];

                        if (newId) {
                            registro[pkName] = newId;
                            ultimosIdsInsertados[tabla] = newId;
                            resultados.ids[tabla] = newId;
                            if (idOriginalDadoPorIA) mapaIds[idOriginalDadoPorIA] = newId;
                        }
                    }
                } catch (errReg) {
                    console.error('Error procesando registro en tabla ' + tabla + ':', errReg);
                }
            });
        }
    });

    return {
        success: true,
        message: "Procesamiento IA completado con UPSERT Inteligente",
        detalles: resultados,
        ids: resultados.ids
    };
}

/**
 * Compara los datos extraídos por la IA con la base de datos actual para detectar cambios.
 */
function analizarCambiosIA(datos) {
    if (!datos) return generarRespuesta(false, "No hay datos para analizar.");

    // --- NORMALIZACIÓN PREVIA (MISMA QUE guardarDatosIA) ---
    const tipo = datos.tipo_documento || "DESCONOCIDO";
    let pCtx = datos.parentContext;
    if (typeof pCtx === 'string') { try { pCtx = JSON.parse(pCtx); } catch (e) { pCtx = null; } }
    const idContratoCtx = datos.ID_Contrato_Contexto || (pCtx ? pCtx.ID_Contrato : null);

    if (tipo === 'APU') {
        const listaConceptos = datos.conceptos || (datos.concepto_padre ? [datos.concepto_padre] : []);
        if (listaConceptos.length > 0 && !datos.Catalogo_Conceptos) {
            datos.Catalogo_Conceptos = listaConceptos.map(c => ({
                Clave: c.Clave_Concepto,
                Descripcion: c.Descripcion_Concepto,
                Unidad: c.Unidad,
                Costo_Directo: c.Costo_Directo,
                Precio_Unitario: c.Precio_Unitario_Total,
                ID_Contrato: idContratoCtx
            }));
        }
        if (!datos.Matriz_Insumos) {
            datos.Matriz_Insumos = [];
            listaConceptos.forEach(c => {
                (c.insumos || []).forEach(ins => {
                    datos.Matriz_Insumos.push({
                        Clave_Concepto_Padre: c.Clave_Concepto,
                        Tipo_Insumo: ins.Tipo_Insumo,
                        Clave_Insumo: ins.Clave_Insumo,
                        Descripcion: ins.Descripcion,
                        Unidad: ins.Unidad,
                        Costo_Unitario: ins.Costo_Unitario,
                        Rendimiento_Cantidad: ins.Rendimiento_Cantidad,
                        Importe: ins.Importe,
                        Porcentaje_Incidencia: ins.Porcentaje_Incidencia
                    });
                });
            });
            // Compatibilidad formato legado
            if (datos.insumos && datos.Matriz_Insumos.length === 0) {
                datos.Matriz_Insumos = datos.insumos.map(ins => ({
                    Clave_Concepto_Padre: ins.Clave_Concepto_Padre,
                    Tipo_Insumo: ins.Tipo_Insumo,
                    Clave_Insumo: ins.Clave_Insumo,
                    Descripcion: ins.Descripcion,
                    Unidad: ins.Unidad,
                    Costo_Unitario: ins.Costo_Unitario,
                    Rendimiento_Cantidad: ins.Rendimiento_Cantidad,
                    Importe: ins.Importe,
                    Porcentaje_Incidencia: ins.Porcentaje_Incidencia
                }));
            }
            // Pre-resolver ID_Concepto para matching
            if (idContratoCtx) {
                datos.Matriz_Insumos.forEach(ins => {
                    if (ins.Clave_Concepto_Padre) {
                        const con = dbSelect('Catalogo_Conceptos', { Clave: ins.Clave_Concepto_Padre, ID_Contrato: idContratoCtx });
                        if (con && con.length > 0) ins.ID_Concepto = con[0].ID_Concepto;
                    }
                });
            }
        }
    } else if (tipo === 'PROGRAMA') {
        if (datos.conceptos_programados && !datos.Catalogo_Conceptos) {
            datos.Catalogo_Conceptos = datos.conceptos_programados.map(cp => ({
                Clave: cp.Clave_Concepto,
                Descripcion: cp.Descripcion_Limpiada,
                Unidad: cp.Unidad,
                Cantidad_Contratada: cp.Cantidad || 1,
                Precio_Unitario: cp.Precio_Unitario || null,
                Importe_Total_Sin_IVA: cp.Importe_Total,
                ID_Contrato: idContratoCtx
            }));
        }
    }

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
