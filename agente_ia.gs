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

5. **FIANZA / GARANTÍA**:
   - Palabras clave: "PÓLIZA DE FIANZA", "AFIANZADORA", "CUMPLIMIENTO", "ANTICIPO", "VICIOS OCULTOS", "INSTITUCIÓN DE SEGUROS Y DE FIANZAS".
   - ESTRUCTURA: Documento emitido por una afianzadora garantizando una obligación de un contrato específico.

INSTRUCCIÓN: Analiza el contenido y responde ÚNICAMENTE con un JSON con el campo "clase".
Clases permitidas: 'Contratos', 'CAF', 'Programa', 'Matriz_Insumos', 'Fianza' o 'Desconocido'.
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

        const response = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=" + GEMINI_API_KEY, options);
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

REGLAS DE EXTRACCIÓN (CRÍTICO):
1. **PORCENTAJES**: Extrae los valores como números (ej. 30 para 30%). SI NO APARECEN, ASIGNA 0. NUNCA DEJES ESTOS CAMPOS VACÍOS.
2. **ESTADO**: Se refiere a la Entidad Federativa de México donde se realiza la obra (ej. "CIUDAD DE MÉXICO", "JALISCO"). NO es el estatus operativo.
3. **MONTOS**: Extrae los montos totales y de fianzas. Si no hay fianza, el monto es 0.0.
4. **FECHAS**: Devuelve siempre en formato YYYY-MM-DD.

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "CONTRATO",
  "datos_extraidos": {
    "Numero_Contrato": "Clave del contrato",
    "Objeto_Contrato": "Descripción",
    "Estado": "ESTADO DE LA REPÚBLICA (Entidad Federativa)",
    "Tipo_Contrato": "OBRA PUBLICA / SERVICIOS",
    "Area_Responsable": "Unidad administrativa",
    "No_Concurso": "Número de licitación",
    "Modalidad_Adjudicacion": "Licitación / Invitación / Adjudicación Directa",
    "Fecha_Adjudicacion": "YYYY-MM-DD",
    "Fecha_Firma": "YYYY-MM-DD",
    "Fecha_Inicio_Obra": "YYYY-MM-DD",
    "Fecha_Fin_Obra": "YYYY-MM-DD",
    "Plazo_Ejecucion_Dias": 120,
    "Montos": {
      "Monto_Total_Sin_IVA": 0.0,
      "Monto_Total_Con_IVA": 0.0
    },
    "Anticipo": {
      "pct_Anticipo": 0.0,
      "Monto_Anticipo": 0.0
    },
    "Fianzas": {
      "Cumplimiento": { "No_Fianza": "N/A", "Monto": 0.0, "pct": 0.0 },
      "Anticipo": { "No_Fianza": "N/A", "Monto": 0.0, "pct": 0.0 },
      "Vicios_Ocultos": { "No_Fianza": "N/A", "Monto": 0.0, "pct": 0.0 }
    },
    "Retenciones_y_Penas": {
      "pct_Penas_Convencionales": 0.0,
      "pct_Retencion_Incumplimiento": 0.0,
      "Otras_Retenciones_Pct": 0.0
    },
    "porcentajes_contrato": {
      "Pct_Indirectos_Oficina": 0.0,
      "Pct_Indirectos_Campo": 0.0,
      "Pct_Indirectos_Totales": 0.0,
      "Pct_Financiamiento": 0.0,
      "Pct_Utilidad": 0.0,
      "Pct_Cargos_SFP": 0.5025,
      "Pct_ISN": 0.0
    },
    "Contratista": {
      "Razon_Social": "Nombre empresa",
      "RFC": "RFC",
      "Representante_Legal": "Nombre"
    },
    "ID_Convenio_Vinculado": "Contexto"
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

REGLAS CRÍTICAS DE EXTRACCIÓN Y ANCLAJE (REGLA DE ORO):
1. **COSTO DIRECTO (CRÍTICO)**: Debes extraer el valor del Costo Directo que aparece en la carátula o resumen de cada concepto. Es la suma antes de indirectos.
2. **INTEGRIDAD DE INSUMOS**: Extrae el 100% de los insumos listados para cada concepto. No resumas ni omitas ninguno.
3. **ANCLAJE VERTICAL**: Los documentos de APU suelen tener columnas muy juntas. Imagina una línea vertical que baja desde los encabezados "Unidad", "Costo Unitario" e "Importe". Cualquier número debe estar alineado bajo su respectiva columna.
4. **DETECCIÓN DE FILAS**: Un insumo termina cuando encuentras una nueva Clave_Insumo o un nuevo Tipo_Insumo. No mezcles descripciones de dos insumos.
5. **RECONSTRUCCIÓN DE TABLAS**: Si el PDF es de texto plano, utiliza el espaciado de caracteres para identificar dónde termina una columna y empieza la otra.
6. **INTEGRIDAD**: Extrae TODOS los conceptos del documento, no solo el primero.
7. **COSTOS**: El "Costo_Directo" de cada concepto es la SUMA de todos los importes de sus insumos si no está explícito. El "Precio_Unitario_Total" incluye además Indirectos y Utilidad.
8. **PORCENTAJES**: Extráelos como decimales (ej. 27.62). Si no aparecen, usa 0.
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
    "Numero_Acuerdo": "Folio o número de acuerdo del convenio (ej. CT/1A-ORD/11-ABRIL-25/X)",
    "Nombre_Fondo": "Nombre del fondo (ej. FONADIN, BANOBRAS)",
    "Monto_Apoyo": 250000000.0,
    "Fecha_Firma": "YYYY-MM-DD",
    "Vigencia_Fin": "YYYY-MM-DD",
    "Objeto_Programa": "Descripción del objeto o programa",
    "Estado": "Estado de la República / Ubicación",
    "Entidades_Involucradas": ["BANOBRAS", "SHCP", "..."]
  }
}

REGLA CRÍTICA:
- El Número_Acuerdo es el identificador principal. No lo confundas con números de fiduciario (ej. 1936).
`;

const PROMPT_IMPORTACION_CATALOGO = `
Rol y Objetivo:
Eres un Experto en Auditoría de Obra y Presupuestos. Tu misión es extraer el CATÁLOGO DE CONCEPTOS de un documento técnico.

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "CATALOGO",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "conceptos": [
      {
        "Clave": "Clave del concepto (ej. PREL-01)",
        "Descripcion": "Descripción completa del concepto",
        "Unidad": "Unidad de medida (ej. M2, Lote, Pza)",
        "Cantidad": 100.50,
        "Precio_Unitario": 1250.33,
        "Importe": 125658.16,
        "Costo_Directo": 1000.00
      }
    ]
  }
}

REGLAS CRÍTICAS DE LECTURA (ANCLAJE ESPACIAL):
1. **COSTO DIRECTO**: Extrae el Costo Directo de cada concepto si está disponible. Es vital para la matriz.
2. **TABLAS DENSAS**: Usa los encabezados "Clave", "Unidad" e "Importe" como anclas verticales.
3. **INTEGRIDAD**: Extrae TODO el listado. Si el documento tiene múltiples páginas, procésalas todas.
4. **CONTEXTO**: Ignora logotipos o textos legales en los encabezados/pies de página; enfócate en el cuerpo de la tabla.
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
    "Periodos_Programados": [
      {
        "Nombre_Periodo": "MES AÑO (ej. DICIEMBRE 2025)",
        "Fecha_Inicio": "YYYY-MM-DD",
        "Fecha_Termino": "YYYY-MM-DD",
        "Conceptos": [
          {
            "Clave_Concepto": "N3-25-5",
            "Descripcion_Contextual": "Descripción",
            "Unidad": "LOTE",
            "Cantidad": 1.00,
            "Precio_Unitario": 255011.33,
            "Importe_Total": 255011.33,
            "Monto_Periodo": 12750.57,
            "Porcentaje_Avance": 5.0
          }
        ]
      }
    ]
  }
}

REGLAS CRÍTICAS DE LECTURA MATRICIAL (ANCLAJE DE PERIODOS):
1. **PERIODO OBLIGATORIO**: El 'Nombre_Periodo' debe ser SIEMPRE el nombre del mes seguido del año (ej. "AGOSTO 2025", "SEPTIEMBRE 2025"). NUNCA lo dejes vacío.
2. **IDENTIFICACIÓN DE COLUMNAS**: Identifica cada Periodo y agrúpalos como objetos raíces en 'Periodos_Programados'.
3. **LECTURA VERTICAL**: Para CADA PERIODO, lee hacia abajo. Registra en 'Conceptos' de ese Periodo TODOS los montos y porcentajes asociados alineados verticalmente bajo esa columna.
4. **INTEGRIDAD POR PERIODO**: Es OBLIGATORIO extraer todos los periodos (desde el primero hasta el último).
5. **NOMBRES EN MAYÚSCULAS**: Escribe siempre los nombres de los periodos en MAYÚSCULAS.
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
function processDocumentWithAI(base64Data, mimeType, targetTable = null, contextJson = null) {
    try {
        let parentContext = contextJson;
        if (contextJson && typeof contextJson === 'string') {
            try {
                parentContext = JSON.parse(contextJson);
            } catch (e) {
                console.error("Error parseando contextJson:", e);
                parentContext = null;
            }
        }

        // Permitimos base64Data nulo solo si el objetivo es Fianza (validación de contexto)
        if (!base64Data && targetTable !== 'Fianza') {
            throw new Error("Datos del archivo faltantes y no es una validación de contexto de fianza.");
        }

        // 1. Preparar payload para Gemini
        let b64 = null;
        if (base64Data) {
            b64 = base64Data;
            if (b64.includes('base64,')) {
                b64 = b64.split('base64,')[1];
            }
        }

        // 2. Seleccionar el Modelo y System Instruction basado en el objetivo (Optimización de Costos)
        let modelId = "gemini-2.5-flash-lite"; // Nivel Económico (CAF, Otros)
        let systemPrompt = PROMPT_IMPORTACION_CONTRATOS;

        if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') {
            systemPrompt = PROMPT_IMPORTACION_MATRICES;
            modelId = "gemini-3.1-flash-lite-preview"; // Nivel Eficiente (Performance)
        } else if (targetTable === 'Estimaciones' || targetTable === 'Facturas') {
            systemPrompt = PROMPT_VALIDACION_ESTIMACIONES;
            modelId = "gemini-2.5-flash-lite"; // Nivel Económico
        } else if (targetTable === 'Convenios_Recurso' || targetTable === 'CAF') {
            systemPrompt = PROMPT_IMPORTACION_CAF;
            modelId = "gemini-2.5-flash-lite"; // Nivel Económico
        } else if (targetTable === 'Programa' || targetTable === 'Programa_Ejecucion') {
            systemPrompt = PROMPT_IMPORTACION_PROGRAMA;
            modelId = "gemini-3.1-flash-lite-preview"; // Nivel Eficiente
        } else if (targetTable === 'Catalogo_Conceptos' || targetTable === 'CATALOGO') {
            systemPrompt = PROMPT_IMPORTACION_CATALOGO;
            modelId = "gemini-3.1-flash-lite-preview"; // Nivel Eficiente
        } else if (targetTable === 'Fianza' || targetTable === 'Contratos') {
            systemPrompt = targetTable === 'Fianza' ? PROMPT_VALIDACION_FIANZA : PROMPT_IMPORTACION_CONTRATOS;
            modelId = "gemini-3.1-pro-preview"; // Nivel Inteligencia (Crítico)
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${GEMINI_API_KEY}`;

        // 3. Construir prompt adicional con contexto dinámico
        let promptAdicional = `Lee este archivo oficial. Debes extraer los datos en formato JSON basado en el esquema solicitado.`;
        if (targetTable) {
            promptAdicional += `\n\nATENCIÓN: El documento se carga para la tabla '${targetTable}'.`;

            if (targetTable === 'Fianza') {
                try {
                    const modeloHtml = HtmlService.createHtmlOutputFromFile('modeloFianza').getContent();
                    promptAdicional += `\n\nCONTEXTO CRÍTICO (MODELO OFICIAL DEL DOF):\n${modeloHtml}`;
                } catch (e) {
                    console.error("Error cargando modeloFianza:", e);
                }
            }

            if (parentContext) {
                promptAdicional += `\n\nCONTEXTO REGISTRO PADRE: ${JSON.stringify(parentContext)}`;

                if (parentContext.Link_Sharepoint) {
                    promptAdicional += `\nINSTRUCCIÓN: Usa el link '${parentContext.Link_Sharepoint}' para el campo 'Link_Sharepoint'.`;
                }

                // Contexto para Matrices: Catálogo Actual
                if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') {
                    const idCont = parentContext.ID_Contrato || parentContext.idContrato;
                    if (idCont) {
                        const cat = dbSelect('Catalogo_Conceptos', { ID_Contrato: idCont });
                        if (cat && cat.length > 0) {
                            promptAdicional += `\n\nCATÁLOGO ACTUAL PARA MATCHING: ${JSON.stringify(cat.map(c => ({ ID_Concepto: c.ID_Concepto, Clave: c.Clave, Descripcion: c.Descripcion })))}`;
                        }
                    }
                }

                // Contexto para Estimaciones/Facturas
                if ((targetTable === 'Estimaciones' || targetTable === 'Facturas') && parentContext.ID_Estimacion) {
                    const est = dbSelect('Estimaciones', { ID_Estimacion: parentContext.ID_Estimacion });
                    if (est && est.length > 0) {
                        const det = dbSelect('Detalle_Estimacion', { ID_Estimacion: parentContext.ID_Estimacion });
                        promptAdicional += `\n\nDATOS SISTEMA (ESTIMACIÓN): ${JSON.stringify(est[0])}\nCONCEPTOS YA REGISTRADOS: ${JSON.stringify(det.map(d => ({ ID_Concepto: d.ID_Concepto, Cantidad: d.Cantidad_Estimada_Periodo })))}`;

                        const vals = dbSelect('Validacion_Archivos', { ID_Estimacion: parentContext.ID_Estimacion });
                        if (vals && vals.length > 0) {
                            const lastVal = vals[vals.length - 1];
                            if (lastVal.Checklist_JSON) {
                                promptAdicional += `\n\nCONTEXTO_AUDITORIA_PREVIA (PRESERVA APROBADOS): ${lastVal.Checklist_JSON}`;
                            }
                        }
                    }
                }

                // Contexto para Programa (Periodos y Conceptos)
                if (targetTable === 'Programa' || targetTable === 'Programa_Ejecucion') {
                    const idCont = parentContext.ID_Contrato || parentContext.idContrato;
                    if (idCont) {
                        const prog = dbSelect('Programa', { ID_Contrato: idCont });
                        if (prog && prog.length > 0) {
                            const idProg = prog[0].ID_Numero_Programa;
                            const pers = dbSelect('Programa_Periodo', { ID_Numero_Programa: idProg });
                            if (pers && pers.length > 0) {
                                promptAdicional += `\n\nPERIODOS EXISTENTES (USA EXACTAMENTE ESTOS NOMBRES): ${JSON.stringify(pers.map(p => ({ Nombre: p.Periodo, Inicio: p.Fecha_Inicio, Fin: p.Fecha_Termino || p.Fecha_Fin })))}`;
                            }
                        }
                        const cat = dbSelect('Catalogo_Conceptos', { ID_Contrato: idCont });
                        if (cat && cat.length > 0) {
                            promptAdicional += `\n\nCATÁLOGO PARA MATCHING DE PROGRAMA: ${JSON.stringify(cat.map(c => ({ Clave: c.Clave, Descripcion: c.Descripcion })))}`;
                        }
                    }
                }
            }
        }

        const parts = [];
        if (b64) {
            parts.push({ inlineData: { mimeType: mimeType || 'application/pdf', data: b64 } });
        }
        parts.push({ text: promptAdicional });

        const payload = {
            system_instruction: { parts: [{ text: systemPrompt }] },
            contents: [{ parts: parts }],
            generationConfig: { temperature: 0.1, responseMimeType: "application/json" }
        };

        const response = UrlFetchApp.fetch(url, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        const resultJson = JSON.parse(response.getContentText());
        if (response.getResponseCode() !== 200) {
            throw new Error(`Gemini Error ${response.getResponseCode()}: ${resultJson.error?.message || response.getContentText()}`);
        }

        const llmString = resultJson.candidates[0].content.parts[0].text;

        // --- LOG PROVISIONAL DE EXTRACCIÓN ---
        console.log("RAW_AI_EXTRACTION [" + (targetTable || "GENERAL") + "]:", llmString);

        const extraction = JSON.parse(llmString.replace(/```json/g, '').replace(/```/g, '').trim());
        const datosFinales = extraction.datos_extraidos || extraction.datos || extraction;

        const resFormateada = {
            success: true,
            accion: (targetTable === 'Estimaciones' || targetTable === 'Facturas') ? "auditoria_validación" : (targetTable === 'Fianza' ? "auditoria_fianza" : "importar_datos"),
            datos: datosFinales,
            datos_extraidos: datosFinales, // Añadido para compatibilidad
            tipo_documento: targetTable === 'Fianza' ? 'FIANZA' : (extraction.tipo_documento || "DESCONOCIDO"),
            nivel_confianza: extraction.nivel_confianza || "Bajo",
            parentContext: parentContext,
            checklist_auditoria: extraction.checklist_auditoria || null,
            informe_auditoria: extraction.informe_auditoria || null,
            estado_global: extraction.estado_global || null
        };

        if (targetTable === 'Fianza' && extraction.validacion) {
            resFormateada.validacion = extraction.validacion;
            resFormateada.tipo_identificado = extraction.tipo_identificado;
        }

        return generarRespuesta(true, resFormateada);

    } catch (e) {
        console.error("Error en processDocumentWithAI:", e);
        return generarRespuesta(false, e.toString(), 'processDocumentWithAI');
    }
}


const PROMPT_VALIDACION_FIANZA = `
### ROL: Perito Legal SGC (Especialista en Fianzas Gubernamentales)
Tu misión es realizar un peritaje técnico y legal comparando la Póliza de Fianza proporcionada contra el "Modelo Oficial del DOF" (modeloFianza.html) y los "Datos del Contrato" de nuestro sistema.

### RECURSOS DISPONIBLES EN EL CONTEXTO:
1. **Archivo Adjunto:** La imagen o PDF de la póliza de fianza.
2. **Modelo Oficial (DOF):** Contenido en modeloFianza.html. 
   - **Ley Aplicable:** Verifica en 'parentContext.Ley_Aplicable' si es LOPSRM (Obras) o LAASSP (Adquisiciones).
   - **Articulado:** 
     - Si es LOPSRM: Los artículos de la LISF aplicables suelen ser 279 (Entidades) o 282 (Dependencias). La fianza DEBE citar la 'Ley de Obras Públicas y Servicios Relacionados con las Mismas'.
     - Si es LAASSP: La fianza DEBE citar la 'Ley de Adquisiciones, Arrendamientos y Servicios del Sector Público'.
3. **Datos del Sistema (parentContext):** Contiene la información oficial (RFCs, Montos, Ley).

### PASO 1: Validación Contextual y Legal (LEY Y ARTICULADO)
- **Cruce de Ley:** Verifica que el cuerpo de la fianza cite la ley correcta indicada en 'parentContext.Ley_Aplicable'. Es un error crítico citar la ley de adquisiciones en un contrato de obra.
- **Fiado (Contratista):** Coincidencia exacta con "Contratista_Razon_Social" y "Contratista_RFC".
- **Monto de la Fianza:** 
  - Si es CUMPLIMIENTO: Verifica contra "Monto_Esperado_Garantia".
  - Si es ANTICIPO: Debe ser el 100% del anticipo ("Monto_Esperado_Garantia").

### PASO 2: Validación de Cláusulas (MODELO DOF)
- **Vigencia:** Debe ser indeterminada hasta cumplimiento total.
- **Indivisibilidad:** Obligatorio mencionar que la garantía es indivisible.
- **Suspensión:** Verifica que la cláusula de suspensión coincida con el texto del modelo DOF para la ley específica (LOPSRM o LAASSP).

### OUTPUT JSON REQUERIDO:
{
  "tipo_identificado": "Cumplimiento / Anticipo - [Ley Detectada]",
  "datos_extraidos": {
    "num_fianza": "Póliza-12345",
    "monto_documento": 123456.78,
    "monto_esperado_system": 0.00,
    "ley_citada_en_documento": "Texto detectado",
    "nombre_afianzadora": "Nombre de la Institución Afianzadora",
    "pct_fianza_documento": 10.0
  },
  "validacion": {
    "estatus": "VALIDA" | "RECHAZADA",
    "checklist": [
      {
        "criterio": "Cruce de Ley Aplicable",
        "pasa": true,
        "observacion": "Cita correctamente la LOPSRM."
      },
      {
        "criterio": "Articulado LISF",
        "pasa": false,
        "observacion": "Cita el Art. 279 pero por ser Dependencia debería citar el 282."
      }
    ],
    "resumen_critico": "Resumen técnico indicando la validez legal y el apego al modelo oficial de la ley correspondiente."
  }
}
`;

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

    // --- LOG DE GUARDADO ---
    console.log("DETALLES_GUARDADO_IA [" + tipo + "]:", JSON.stringify(datos));

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

        // Helper para escalar porcentajes (IA devuelve 30 para 30%, guardamos 30)
        const scalePct = (val) => {
            let n = parseFloat(val);
            if (isNaN(n)) return 0;
            // Si el valor ya es una fracción (ej. 0.3), lo llevamos a entero (30)
            // Si es > 1 (ej. 30), asumimos que ya es entero y lo dejamos tal cual
            return (n > 0 && n <= 1) ? n * 100 : n;
        };

        const contratista = {
            Razon_Social: c.Contratista?.Razon_Social,
            RFC: c.Contratista?.RFC,
            Representante_Legal: c.Contratista?.Representante_Legal
        };
        const contrato = {
            Numero_Contrato: c.Numero_Contrato,
            Objeto_Contrato: c.Objeto_Contrato,
            Estado: c.Estado || 'CDMX', // Ahora es ubicación geográfica
            Tipo_Contrato: c.Tipo_Contrato,
            Area_Responsable: c.Area_Responsable,
            No_Concurso: c.No_Concurso,
            Modalidad_Adjudicacion: c.Modalidad_Adjudicacion,
            Fecha_Adjudicacion: c.Fecha_Adjud_acion || c.Fecha_Adjudicacion,
            Fecha_Firma: c.Fecha_Firma,
            Fecha_Inicio_Obra: c.Fecha_Inicio_Obra,
            Fecha_Fin_Obra: c.Fecha_Fin_Obra,
            Plazo_Ejecucion_Dias: c.Plazo_Ejecucion_Dias,
            Monto_Total_Sin_IVA: c.Montos?.Monto_Total_Sin_IVA,
            Monto_Total_Con_IVA: c.Montos?.Monto_Total_Con_IVA,
            ID_Convenio_Vinculado: idConvenioVinculado || c.ID_Convenio_Vinculado,
            ID_Contratista: c.ID_Contratista || datos.ID_Contratista,
            // Anticipo
            pct_Anticipo: scalePct(c.Anticipo?.pct_Anticipo || c.pct_Anticipo),
            Monto_Anticipo: c.Anticipo?.Monto_Anticipo || c.Monto_Anticipo || 0,
            No_Fianza_Anticipo: c.Fianzas?.Anticipo?.No_Fianza || c.No_Fianza_Anticipo,
            // Cumplimiento
            pct_Cumplimiento: scalePct(c.Fianzas?.Cumplimiento?.pct),
            Monto_Fianza_Cumplimiento: c.Fianzas?.Cumplimiento?.Monto || 0,
            No_Fianza_Cumplimiento: c.Fianzas?.Cumplimiento?.No_Fianza || c.No_Fianza_Cumplimiento,
            // Vicios Ocultos
            pct_Vicios_Ocultos: scalePct(c.Fianzas?.Vicios_Ocultos?.pct),
            No_Fianza_Vicios_Ocultos: c.Fianzas?.Vicios_Ocultos?.No_Fianza,
            Monto_Fianza_Vicios_Ocultos: c.Fianzas?.Vicios_Ocultos?.Monto,
            // Indirectos y Otros (Sobrecostos escalados)
            Pct_Indirectos_Oficina: scalePct(c.porcentajes_contrato?.Pct_Indirectos_Oficina),
            Pct_Indirectos_Campo: scalePct(c.porcentajes_contrato?.Pct_Indirectos_Campo),
            Pct_Indirectos_Totales: scalePct(c.porcentajes_contrato?.Pct_Indirectos_Totales),
            Pct_Financiamiento: scalePct(c.porcentajes_contrato?.Pct_Financiamiento),
            Pct_Utilidad: scalePct(c.porcentajes_contrato?.Pct_Utilidad),
            Pct_Cargos_SFP: scalePct(c.porcentajes_contrato?.Pct_Cargos_SFP),
            Pct_ISN: scalePct(c.porcentajes_contrato?.Pct_ISN),
            // Penas y Retenciones
            pct_Penas_Convencionales: scalePct(c.Retenciones_y_Penas?.pct_Penas_Convencionales),
            pct_Retencion_Incumplimiento: scalePct(c.Retenciones_y_Penas?.pct_Retencion_Incumplimiento),
            Otras_Retenciones_Pct: scalePct(c.Retenciones_y_Penas?.Otras_Retenciones_Pct),
            Link_Sharepoint: c.Link_Sharepoint
        };
        datos.Contratistas = [contratista];
        datos.Contratos = [contrato];
    } else if (tipo === 'CAF') {
        datos.Convenios_Recurso = [{
            Numero_Acuerdo: datos.Numero_Acuerdo,
            Nombre_Fondo: datos.Nombre_Fondo,
            Monto_Apoyo: datos.Monto_Apoyo,
            Fecha_Firma: datos.Fecha_Firma,
            Vigencia_Fin: datos.Vigencia_Fin,
            Objeto_Programa: datos.Objeto_Programa,
            Estado: datos.Estado,
            Link_Sharepoint: datos.Link_Sharepoint
        }];
    } else if (tipo === 'CATALOGO') {
        if (datos.conceptos && Array.isArray(datos.conceptos)) {
            datos.Catalogo_Conceptos = datos.conceptos;
        }
    } else if (tipo === 'PROGRAMA') {
        // Mapear Programa
        datos.Programa = [{
            Programa: datos.Nombre_Programa || datos.Programa || "Programa General de Obra",
            Tipo_Programa: datos.Tipo_Programa || "MENSUAL",
            Fecha_Inicio: datos.Fechas_Generales?.Fecha_Inicio || null,
            Fecha_Termino: datos.Fechas_Generales?.Fecha_Fin || null,
            Plazo_Ejecucion_Dias: datos.Fechas_Generales?.Plazo_Dias || null
        }];
        // Períodos y Conceptos (NUEVO FORMATO ORIENTADO A PERIODOS)
        if (datos.Periodos_Programados) {
            datos.Programa_Periodo = [];
            datos.Programa_Ejecucion = [];

            const conceptosUnicos = {};

            datos.Periodos_Programados.forEach((per, idx) => {
                const nombrePer = String(per.Nombre_Periodo || "").trim().toUpperCase();

                // 1. Guardar Periodo
                const perObj = {
                    Numero_Periodo: idx + 1,
                    Periodo: nombrePer,
                    Fecha_Inicio: per.Fecha_Inicio,
                    Fecha_Termino: per.Fecha_Termino || per.Fecha_Fin
                };
                datos.Programa_Periodo.push(perObj);

                // 2. Procesar Conceptos del Periodo
                if (per.Conceptos && Array.isArray(per.Conceptos)) {
                    per.Conceptos.forEach(cp => {
                        // Rescatar concepto único para Catálogo_Conceptos si no existe
                        if (cp.Clave_Concepto && !conceptosUnicos[cp.Clave_Concepto]) {
                            conceptosUnicos[cp.Clave_Concepto] = {
                                Clave: cp.Clave_Concepto,
                                Descripcion: cp.Descripcion_Contextual || cp.Descripcion_Limpiada || "N/A",
                                Unidad: cp.Unidad || "SRV",
                                Cantidad_Contratada: cp.Cantidad || 1,
                                Precio_Unitario: cp.Precio_Unitario || null,
                                Importe_Total_Sin_IVA: cp.Importe_Total || 0
                            };
                        }

                        // SOLO DELTAS: no registrar actividad en cero
                        const monto = parseFloat(cp.Monto_Periodo) || 0;
                        const pct = parseFloat(cp.Porcentaje_Avance) || 0;
                        if (monto === 0 && pct === 0) return;

                        // NUEVO: Registrar Programa_Ejecucion INYECTANDO Fechas del Periodo
                        datos.Programa_Ejecucion.push({
                            Clave_Concepto_Temp: cp.Clave_Concepto,
                            Descripcion_Temp: cp.Descripcion_Contextual || cp.Descripcion_Limpiada,
                            Periodo_Temp: nombrePer,
                            Fecha_Inicio: per.Fecha_Inicio,
                            Fecha_Fin: per.Fecha_Termino,
                            Monto_Programado: cp.Monto_Periodo,
                            Avance_Programado_Pct: cp.Porcentaje_Avance
                        });
                    });
                }
            });

            // Reconstruir Catálogo de Conceptos uniendo todos los rescatados
            datos.Catalogo_Conceptos = Object.values(conceptosUnicos);

        } else if (datos.Periodos_Identificados && datos.conceptos_programados) {
            // LOGICA VIEJA (Por retrocompatibilidad si la IA aún devuelve el formato anterior)
            datos.Programa_Periodo = datos.Periodos_Identificados.map((p, i) => {
                const esObjeto = typeof p === 'object';
                return { Numero_Periodo: i + 1, Periodo: (esObjeto ? p.Nombre : p).trim().toUpperCase(), Fecha_Inicio: esObjeto ? p.Fecha_Inicio : null, Fecha_Termino: esObjeto ? p.Fecha_Termino : null };
            });
            datos.Catalogo_Conceptos = datos.conceptos_programados.map(cp => ({ Clave: cp.Clave_Concepto, Descripcion: cp.Descripcion_Limpiada, Unidad: cp.Unidad, Cantidad_Contratada: cp.Cantidad || 1, Precio_Unitario: cp.Precio_Unitario || null, Importe_Total_Sin_IVA: cp.Importe_Total }));
            datos.Programa_Ejecucion = [];
            datos.conceptos_programados.forEach(cp => {
                cp.avances_por_periodo?.forEach(av => {
                    const monto = parseFloat(av.Monto_Periodo) || 0; const pct = parseFloat(av.Porcentaje_Avance) || 0; if (monto === 0 && pct === 0) return;
                    let nombrePer = String(av.Nombre_Periodo || "").trim().toUpperCase();
                    datos.Programa_Ejecucion.push({ Clave_Concepto_Temp: cp.Clave_Concepto, Periodo_Temp: nombrePer, Monto_Programado: av.Monto_Periodo, Avance_Programado_Pct: av.Porcentaje_Avance });
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
    } else if (tipo === 'FIANZA') {
        const estatusValidacion = respuestaIA.validacion?.estatus || 'RECHAZADA';
        const detalles = respuestaIA.validacion?.diferencias_encontradas?.join(', ') || '';

        let contratoFianza = {
            Numero_Contrato: datos.contrato || datos.Numero_Contrato
        };

        if (estatusValidacion === 'VALIDA') {
            contratoFianza["Fianza_Numero"] = datos.num_fianza;
            contratoFianza["Fianza_Estatus"] = "Validada conforme a DOF";

            // Asignar los campos en BD según el tipo identificado por la IA
            const tIdent = (respuestaIA.tipo_identificado || "").toLowerCase();
            if (tIdent.includes("anticipo")) {
                contratoFianza["No_Fianza_Anticipo"] = datos.num_fianza;
                contratoFianza["Monto_Anticipo"] = datos.monto || datos.monto_documento;
            } else if (tIdent.includes("cumplimiento")) {
                contratoFianza["No_Fianza_Cumplimiento"] = datos.num_fianza;
                contratoFianza["Monto_Fianza_Cumplimiento"] = datos.monto || datos.monto_documento;
            } else if (tIdent.includes("vicio")) {
                contratoFianza["No_Fianza_Vicios_Ocultos"] = datos.num_fianza;
                contratoFianza["Monto_Fianza_Vicios_Ocultos"] = datos.monto || datos.monto_documento;
            }
        } else {
            contratoFianza["Fianza_Estatus"] = "Rechazada (" + detalles.slice(0, 100) + "...)";
        }

        // Si hay un contexto padre forzado, usarlo
        let pCtxF = respuestaIA.parentContext;
        if (typeof pCtxF === 'string') { try { pCtxF = JSON.parse(pCtxF); } catch (e) { } }
        if (pCtxF && pCtxF.ID_Contrato) {
            contratoFianza.ID_Contrato = pCtxF.ID_Contrato;
        }

        datos.Contratos = [contratoFianza];
    }

    // --- RESOLUCIÓN GLOBAL DE CONTRATO DESDE LA RAÍZ ---
    let pContext = respuestaIA.parentContext;
    if (typeof pContext === 'string') {
        try { pContext = JSON.parse(pContext); } catch (e) { pContext = null; }
    }

    let globalIdContrato = respuestaIA.ID_Contrato_Contexto || 
                           (pContext ? (pContext.ID_Contrato || pContext.idContrato || pContext.id_contrato) : null);

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
    const mapaPeriodosReales = {}; // NUEVO: Para vincular periodos por nombre sin usar IDs temporales
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

                    // --- RESOLUCIÓN DE TEMPORALES Y LINKS JERÁRQUICOS ---
                    const idContratoContextoActual = resultados.ids['Contratos'] || globalIdContrato;

                    if (tabla === 'Matriz_Insumos') {
                        // Cruzar para obtener el ID_Concepto real desde la base de datos usando la Clave
                        if (!registro.ID_Concepto && (registro.Clave_Concepto_Padre || registro.Clave_Padre)) {
                            const clavePadre = (registro.Clave_Concepto_Padre || registro.Clave_Padre).toString().trim();
                            if (idContratoContextoActual) {
                                // Búsqueda insensible a mayúsculas/minúsculas para mayor robustez
                                const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: idContratoContextoActual });
                                const conMatch = catalogo.find(c => String(c.Clave).trim().toUpperCase() === clavePadre.toUpperCase());
                                if (conMatch) {
                                    registro.ID_Concepto = conMatch.ID_Concepto;
                                }
                            }
                        }
                        // Limpiar campos temporales que no existen en el esquema
                        delete registro.Clave_Concepto_Padre;
                        delete registro.Clave_Padre;
                        delete registro.ID_Contrato; // No existe en esquema Matriz_Insumos
                    }
                    if (tabla === 'Programa_Ejecucion') {
                        // 1. Resolver ID_Concepto (Doble Match: Clave + Descripción)
                        if (!registro.ID_Concepto && (registro.Clave_Concepto_Temp || registro.Clave)) {
                            const clave = (registro.Clave_Concepto_Temp || registro.Clave).toString().trim();
                            if (idContratoContextoActual) {
                                // Intento 1: Por Clave (Insensible)
                                const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: idContratoContextoActual });
                                let matchClave = catalogo.find(c => String(c.Clave).trim().toUpperCase() === clave.toUpperCase());
                                
                                if (matchClave) {
                                    registro.ID_Concepto = matchClave.ID_Concepto;
                                } else {
                                    // Intento 2: Por Descripción (Contextual)
                                    const descNorm = normalizeText(registro.Descripcion_Temp || registro.Descripcion || registro.Descripcion_Limpiada || "");
                                    if (descNorm) {
                                        const matchDesc = catalogo.find(c => normalizeText(c.Descripcion) === descNorm);
                                        if (matchDesc) {
                                            registro.ID_Concepto = matchDesc.ID_Concepto;
                                        } else {
                                            // NUEVO: Si no existe el concepto, crearlo para evitar pérdida de datos
                                            const nuevoConcepto = dbInsert('Catalogo_Conceptos', {
                                                ID_Contrato: idContratoContextoActual,
                                                Clave: clave,
                                                Descripcion: registro.Descripcion_Temp || registro.Descripcion || "Concepto importado vía Programa",
                                                Unidad: registro.Unidad || 'S/U',
                                                Cantidad_Contratada: 1,
                                                Precio_Unitario: 0
                                            });
                                            if (nuevoConcepto) registro.ID_Concepto = nuevoConcepto.ID_Concepto;
                                        }
                                    }
                                }
                            }
                        }

                        // 3. Resolver ID_Programa_Periodo (NORMALIZADO)
                        if (!registro.ID_Programa_Periodo && (registro.Periodo_Temp || registro.Periodo)) {
                            const perLabel = (registro.Periodo_Temp || registro.Periodo).toString().trim().toUpperCase();

                            // Intentar recuperar el ID real insertado en esta sesión para este nombre de periodo
                            const keyNormalizada = Object.keys(mapaPeriodosReales).find(k => k.trim().toUpperCase() === perLabel);

                            if (keyNormalizada && mapaPeriodosReales[keyNormalizada]) {
                                registro.ID_Programa_Periodo = mapaPeriodosReales[keyNormalizada].ID_Programa_Periodo;
                            } else {
                                // Buscar en DB
                                const idProgPadre = ultimosIdsInsertados['Programa'] || (dbSelect('Programa', { ID_Contrato: idContratoContextoActual })[0] || {}).ID_Numero_Programa;
                                if (idProgPadre) {
                                    const periodosDB = dbSelect('Programa_Periodo', { ID_Numero_Programa: idProgPadre });
                                    let matchDB = periodosDB.find(p => (p.Periodo || "").toString().trim().toUpperCase() === perLabel);
                                    
                                    if (!matchDB) {
                                        matchDB = periodosDB.find(p => (p.Periodo || "").toString().trim().toUpperCase().includes(perLabel));
                                    }
                                    
                                    if (matchDB) registro.ID_Programa_Periodo = matchDB.ID_Programa_Periodo;
                                }
                            }

                            // Propagar fechas si ya tenemos el ID
                            if (registro.ID_Programa_Periodo) {
                                if (!registro.Fecha_Inicio || !registro.Fecha_Fin) {
                                    const pData = mapaPeriodosReales[perLabel] || dbSelect('Programa_Periodo', { ID_Programa_Periodo: registro.ID_Programa_Periodo })[0];
                                    if (pData) {
                                        if (!registro.Fecha_Inicio) registro.Fecha_Inicio = pData.Fecha_Inicio;
                                        if (!registro.Fecha_Fin) registro.Fecha_Fin = pData.Fecha_Termino || pData.Fecha_Fin;
                                    }
                                }
                            }
                        }

                        // NUEVO: RECÁLCULO DE IMPORTES Y PORCENTAJES RE-NORMALIZADOS ("PARCIALEZ")
                        // Resolvemos el monto del contrato aquí para asegurar que si se creó en esta sesión, tengamos el dato.
                        let montoTotalContrato = 0;
                        if (idContratoContextoActual) {
                            const cDB = dbSelect('Contratos', { ID_Contrato: idContratoContextoActual });
                            if (cDB && cDB.length > 0) montoTotalContrato = parseFloat(cDB[0].Monto_Total_Con_IVA) || 0;
                        }

                        if (montoTotalContrato > 0) {
                            const pctIA = parseFloat(registro.Avance_Programado_Pct) || 0;
                            const montoIA = parseFloat(registro.Monto_Programado) || 0;

                            if (montoIA > 0) {
                                // Si tenemos monto, el % SIEMPRE se recalcula para ser "Parcial" (vs contrato total)
                                registro.Avance_Programado_Pct = (montoIA / montoTotalContrato) * 100;
                            } else if (pctIA > 0) {
                                // Si solo tenemos %, el usuario indica que suelen ser conceptuales (por fila)
                                // Intentamos obtener el total del concepto para normalizar
                                let montoConceptoTotal = 0;
                                if (registro.ID_Concepto) {
                                    const conDB = dbSelect('Catalogo_Conceptos', { ID_Concepto: registro.ID_Concepto });
                                    if (conDB && conDB.length > 0) montoConceptoTotal = parseFloat(conDB[0].Importe_Total_Sin_IVA) || 0;
                                }

                                if (montoConceptoTotal > 0) {
                                    // Calculamos el monto real del periodo usando el % conceptual
                                    const montoCalculado = (montoConceptoTotal * (pctIA / 100));
                                    registro.Monto_Programado = montoCalculado.toFixed(2);
                                    // Y fijamos el % como parcial vs el contrato
                                    registro.Avance_Programado_Pct = (montoCalculado / montoTotalContrato) * 100;
                                } else {
                                    // Fallback: Si no hay monto de concepto, asumimos que el % ya es parcial
                                    registro.Monto_Programado = (montoTotalContrato * (pctIA / 100)).toFixed(2);
                                    registro.Avance_Programado_Pct = pctIA;
                                }
                            }
                        }

                        // Limpiar temporales
                        delete registro.Clave_Concepto_Temp;
                        delete registro.Periodo_Temp;
                        delete registro.ID_Numero_Programa;
                    }

                    // INYECCIÓN FORZOSA DE LLAVES FORÁNEAS (Basado en RELACIONES_BD)
                    const relacion = typeof RELACIONES_BD !== 'undefined' ? RELACIONES_BD[tabla] : null;

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
                        if (tabla === 'Contratistas' && registro.RFC) {
                            const resC = dbSelect('Contratistas', { RFC: registro.RFC });
                            if (resC && resC.length > 0) match = resC[0];
                        } else if (tabla === 'Convenios_Recurso' && registro.Numero_Acuerdo) {
                            const resConv = dbSelect('Convenios_Recurso', { Numero_Acuerdo: registro.Numero_Acuerdo });
                            if (resConv && resConv.length > 0) match = resConv[0];
                        } else if (tabla === 'Contratos' && registro.Numero_Contrato) {
                            const resCon = dbSelect('Contratos', { Numero_Contrato: registro.Numero_Contrato });
                            if (resCon && resCon.length > 0) match = resCon[0];
                        } else if (tabla === 'Catalogo_Conceptos' && registro.ID_Contrato) {
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
                            // MATCH POR NOMBRE DE PERIODO (Natural Key)
                            const resPer = dbSelect('Programa_Periodo', { ID_Numero_Programa: registro.ID_Numero_Programa, Periodo: registro.Periodo });
                            if (resPer && resPer.length > 0) match = resPer[0];
                        } else if (tabla === 'Matriz_Insumos' && registro.ID_Concepto && registro.Clave_Insumo) {
                            const resMat = dbSelect('Matriz_Insumos', { ID_Concepto: registro.ID_Concepto, Clave_Insumo: registro.Clave_Insumo });
                            if (resMat && resMat.length > 0) match = resMat[0];
                        } else if (tabla === 'Programa_Ejecucion' && registro.ID_Concepto && registro.ID_Programa_Periodo) {
                            // MATCHING DINÁMICO: Basar en Concepto + Periodo
                            const resEjec = dbSelect('Programa_Ejecucion', {
                                ID_Concepto: registro.ID_Concepto,
                                ID_Programa_Periodo: registro.ID_Programa_Periodo
                            });
                            if (resEjec && resEjec.length > 0) {
                                match = resEjec[0];
                                if (resEjec && resEjec.length > 1) {
                                    // Dejar el primero que macha y borrar los demas para evitar duplicados infinitos
                                    for (let k = 1; k < resEjec.length; k++) {
                                        const condDel = {};
                                        condDel[pkName] = resEjec[k][pkName];
                                        dbDelete('Programa_Ejecucion', condDel);
                                    }
                                }
                            }
                        }
                    }

                    // Limpiar campos temporales para no insertarlos en columnas basura
                    delete registro.Clave_Concepto_Temp;
                    delete registro.Descripcion_Temp;
                    delete registro.Periodo_Temp;
                    // NO BORRAR Clave ni Periodo, son parte del esquema
                    // delete registro.Clave; 
                    // delete registro.Periodo;

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

                        // NUEVO: Registrar en el mapa de periodos por nombre (OBJETO COMPLETO PARA FECHAS)
                        if (tabla === 'Programa_Periodo' && registro.Periodo) {
                            mapaPeriodosReales[registro.Periodo] = { ...dataMerged, ID_Programa_Periodo: match[pkName] };
                        }

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

                            // NUEVO: Registrar en el mapa de periodos por nombre (OBJETO COMPLETO PARA FECHAS)
                            if (tabla === 'Programa_Periodo' && registro.Periodo) {
                                mapaPeriodosReales[registro.Periodo] = { ...registro, ID_Programa_Periodo: newId };
                            }
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
