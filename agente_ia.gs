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

REGLA DE ORO (Criterios de Clasificación - PENSAMIENTO VISUAL):

**Paso 1: ¿Es una Cuadrícula/Tabla o es Prosa?**
- Si ves una estructura dominante de **COLUMNAS, FILAS y CELDAS** con números, descripciones técnicas y precios unitarios -> ES UN **APU / MATRIZ**.
- Si ves un documento escrito en **PÁRRAFOS de texto legal** articulado en cláusulas -> ES UN **CONTRATO**.

**Reglas Específicas:**

1. **CAF (Convenio de Apoyo Financiero / Suficiencia)**: 
   - Estructura: Oficio informativo (1-3 págs).
   - Claves: "Oficio de Suficiencia", "BANOBRAS", "Recursos Federales".
   - Identificador: No es una tabla densa de insumos.

2. **APU / MATRIZ / CATÁLOGO (Análisis de Precios Unitarios)**: 
   - **MÁXIMA PRIORIDAD**: Si el documento contiene desgloses de "Materiales", "Mano de Obra" o "Equipo", o si es una tabla de conceptos con Clave, Unidad, Cantidad y P.U.
   - **FILTRO DE ARCHIVO**: Archivos con nombres como "FORMA E5", "E7", "Anexo Técnico", "Catálogo" son casi siempre APUs/Catálogos.
   - **REGLA VISUAL**: Tiene muchas celdas. Se ve como una hoja de cálculo impresa.

3. **CONTRATO (Documento Legal)**: 
   - **IDENTIFICADOR EXCLUSIVO**: Debe contener una narrativa legal con "DECLARACIONES" y "CLÁUSULAS" (Ej: PRIMERA, SEGUNDA...).
   - **REGLA NEGATIVA**: Un contrato legal NUNCA es un listado detallado de insumos fila por fila. Si el documento se parece a un Excel, NO es un contrato.

4. **PROGRAMA (Programa de Obra)**: 
   - Estructura: Tabla con una línea de tiempo (Meses, Semanas) y barras o montos en las celdas para indicar ejecución.

5. **FIANZA**:
   - Estructura: Formato de afianzadora/aseguradora con logotipos institucionales y datos de la póliza.

INSTRUCCIÓN: Analiza visualmente el documento y responde ÚNICAMENTE con un JSON.
Clases permitidas: 'Contratos', 'CAF', 'Programa', 'Matriz_Insumos', 'Fianza'.
Ejemplo: {"clase": "Matriz_Insumos", "razon_visual": "Contiene tablas densas de materiales y precios unitarios"}
`;

/**
 * Clasifica un documento usando IA para identificar si es Contrato, CAF, Programa o Matriz.
 * Implementa un mecanismo de fallback si el modelo principal falla.
 */
function clasificarDocumentoIA(base64Content, mimeType, filename = "documento.pdf") {
    const models = ["gemini-3.1-pro", "gemini-2.5-flash"];
    let lastError = null;

    for (const modelId of models) {
        try {
            const finalPrompt = PROMPT_CLASIFICADOR_DOCUMENTOS.replace('{{FILENAME}}', filename);
            const payload = {
                contents: [{
                    parts: [
                        { text: finalPrompt },
                        { inline_data: { mime_type: mimeType, data: base64Content.split(',')[1] || base64Content } }
                    ]
                }],
                generationConfig: {
                    temperature: 0.1,
                    response_mime_type: "application/json"
                }
            };

            const options = {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true // Para capturar el 404
            };

            const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=` + GEMINI_API_KEY, options);
            const respCode = response.getResponseCode();
            const respText = response.getContentText();

            if (respCode !== 200) {
                console.warn(`Clasificador Falló con ${modelId} (${respCode}): ${respText}`);
                lastError = respText;
                continue; // Probar siguiente modelo
            }

            const result = JSON.parse(respText);
            const resultText = result.candidates[0].content.parts[0].text;
            const cleanedJson = resultText.replace(/```json|```/g, "").trim();
            const jsonResponse = JSON.parse(cleanedJson);

            return { success: true, clase: jsonResponse.clase, razon: jsonResponse.razon_visual };
        } catch (e) {
            lastError = e.toString();
            console.error(`Error crítico en clasificarDocumentoIA con ${modelId}:`, e);
        }
    }

    return { success: false, error: lastError || "No se pudo clasificar con ningún modelo disponible" };
}

/**
 * Clasificador Robusto Multidocumento
 * Escanea el texto y determina el tipo de documento legal/administrativo.
 */
function clasificarDocumentoRobusto(textoPDF) {
  if (!textoPDF) return { tipoDocumento: 'DESCONOCIDO', confianzaPorcentaje: 0, numeroContratoIdentificado: null, esValido: false };
  // Normalización estricta: todo a mayúsculas y espacios uniformes
  const textoNormalizado = String(textoPDF).toUpperCase().replace(/\s+/g, ' ');

  // 1. Matriz de huellas dactilares (Palabras clave únicas por documento)
  const huellas = {
    'CAF': [
      "CONVENIO DE APOYO FINANCIERO",
      "FONADIN",
      "BANOBRAS",
      "UNIDAD DE BANCA DE INVERSIÓN",
      "DELEGADO FIDUCIARIO"
    ],
    'Contratos': [
      "CONTRATO DE SERVICIOS RELACIONADOS CON OBRA PÚBLICA",
      "DIRECCIÓN GENERAL DE DESARROLLO CARRETERO",
      "PLURIANUAL CON CARÁCTER NACIONAL",
      "EL CONTRATISTA",
      "LA DEPENDENCIA"
    ],
    'Matriz_Insumos': [
      "ANALISIS DE PRECIOS UNITARIOS",
      "P. UNITARIO",
      "CARGOS ADICIONALES",
      "MANO DE OBRA GRAVABLE",
      "FORMA E5",
      "INSUMOS QUE INTERVIENEN EN LA INTEGRACIÓN"
    ],
    'Fianza': [
      "POLIZA DE FIANZA",
      "COMISION NACIONAL DE SEGUROS Y FIANZAS",
      "INSTITUCIÓN DE GARANTÍAS",
      "FIADORA",
      "MONTO TOTAL DE LA FIANZA"
    ],
    'Programa': [
      "PROGRAMA DE TRABAJO",
      "DIAS NATURALES",
      "ESTIMACIÓN 1",
      "FECHA DE CONTRATO",
      "PLAZO:"
    ]
  };

  // 2. Variables para rastrear al ganador
  let mejorCategoria = 'DESCONOCIDO';
  let puntajeMaximo = 0;
  let confianzaGanadora = 0;

  // 3. Motor de evaluación
  for (const [categoria, palabrasClave] of Object.entries(huellas)) {
    let puntosObtenidos = 0;
    
    // Contar cuántas palabras clave se encontraron en el texto
    palabrasClave.forEach(patron => {
      if (textoNormalizado.includes(patron)) {
        puntosObtenidos++;
      }
    });

    // Actualizar el ganador si esta categoría tiene más puntos
    if (puntosObtenidos > puntajeMaximo) {
      puntajeMaximo = puntosObtenidos;
      mejorCategoria = categoria;
      confianzaGanadora = (puntosObtenidos / palabrasClave.length) * 100;
    }
  }

  // 4. Umbral de seguridad (Mínimo 2 coincidencias para evitar falsos positivos)
  if (puntajeMaximo < 2) {
    mejorCategoria = 'DESCONOCIDO';
    confianzaGanadora = 0;
  }

  // 5. Extracción temprana de metadatos transversales (si aplican)
  let numContrato = null;
  const matchContrato = textoNormalizado.match(/(?:CONTRATO N[UÚ]MERO|CONTRATO:?)\s*([A-Z0-9\-]+)/);
  if (matchContrato) {
    numContrato = matchContrato[1].trim();
  }

  // 6. Retorno de la decisión
  return {
    tipoDocumento: mejorCategoria,
    confianzaPorcentaje: typeof confianzaGanadora === 'number' ? confianzaGanadora.toFixed(2) : "0.00",
    numeroContratoIdentificado: numContrato,
    esValido: mejorCategoria !== 'DESCONOCIDO'
  };
}

/**
 * Switch de Enrutamiento Maestro Integrado
 * Despacha la extracción a funciones especializadas o, en su defecto, pide ayuda al agente de IA.
 */
function procesarDocumentoEntrante(textoExtraido, base64Data = null, mimeType = null, contextJson = null) {
  const clasificacion = clasificarDocumentoRobusto(textoExtraido);
  console.log("Documento clasificado por matriz de pesos como: " + clasificacion.tipoDocumento + " con " + clasificacion.confianzaPorcentaje + "% de confianza.");

  let extContext = contextJson;
  if (typeof extContext === 'string') {
      try { extContext = JSON.parse(extContext); } catch (e) { extContext = {}; }
  } else if (!extContext) {
      extContext = {};
  }

  if (clasificacion.numeroContratoIdentificado) {
      extContext.Numero_Contrato_Identificado = clasificacion.numeroContratoIdentificado;
  }

  const strContext = Object.keys(extContext).length > 0 ? JSON.stringify(extContext) : null;

  switch (clasificacion.tipoDocumento) {
    case 'CAF':
    case 'Contratos':
    case 'Matriz_Insumos':
    case 'Fianza':
    case 'Programa':
      // Usar nuestro motor IA actual pre-enrutado a la tabla correcta
      return procesarDocumentoConIA(base64Data, mimeType, clasificacion.tipoDocumento, strContext);
      
    default:
      // Derivar al agente_ia.gs al flujo libre inferido
      return procesarDocumentoConIA(base64Data, mimeType, null, "Intenta determinar qué tipo de documento es y extrae sus datos principales. " + (strContext || ""));
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
5. **GUARDIA DE SEGURIDAD (CRÍTICO)**: Si detectas que el documento contiene desgloses exhaustivos de INSUMOS por concepto (tablas de materiales, equipo y mano de obra), entonces NO ES UN CONTRATO LEGAL. En ese caso, devuelve {"tipo_documento": "ERROR_ES_APU"}.

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
Rol: Experto en Ingeniería de Costos y Extracción de Datos JSON.
Objetivo: Extraer TODOS los conceptos y desgloses de APU (Precios Unitarios) del documento.

REGLAS DE ORO (ESTRICTAS):
1. **SALIDA ÚNICA**: Responde EXCLUSIVAMENTE con un JSON válido. Sin texto explicativo.
2. **JERARQUÍA**: Cada concepto DEBE tener su array de 'insumos'. Es el corazón del dato.
3. **TABLAS LARGAS**: Si un concepto continúa en varias páginas, acumula sus insumos sin cerrar el objeto hasta terminarlo.
4. **COSTO DIRECTO**: Extrae el subtotal (antes de indirectos) para cada concepto. Si no está explícito, suma los importes de sus insumos.
5. **JSON COMPLETO**: Asegúrate de cerrar todas las llaves y corchetes, incluso si el documento es muy extenso.
6. **TIPO INSUMO**: El tipo de insumo es un número (1, 2 o 3) que indica el tipo de insumo (Material, Mano de Obra, Equipo).
7. **GUARDIA DE SEGURIDAD (CRÍTICO)**: Si el documento es meramente un texto legal de cláusulas y NO contiene tablas de desgloses de precios unitarios con insumos, devuelve {"tipo_documento": "ERROR_ES_CONTRATO"}.

Estructura JSON:
{
  "tipo_documento": "APU",
  "datos_extraidos": {
    "conceptos": [
      {
        "Clave_Concepto": "Clave (ej. N3-25-1)",
        "Descripcion_Concepto": "Descripción técnica",
        "Unidad": "Unidad (ej. ESTUDIO, MES, LOTE)",
        "Costo_Directo": 0.00,
        "Precio_Unitario_Total": 0.00,
        "insumos": [
          {
            "Tipo_Insumo": "1 | 2 | 3",
            "Clave_Insumo": "Clave",
            "Descripcion": "Nombre del insumo",
            "Unidad": "JOR, Hr, Pza, kg",
            "Costo_Unitario": 0.00,
            "Rendimiento_Cantidad": 0.00,
            "Importe": 0.00,
            "Porcentaje_Incidencia": 0.00
          }
        ]
      }
    ]
  }
}
`;

const PROMPT_IMPORTACION_CAF = `
Rol y Objetivo:
Eres un Experto en Gestión de Recursos. Tu misión es extraer datos de CAF (Convenio de Apoyo Financiero / Suficiencia Presupuestal).

PASO: EXTRACCIÓN Y SALIDA JSON
Devuelve ÚNICAMENTE un objeto JSON válido.

Estructura requerida:
{
  "tipo_documento": "CAF",
  "nivel_confianza": "Alto",
  "datos_extraidos": {
    "Numero_Acuerdo": "Folio o número de acuerdo del convenio (ej. CT/1A-ORD/11-ABRIL-25/X)",
    "Nombre_Fondo": "Nombre COMPLETO y formal del fondo (ej. FONDO NACIONAL DE INFRAESTRUCTURA (FONADIN))",
    "Monto_Apoyo": 250000000.0,
    "Fecha_Firma": "YYYY-MM-DD",
    "Vigencia_Fin": "Fecha de vencimiento o fin de vigencia (YYYY-MM-DD)",
    "Objeto_Programa": "Descripción del objeto o programa",
    "Estado": "STATUS DEL CAF: Escribe 'VIGENTE' o 'VENCIDO' (según la fecha de hoy vs Vigencia_Fin)",
    "Entidades_Involucradas": ["BANOBRAS", "SHCP", "..."]
  }
}

REGLAS CRÍTICAS:
1. **Numero_Acuerdo**: Es el identificador principal. Ignora guiones o caracteres extra si son claramente errores de formato.
2. **Nombre_Fondo**: NO uses abreviaturas solas. Busca el nombre institucional completo.
3. **Estado**: En este documento específico, este campo DEBE contener el estatus operativo ('VIGENTE' o 'VENCIDO').
4. **Vigencia**: Busca palabras como "Vigencia hasta", "Vencimiento", "Fecha final".
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
function procesarDocumentoConIA(base64Data, mimeType, targetTable = null, contextJson = null) {
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

        // 2. Seleccionar el Modelo y Fallbacks (Optimización de Resiliencia y Costos)
        const modelsToTry = (function () {
            // Alta Complejidad (Tablas y Razonamiento Espacial)
            if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') return ["gemini-3.1-pro", "gemini-2.5-pro"];
            if (targetTable === 'Programa' || targetTable === 'Programa_Ejecucion') return ["gemini-3.1-pro", "gemini-2.5-pro"];

            // Media Complejidad (Prosa y Extracción Mixta)
            if (targetTable === 'Convenios_Recurso' || targetTable === 'CAF') return ["gemini-3.0-flash", "gemini-2.5-flash"];
            if (targetTable === 'Contratos') return ["gemini-3.0-flash", "gemini-2.5-flash"];
            if (targetTable === 'Catalogo_Conceptos' || targetTable === 'Catalogo') return ["gemini-3.0-flash", "gemini-2.5-flash"];
            if (targetTable === 'Estimaciones' || targetTable === 'Facturas') return ["gemini-3.0-flash", "gemini-2.5-flash"];

            // Baja Complejidad (Texto legal plano)
            if (targetTable === 'Fianza') return ["gemini-3.1-flash-lite", "gemini-2.5-flash-lite"];

            // Default Fallback
            return ["gemini-3-flash", "gemini-2.5-flash"];
        })();

        let systemPrompt = (function () {
            if (targetTable === 'Matriz_Insumos' || targetTable === 'Análisis_P_U') return PROMPT_IMPORTACION_MATRICES;
            if (targetTable === 'Estimaciones' || targetTable === 'Facturas') return PROMPT_VALIDACION_ESTIMACIONES;
            if (targetTable === 'Convenios_Recurso' || targetTable === 'CAF') return PROMPT_IMPORTACION_CAF;
            if (targetTable === 'Catalogo_Conceptos' || targetTable === 'Catalogo') return PROMPT_IMPORTACION_CATALOGO;
            if (targetTable === 'Contratos') return PROMPT_IMPORTACION_CONTRATOS;
            if (targetTable === 'Programa' || targetTable === 'Programa_Ejecucion') return PROMPT_IMPORTACION_PROGRAMA;
            if (targetTable === 'Fianza') return PROMPT_VALIDACION_FIANZA;
            return PROMPT_IMPORTACION_CONTRATOS;
        })();


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

        const partsArr = [
            { text: "INSTRUCCIONES_SISTEMA: " + systemPrompt },
            { text: "CONTEXTO_ADICIONAL: " + promptAdicional }
        ];

        if (b64) {
            partsArr.push({ inline_data: { mime_type: mimeType || 'application/pdf', data: b64 } });
        }

        const payload = {
            contents: [{ parts: partsArr }],
            generationConfig: {
                temperature: 0.1,
                response_mime_type: "application/json"
            }
        };

        const options = {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        let llmString = null;
        let lastError = null;

        for (const modelId of modelsToTry) {
            try {
                const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${GEMINI_API_KEY}`;
                const response = UrlFetchApp.fetch(url, options);
                const resultJson = JSON.parse(response.getContentText());

                if (response.getResponseCode() !== 200) {
                    console.warn(`Extracción falló con ${modelId} (${response.getResponseCode()}): ${response.getContentText()}`);
                    lastError = resultJson.error?.message || response.getContentText();
                    continue;
                }

                llmString = resultJson.candidates[0].content.parts[0].text;
                if (llmString) break; // Éxito
            } catch (err) {
                lastError = err.toString();
                console.error(`Error de red con ${modelId}:`, err);
            }
        }

        if (!llmString) {
            throw new Error(`No se pudo extraer información del documento después de varios intentos. Último error: ${lastError}`);
        }

        // --- LOG PROVISIONAL DE EXTRACCIÓN ---
        console.log("RAW_AI_EXTRACTION [" + (targetTable || "GENERAL") + "]:", llmString);

        let extraction;
        try {
            extraction = JSON.parse(limpiarJSONIA(llmString));
        } catch (errJson) {
            console.error("Error parsing AI JSON:", errJson, "Raw text:", llmString);
            // Intento desesperado: Si está truncado pero el inicio es válido, intentar extraer lo que se pueda regex
            throw new Error("El documento es demasiado complejo y la respuesta de la IA contiene errores de formato (JSON Malformed). Por favor, intenta procesar menos páginas a la vez.");
        }
        const datosFinales = extraction.datos_extraidos || extraction.datos || extraction;

        // --- VALIDACIÓN DE GUARDIA DE SEGURIDAD (EXTRACCIÓN) ---
        const tipoDet = (extraction.tipo_documento || "").toString().toUpperCase();
        if (tipoDet.startsWith('ERROR_')) {
            const errorMsg = tipoDet === 'ERROR_ES_APU'
                ? "CRÍTICO: El documento parece ser un APU/MATRIZ (Celdas/Columnas) pero se intentó procesar como CONTRATO (Legal)."
                : "CRÍTICO: El documento parece ser un CONTRATO (Legal) pero se intentó procesar como APU/MATRIZ (Celdas/Columnas).";
            return { success: false, error: errorMsg };
        }

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
        console.error("Error en procesarDocumentoConIA:", e);
        return generarRespuesta(false, e.toString(), 'procesarDocumentoConIA');
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

    // CACHÉ LOCAL: Evita duplicados si Sheets tiene latencia o si el JSON trae repetidos
    const cacheLocal = {
        'Contratistas': [],
        'Programa': [],
        'Programa_Periodo': [],
        'Catalogo_Conceptos': []
    };

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

    // --- RESOLUCIÓN GLOBAL DE CONTRATO (DIAGNÓSTICO PREVENTIVO) ---
    let pContext = respuestaIA.parentContext;
    if (typeof pContext === 'string') {
        try { pContext = JSON.parse(pContext); } catch (e) { pContext = null; }
    }

    let globalIdContrato = respuestaIA.ID_Contrato_Contexto ||
        (pContext ? (pContext.ID_Contrato || pContext.idContrato || pContext.id_contrato) : null);

    if (!globalIdContrato) {
        const rootKeys = Object.keys(datos);
        const rootIdKey = rootKeys.find(k => k.toLowerCase() === 'id_contrato' || k.toLowerCase() === 'idcontrato');
        if (rootIdKey) globalIdContrato = datos[rootIdKey];
    }

    if (!globalIdContrato) {
        const lookIn = ['Contratos', 'Catalogo_Conceptos', 'Programa', 'Estimaciones'];
        for (const t of lookIn) {
            if (datos[t] && Array.isArray(datos[t]) && datos[t][0]) {
                const itemKeys = Object.keys(datos[t][0]);
                const itemIdKey = itemKeys.find(k => k.toLowerCase() === 'id_contrato' || k.toLowerCase() === 'idcontrato');
                if (itemIdKey) {
                    globalIdContrato = datos[t][0][itemIdKey];
                    break;
                }
            }
        }
    }

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
            Estado: c.Estado || 'Ciudad de México', // Ahora es ubicación geográfica
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
        if (contratista.Razon_Social || contratista.RFC) {
            datos.Contratistas = [contratista];
        }
        if (contrato.Numero_Contrato || contrato.ID_Contrato) {
            datos.Contratos = [contrato];
        }
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
            datos.Catalogo_Conceptos = datos.conceptos.map(c => ({
                ...c,
                ID_Contrato: globalIdContrato
            }));
        }
    } else if (tipo === 'PROGRAMA') {
        const nombreProg = datos.Nombre_Programa || datos.Programa || "Programa General de Obra";

        datos.Programa = [{
            ID_Contrato: globalIdContrato,
            Programa: nombreProg,
            Tipo_Programa: datos.Tipo_Programa || "MENSUAL",
            Fecha_Inicio: datos.Fechas_Generales?.Fecha_Inicio || null,
            Fecha_Termino: datos.Fechas_Generales?.Fecha_Fin || null,
            Plazo_Ejecucion_Dias: datos.Fechas_Generales?.Plazo_Dias || null
        }];

        if (datos.Periodos_Programados) {
            datos.Programa_Periodo = [];
            datos.Programa_Ejecucion = [];
            const conceptosUnicos = {};

            datos.Periodos_Programados.forEach((per, idx) => {
                const nombrePer = String(per.Nombre_Periodo || "").trim().toUpperCase();

                // 1. Guardar Periodo (Linkage will happen in the loop via ultimosIdsInsertados or natural match)
                const perObj = {
                    ID_Numero_Programa_TmpLink: nombreProg, // Temporary link to find parent by name in loop
                    Numero_Periodo: idx + 1,
                    Periodo: nombrePer,
                    Fecha_Inicio: per.Fecha_Inicio,
                    Fecha_Termino: per.Fecha_Termino || per.Fecha_Fin
                };
                datos.Programa_Periodo.push(perObj);

                // 2. Procesar Conceptos del Periodo
                if (per.Conceptos && Array.isArray(per.Conceptos)) {
                    per.Conceptos.forEach(cp => {
                        if (cp.Clave_Concepto && !conceptosUnicos[cp.Clave_Concepto]) {
                            conceptosUnicos[cp.Clave_Concepto] = {
                                Clave: cp.Clave_Concepto,
                                Descripcion: cp.Descripcion_Contextual || cp.Descripcion_Limpiada || "N/A",
                                Unidad: cp.Unidad || "SRV",
                                Cantidad_Contratada: cp.Cantidad || 1,
                                Precio_Unitario: cp.Precio_Unitario || null,
                                Importe_Total_Sin_IVA: cp.Importe_Total || 0,
                                ID_Contrato: globalIdContrato
                            };
                        }

                        const monto = parseFloat(cp.Monto_Periodo) || 0;
                        const pct = parseFloat(cp.Porcentaje_Avance) || 0;
                        if (monto === 0 && pct === 0) return;

                        datos.Programa_Ejecucion.push({
                            Programa_TmpLink: nombreProg,
                            Periodo_TmpLink: nombrePer,
                            Clave_Concepto_Temp: cp.Clave_Concepto,
                            Descripcion_Temp: cp.Descripcion_Contextual || cp.Descripcion_Limpiada,
                            Periodo_Temp: nombrePer,
                            Fecha_Inicio: per.Fecha_Inicio,
                            Fecha_Fin: per.Fecha_Termino || per.Fecha_Fin,
                            Monto_Programado: cp.Monto_Periodo,
                            Avance_Programado_Pct: cp.Porcentaje_Avance
                        });
                    });
                }
            });
            datos.Catalogo_Conceptos = Object.values(conceptosUnicos);

        } else if (datos.Periodos_Identificados && datos.conceptos_programados) {
            // LOGICA VIEJA
            datos.Programa_Periodo = datos.Periodos_Identificados.map((p, i) => {
                const esObjeto = typeof p === 'object';
                return {
                    ID_Numero_Programa_TmpLink: nombreProg,
                    Numero_Periodo: i + 1,
                    Periodo: (esObjeto ? p.Nombre : p).trim().toUpperCase(),
                    Fecha_Inicio: esObjeto ? p.Fecha_Inicio : null,
                    Fecha_Termino: esObjeto ? p.Fecha_Termino : null
                };
            });
            datos.Catalogo_Conceptos = datos.conceptos_programados.map(cp => ({
                Clave: cp.Clave_Concepto,
                Descripcion: cp.Descripcion_Limpiada,
                Unidad: cp.Unidad,
                Cantidad_Contratada: cp.Cantidad || 1,
                Precio_Unitario: cp.Precio_Unitario || null,
                Importe_Total_Sin_IVA: cp.Importe_Total,
                ID_Contrato: globalIdContrato
            }));
            datos.Programa_Ejecucion = [];
            datos.conceptos_programados.forEach(cp => {
                cp.avances_por_periodo?.forEach((av) => {
                    const monto = parseFloat(av.Monto_Periodo) || 0; const pct = parseFloat(av.Porcentaje_Avance) || 0; if (monto === 0 && pct === 0) return;
                    let nombrePer = String(av.Nombre_Periodo || "").trim().toUpperCase();
                    datos.Programa_Ejecucion.push({
                        Programa_TmpLink: nombreProg,
                        Periodo_TmpLink: nombrePer,
                        Clave_Concepto_Temp: cp.Clave_Concepto,
                        Periodo_Temp: nombrePer,
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
                ID_Contrato: globalIdContrato,
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


    // Si aún no hay ID global pero la IA extrajo un Numero_Contrato, buscarlo o crearlo
    let needleText = datos.Numero_Contrato ||
        (datos.Contratos && datos.Contratos[0] ? datos.Contratos[0].Numero_Contrato : null) ||
        (datos.Contrato && typeof datos.Contrato === 'string' ? datos.Contrato : null);

    // Normalizar needleText (Quitar espacios extra)
    if (needleText) needleText = needleText.toString().trim();

    if (!globalIdContrato && needleText && needleText.length > 0) {
        if (!isNaN(parseInt(needleText)) && String(needleText).length < 6) {
            globalIdContrato = parseInt(needleText);
        } else {
            const contDb = dbSelect('Contratos', { Numero_Contrato: needleText });
            if (contDb && contDb.length > 0) {
                globalIdContrato = contDb[0].ID_Contrato;
            } else if (tipo !== 'CONTRATO' && !globalIdContrato) {
                // Solo auto-crear si:
                // 1. El documento NO es un contrato (es CAF, etc)
                // 2. NO tenemos un ID global ya resuelto (protección extra contra duplicados)
                // 3. Hay un Numero_Contrato válido
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
    const mapaConceptosReales = {}; // NUEVO: Para vincular conceptos por clave sin usar IDs temporales
    if (globalIdContrato) {
        ultimosIdsInsertados['Contratos'] = globalIdContrato;
        // Pre-cargar catálogos para resolución instantánea (Evita dbSelect masivos)
        const catExistente = dbSelect('Catalogo_Conceptos', { ID_Contrato: globalIdContrato });
        // CORRECCIÓN: Agregar .replace(/\\s+/g, '') para evitar falsos negativos por espacios en blanco extraños del PDF o IA
        if (catExistente) catExistente.forEach(c => { if (c.Clave) mapaConceptosReales[c.Clave.toString().replace(/\\s+/g, '').toUpperCase()] = c.ID_Concepto; });

        // Pre-cargar periodos si existen programas
        const progExistente = dbSelect('Programa', { ID_Contrato: globalIdContrato });
        if (progExistente) progExistente.forEach(p => {
            const persExistentes = dbSelect('Programa_Periodo', { ID_Numero_Programa: p.ID_Numero_Programa });
            if (persExistentes) persExistentes.forEach(per => { if (per.Periodo) mapaPeriodosReales[per.Periodo.toString().toUpperCase()] = { ...per }; });
        });
    }

    jerarquia.forEach(tabla => {
        if (datos[tabla] && Array.isArray(datos[tabla])) {
            const headers = ESQUEMA_BD[tabla];
            const pkName = headers[0];

            // Cuando se van a guardar registros de Matriz_Insumos
            if (tabla === 'Matriz_Insumos') {
                datos[tabla] = datos[tabla].map(insumo => {
                    // Si la IA mandó una Clave_Concepto_Temp pero no el ID numérico
                    const claveTemp = insumo.Clave_Concepto_Temp || insumo.Clave_Concepto_Padre || insumo.Clave_Padre;
                    if (!insumo.ID_Concepto && claveTemp) {
                        // Buscar el ID en el mapa de conceptos en caché
                        const claveStr = String(claveTemp).replace(/\\s+/g, '').toUpperCase();
                        if (mapaConceptosReales[claveStr]) {
                            insumo.ID_Concepto = mapaConceptosReales[claveStr];
                        }
                    }
                    return insumo;
                });
            }

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
            else if (tabla === 'Convenios_Modificatorios') skName = 'Numero_Convenio_Mod';
            else if (tabla === 'Catalogo_Conceptos') skName = 'Clave';
            else if (tabla === 'Programa') skName = 'Tipo_Programa';

            let conceptCounter = 0; // Contador secuencial para ordenamiento visual

            datos[tabla].forEach(registro => {
                try {
                    // --- ORDEN SECUENCIAL PARA CONCEPTOS ---
                    if (tabla === 'Catalogo_Conceptos') {
                        conceptCounter++;
                        registro.Orden = conceptCounter;
                    }

                    // --- SANITIZACIÓN DE CAMPOS (MAPPING FALLBACK) ---
                    if (tabla === 'Catalogo_Conceptos') {
                        if (registro.Clave_Concepto && !registro.Clave) registro.Clave = registro.Clave_Concepto;
                        if (registro.Descripcion_Concepto && !registro.Descripcion) registro.Descripcion = registro.Descripcion_Concepto;
                        if (registro.Unidad_Medida && !registro.Unidad) registro.Unidad = registro.Unidad_Medida;

                        // Antes de decidir si es nuevo o actualización
                        const claveBusqueda = (registro.Clave_Concepto_Temp || registro.Clave || "").toString().replace(/\\s+/g, '').toUpperCase();
                        const idExistente = mapaConceptosReales[claveBusqueda];
                        if (idExistente) {
                            // Forzamos el ID para que el sistema sepa que es un UPDATE (UPSERT)
                            registro.ID_Concepto = idExistente;
                        }
                    }

                    Object.keys(registro).forEach(k => {
                        const lowK = k.toLowerCase();
                        if ((lowK.startsWith('id_') || lowK.startsWith('idcont') || lowK.startsWith('idconcep')) && k !== pkName && registro[k]) {
                            // Si es un temporal, buscar en mapa
                            if (mapaIds[registro[k]]) {
                                registro[k] = mapaIds[registro[k]];
                            }
                            // ALIASING: Si el campo es una variante de ID_Contrato, normalizar la KEY para la DB
                            if (lowK === 'id_contrato' || lowK === 'idcontrato') {
                                if (k !== 'ID_Contrato') {
                                    registro['ID_Contrato'] = registro[k];
                                    // No borramos la original por si la IA la necesita, pero la DB usará ID_Contrato
                                }
                            }
                        }
                    });

                    // --- RESOLUCIÓN DE LINKS JERÁRQUICOS (REAL ID APPROACH) ---
                    const idContratoContextoActual = resultados.ids['Contratos'] || globalIdContrato;

                    if (tabla === 'Programa_Periodo' && !registro.ID_Numero_Programa && registro.ID_Numero_Programa_TmpLink) {
                        const progName = registro.ID_Numero_Programa_TmpLink;
                        // Intentar match con lo recién insertado/seleccionado por nombre
                        const progMatch = dbSelect('Programa', { ID_Contrato: idContratoContextoActual, Programa: progName });
                        if (progMatch && progMatch.length > 0) {
                            registro.ID_Numero_Programa = progMatch[0].ID_Numero_Programa;
                        } else if (ultimosIdsInsertados['Programa']) {
                            registro.ID_Numero_Programa = ultimosIdsInsertados['Programa'];
                        }
                    }

                    if (tabla === 'Programa_Ejecucion') {
                        // 1. Resolver ID_Programa
                        if (!registro.ID_Programa && registro.Programa_TmpLink) {
                            const progMatch = dbSelect('Programa', { ID_Contrato: idContratoContextoActual, Programa: registro.Programa_TmpLink });
                            if (progMatch && progMatch.length > 0) registro.ID_Programa = progMatch[0].ID_Numero_Programa;
                            else registro.ID_Programa = ultimosIdsInsertados['Programa'];
                        }
                        // 2. Resolver ID_Programa_Periodo
                        if (!registro.ID_Programa_Periodo && (registro.Periodo_TmpLink || registro.Periodo_Temp)) {
                            const perName = registro.Periodo_TmpLink || registro.Periodo_Temp;
                            const idProg = registro.ID_Programa || ultimosIdsInsertados['Programa'];
                            if (idProg) {
                                // Buscar periodo por nombre bajo el programa resuelto
                                const perMatch = dbSelect('Programa_Periodo', { ID_Numero_Programa: idProg }).find(p => normalizeText(p.Periodo) === normalizeText(perName));
                                if (perMatch) registro.ID_Programa_Periodo = perMatch.ID_Programa_Periodo;
                            }
                        }
                    }

                    if (tabla === 'Matriz_Insumos') {
                        // Cruzar para obtener el ID_Concepto real usando el mapa de sesión (Súper rápido)
                        if (!registro.ID_Concepto && (registro.Clave_Concepto_Padre || registro.Clave_Padre)) {
                            const clavePadre = (registro.Clave_Concepto_Padre || registro.Clave_Padre).toString().replace(/\\s+/g, '').toUpperCase();
                            if (mapaConceptosReales[clavePadre]) {
                                registro.ID_Concepto = mapaConceptosReales[clavePadre];
                            } else if (idContratoContextoActual) {
                                // Fallback: Búsqueda de último recurso si no está en mapa de sesión
                                const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: idContratoContextoActual });
                                const conMatch = catalogo.find(c => String(c.Clave).replace(/\\s+/g, '').toUpperCase() === clavePadre);
                                if (conMatch) registro.ID_Concepto = conMatch.ID_Concepto;
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
                            const clave = (registro.Clave_Concepto_Temp || registro.Clave).toString().replace(/\\s+/g, '').toUpperCase();
                            if (mapaConceptosReales[clave]) {
                                registro.ID_Concepto = mapaConceptosReales[clave];
                            } else if (idContratoContextoActual) {
                                // Intento 1: Por Clave (Insensible)
                                const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: idContratoContextoActual });
                                let matchClave = catalogo.find(c => String(c.Clave).replace(/\\s+/g, '').toUpperCase() === clave);

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
                        const relaciones = Array.isArray(relacion) ? relacion : [relacion];
                        relaciones.forEach(rel => {
                            let tablaPadre = rel.padre;
                            let llaveForanea = rel.fk;

                            // Casos especiales de inyección:
                            // 1. Si la tabla es hija de Contratos, SIEMPRE forzar el ID_Contrato correcto
                            if (tablaPadre === 'Contratos' && idContratoContextoActual) {
                                registro[llaveForanea] = idContratoContextoActual;
                            }
                            // 2. Si venimos de un padre que acabamos de insertar/mapear
                            else if (ultimosIdsInsertados[tablaPadre]) {
                                if (!registro[llaveForanea]) {
                                    registro[llaveForanea] = ultimosIdsInsertados[tablaPadre];
                                }
                            }
                        });
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
                        if (tabla === 'Contratistas') {
                            // --- DOBLE VALIDACIÓN REFORZADA (RFC + Razón Social + Caché Local) ---
                            const rfcNormLimpio = normalizeText(registro.RFC || "");
                            const razonSocialNorm = normalizeText(registro.Razon_Social || "");

                            // 1. Buscar en Caché Local primero (Súper rápido y evita latencia de Sheets)
                            match = (cacheLocal.Contratistas || []).find(c => {
                                const rfcDB = normalizeText(c.RFC || "");
                                const matchRFC = rfcNormLimpio && (rfcDB.includes(rfcNormLimpio) || rfcNormLimpio.includes(rfcDB));
                                const matchNombre = razonSocialNorm && normalizeText(c.Razon_Social || c.Nombre_Contratista) === razonSocialNorm;
                                return matchRFC || matchNombre;
                            });

                            // 2. Si no está en caché, buscar en DB
                            if (!match) {
                                const contratistasDB = dbSelect('Contratistas') || [];

                                // 2.1 Validar por RFC (Normalizado e inclusivo para participaciones conjuntas)
                                if (rfcNormLimpio) {
                                    match = contratistasDB.find(c => {
                                        const rfcDB = normalizeText(c.RFC || "");
                                        return rfcDB.includes(rfcNormLimpio) || rfcNormLimpio.includes(rfcDB);
                                    });
                                }

                                // 2.2 Validar por Razón Social (Fallback)
                                if (!match && razonSocialNorm) {
                                    match = contratistasDB.find(c => {
                                        const nombreDB = normalizeText(c.Razon_Social || c.Nombre_Contratista);
                                        return nombreDB === razonSocialNorm;
                                    });
                                }
                            }

                            if (match) {
                                console.log(`SmartMatch contratista encontrado: ${match.Razon_Social} (ID: ${match.ID_Contratista})`);
                            }
                        } else if (tabla === 'Convenios_Recurso' && registro.Numero_Acuerdo) {
                            const resConv = dbSelect('Convenios_Recurso');
                            const numNorm = normalizeText(registro.Numero_Acuerdo);
                            match = (resConv || []).find(c => normalizeText(c.Numero_Acuerdo) === numNorm);
                        } else if (tabla === 'Contratos' && (registro.Numero_Contrato || globalIdContrato)) {
                            // Intentar match con ID Global primero si existe
                            if (globalIdContrato) {
                                const resCon = dbSelect('Contratos', { ID_Contrato: globalIdContrato });
                                if (resCon && resCon.length > 0) match = resCon[0];
                            }
                            // Si no hay match por ID, intentar por Numero_Contrato
                            if (!match && registro.Numero_Contrato) {
                                const resCon = dbSelect('Contratos', { Numero_Contrato: registro.Numero_Contrato });
                                if (resCon && resCon.length > 0) match = resCon[0];
                            }
                        } else if (tabla === 'Catalogo_Conceptos' && registro.ID_Contrato) {
                            const descNorm = normalizeText(registro.Descripcion || registro.Concepto || "");

                            // 1. Match en Caché Local
                            match = (cacheLocal.Catalogo_Conceptos || []).find(c => {
                                const matchClave = registro.Clave && String(c.Clave).replace(/\\s+/g, '').toUpperCase() === String(registro.Clave).replace(/\\s+/g, '').toUpperCase();
                                const matchDesc = descNorm && normalizeText(c.Descripcion) === descNorm;
                                return c.ID_Contrato === registro.ID_Contrato && (matchClave || matchDesc);
                            });

                            if (!match) {
                                // Match por Clave primero
                                if (registro.Clave) {
                                    const resClave = dbSelect('Catalogo_Conceptos', { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato });
                                    if (resClave && resClave.length > 0) match = resClave[0];
                                }

                                // Match por Descripción como fallback
                                if (!match && descNorm) {
                                    const catalogo = dbSelect('Catalogo_Conceptos', { ID_Contrato: registro.ID_Contrato });
                                    match = catalogo.find(c => normalizeText(c.Descripcion) === descNorm);
                                }
                            }

                            // NUEVO: RECUPERACIÓN GENÉRICA DE HUÉRFANOS (Natural Key match con ID_Contrato vacío)
                            if (!match && skName && registro[skName]) {
                                const rel = typeof RELACIONES_BD !== 'undefined' ? RELACIONES_BD[tabla] : null;
                                const rels = Array.isArray(rel) ? rel : [rel];
                                // Solo aplica si es hijo directo de Contratos con FK ID_Contrato
                                if (rels.some(r => r && r.padre === 'Contratos' && r.fk === 'ID_Contrato')) {
                                    const resHuertano = dbSelect(tabla, { [skName]: registro[skName], ID_Contrato: "" });
                                    if (resHuertano && resHuertano.length > 0) {
                                        match = resHuertano[0];
                                        console.log(`SmartMatch: Reclamando registro huérfano en ${tabla} por ${skName}=${registro[skName]}`);
                                    }
                                }
                            }
                        } else if (tabla === 'Programa' && registro.ID_Contrato) {
                            // SmartMatch reforzado: ID_Contrato + Tipo_Programa + Nombre del Programa
                            match = (cacheLocal.Programa || []).find(p =>
                                p.ID_Contrato === registro.ID_Contrato &&
                                (!registro.Tipo_Programa || p.Tipo_Programa === registro.Tipo_Programa) &&
                                (!registro.Programa || p.Programa === registro.Programa)
                            );

                            if (!match) {
                                const condProg = { ID_Contrato: registro.ID_Contrato };
                                if (registro.Tipo_Programa) condProg.Tipo_Programa = registro.Tipo_Programa;
                                if (registro.Programa) condProg.Programa = registro.Programa;
                                const resProg = dbSelect('Programa', condProg);
                                if (resProg && resProg.length > 0) match = resProg[0];
                            }
                        } else if (tabla === 'Programa_Periodo' && registro.ID_Numero_Programa && registro.Periodo) {
                            // SmartMatch reforzado: El periodo es único dentro de su programa padre
                            const perLabelMatch = registro.Periodo.trim().toUpperCase();
                            const normLabel = normalizeText(perLabelMatch);

                            // 1. Match en Caché Local (Filtrando por el programa padre específico)
                            match = (cacheLocal.Programa_Periodo || []).find(p =>
                                p.ID_Numero_Programa === registro.ID_Numero_Programa &&
                                normalizeText(p.Periodo) === normLabel
                            );

                            if (!match) {
                                const periodosDB = dbSelect('Programa_Periodo', { ID_Numero_Programa: registro.ID_Numero_Programa });
                                match = periodosDB.find(p => normalizeText(p.Periodo) === normLabel);
                            }

                            if (match) {
                                // Forzar el nombre a la versión normalizada para evitar diferencias en la DB
                                registro.Periodo = perLabelMatch;
                            }
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
                    delete registro.ID_Numero_Programa_TmpLink;
                    delete registro.Programa_TmpLink;
                    delete registro.Periodo_TmpLink;
                    delete registro.ID_Contrato_Tmp;

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
                        
                        // Si estamos procesando APU/Matriz y encontramos el concepto del Programa,
                        // DEBEMOS permitir que la APU actualice los precios y descripciones para darles precisión exacta.
                        const esActualizacionDesdeAPU = (tabla === 'Catalogo_Conceptos' && (tipo === 'MATRIZ_INSUMOS' || tipo === 'APU'));
                        
                        // Protegemos la clave, pero si viene de APU, dejamos que actualice Costo, Precio y Unidad.
                        const camposProtegidos = esActualizacionDesdeAPU 
                            ? ['Clave'] // Solo protegemos la clave, dejamos que actualice lo financiero
                            : ['Clave', 'Descripcion', 'Unidad']; 

                        Object.keys(registro).forEach(k => {
                            if (registro[k] !== undefined && registro[k] !== null && registro[k] !== '') {
                                // Si es actualización de APU, damos prioridad a los datos entrantes (registro) sobre los existentes (match) en campos financieros
                                if (esActualizacionDesdeAPU && ['Costo_Directo', 'Precio_Unitario', 'Importe_Total_Sin_IVA', 'Descripcion'].includes(k)) {
                                     dataMerged[k] = registro[k];
                                } else if (!camposProtegidos.includes(k) || !match[k]) {
                                    dataMerged[k] = registro[k];
                                }
                            }
                        });

                        // Verificación de la "Cascada de Importes"
                        if (esActualizacionDesdeAPU && tabla === 'Catalogo_Conceptos') {
                            const pu = parseFloat(dataMerged.Precio_Unitario) || 0;
                            const cant = parseFloat(dataMerged.Cantidad_Contratada) || 0;
                            dataMerged.Importe_Total_Sin_IVA = (pu * cant);
                        }

                        dbUpdate(tabla, dataMerged, { [pkName]: match[pkName] });
                        resultados.actualizados++;
                        resultados.ids[tabla] = match[pkName];
                        if (idOriginalDadoPorIA) mapaIds[idOriginalDadoPorIA] = match[pkName];
                        ultimosIdsInsertados[tabla] = match[pkName];

                        // Acualizar mapa de sesión para resolución instantánea de hijos (APUs/Programa)
                        if (tabla === 'Catalogo_Conceptos' && dataMerged.Clave) {
                            mapaConceptosReales[dataMerged.Clave.toString().replace(/\\s+/g, '').toUpperCase()] = match[pkName];
                        }

                        // Sincronizar contexto global si es un contrato
                        if (tabla === 'Contratos') {
                            globalIdContrato = match[pkName];
                            ultimosIdsInsertados['Contratos'] = globalIdContrato;
                        }

                        // NUEVO: VINCULAR CAF AL CONTRATO (Post-Update)
                        if (tabla === 'Convenios_Recurso' && globalIdContrato) {
                            dbUpdate('Contratos', { ID_Convenio_Vinculado: match[pkName] }, { ID_Contrato: globalIdContrato });
                        }

                        // Actualizar mapa de sesión para resolución instantánea
                        if (tabla === 'Programa_Periodo' && dataMerged.Periodo) {
                            mapaPeriodosReales[dataMerged.Periodo.toString().toUpperCase()] = { ...dataMerged };
                        }

                        // Actualizar Caché Local
                        if (cacheLocal[tabla]) {
                            const idx = cacheLocal[tabla].findIndex(item => item[pkName] === match[pkName]);
                            if (idx !== -1) cacheLocal[tabla][idx] = { ...dataMerged };
                            else cacheLocal[tabla].push({ ...dataMerged });
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

                            // Actualizar mapa de sesión para resolución instantánea de hijos (APUs/Programa)
                            if (tabla === 'Catalogo_Conceptos' && registro.Clave) {
                                mapaConceptosReales[registro.Clave.toString().toUpperCase()] = newId;
                            }

                            // Sincronizar contexto global si es un contrato
                            if (tabla === 'Contratos') {
                                globalIdContrato = newId;
                                ultimosIdsInsertados['Contratos'] = globalIdContrato;
                            }

                            // NUEVO: VINCULAR CAF AL CONTRATO (Post-Insert)
                            if (tabla === 'Convenios_Recurso' && globalIdContrato) {
                                dbUpdate('Contratos', { ID_Convenio_Vinculado: newId }, { ID_Contrato: globalIdContrato });
                            }

                            // NUEVO: Registrar en el mapa de periodos por nombre (OBJETO COMPLETO PARA FECHAS)
                            if (tabla === 'Programa_Periodo' && registro.Periodo) {
                                mapaPeriodosReales[registro.Periodo] = { ...registro, ID_Programa_Periodo: newId };
                            }

                            // Actualizar Caché Local para matching inmediato de siguientes registros en el mismo loop
                            if (cacheLocal[tabla]) cacheLocal[tabla].push({ ...registro });
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

    // --- RESOLUCIÓN ROBUSTA DE ID_CONTRATO PARA ANÁLISIS ---
    let idContratoCtx = datos.ID_Contrato_Contexto ||
        (pCtx ? (pCtx.ID_Contrato || pCtx.idContrato || pCtx.id_contrato) : null);

    if (!idContratoCtx) {
        const rootKeys = Object.keys(datos);
        const rootIdKey = rootKeys.find(k => k.toLowerCase() === 'id_contrato' || k.toLowerCase() === 'idcontrato');
        if (rootIdKey) idContratoCtx = datos[rootIdKey];
    }

    if (!idContratoCtx) {
        const lookIn = ['Contratos', 'Catalogo_Conceptos', 'Programa'];
        for (const t of lookIn) {
            if (datos[t] && Array.isArray(datos[t]) && datos[t][0]) {
                const itemKeys = Object.keys(datos[t][0]);
                const itemIdKey = itemKeys.find(k => k.toLowerCase() === 'id_contrato' || k.toLowerCase() === 'idcontrato');
                if (itemIdKey) {
                    idContratoCtx = datos[t][0][itemIdKey];
                    break;
                }
            }
        }
    }

    // Si aún no hay ID, intentar por Numero_Contrato (Búsqueda proactiva antes de mapear hijos)
    if (!idContratoCtx) {
        let needle = datos.Numero_Contrato ||
            (datos.Contratos && datos.Contratos[0] ? datos.Contratos[0].Numero_Contrato : null) ||
            (datos.Contrato && typeof datos.Contrato === 'string' ? datos.Contrato : null);

        if (needle) {
            needle = needle.toString().trim();
            if (needle.length > 0) {
                const dbSearch = dbSelect('Contratos', { Numero_Contrato: needle });
                if (dbSearch && dbSearch.length > 0) idContratoCtx = dbSearch[0].ID_Contrato;
            }
        }
    }

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
            else if (tabla === 'Estimaciones') skName = 'No_Estimacion';
            else if (tabla === 'Convenios_Modificatorios') skName = 'Numero_Convenio_Mod';
            else if (tabla === 'Catalogo_Conceptos') skName = 'Clave';
            let conceptCounter = 0; // Contador secuencial para ordenamiento visual

            datos[tabla].forEach(registro => {
                // --- ORDEN SECUENCIAL PARA CONCEPTOS ---
                if (tabla === 'Catalogo_Conceptos') {
                    conceptCounter++;
                    registro.Orden = conceptCounter;
                }

                // Inyectar contexto (Forzar ID real para matching correcto)
                if (idContratoCtx) {
                    const rel = RELACIONES_BD[tabla];
                    const rels = Array.isArray(rel) ? rel : [rel];
                    rels.forEach(r => {
                        if (r && r.padre === 'Contratos') {
                            registro.ID_Contrato = idContratoCtx;
                        }
                    });
                }

                let match = null;
                // Buscar por PK o SK
                if (registro[pkName]) {
                    const resPK = dbSelect(tabla, { [pkName]: registro[pkName] });
                    if (resPK && resPK.length > 0) match = resPK[0];
                }

                if (!match && tabla === 'Contratos' && idContratoCtx) {
                    const resCtx = dbSelect(tabla, { ID_Contrato: idContratoCtx });
                    if (resCtx && resCtx.length > 0) match = resCtx[0];
                }
                if (!match && skName && registro[skName]) {
                    const res = dbSelect(tabla, { [skName]: registro[skName] });
                    if (res && res.length > 0) match = res[0];
                }

                // Fuzzy match para Convenios_Recurso
                if (!match && tabla === 'Convenios_Recurso' && registro.Numero_Acuerdo) {
                    const numNorm = normalizeText(registro.Numero_Acuerdo);
                    const convenios = dbSelect('Convenios_Recurso');
                    match = (convenios || []).find(c => normalizeText(c.Numero_Acuerdo) === numNorm);
                }

                // Caso especial Catalogo (por Clave + Contrato)
                if (!match && tabla === 'Catalogo_Conceptos' && registro.ID_Contrato) {
                    const descNorm = normalizeText(registro.Descripcion || "");
                    const resClave = registro.Clave ? dbSelect(tabla, { Clave: registro.Clave, ID_Contrato: registro.ID_Contrato }) : [];
                    if (resClave && resClave.length > 0) match = resClave[0];

                    if (!match && descNorm) {
                        const catalogo = dbSelect(tabla, { ID_Contrato: registro.ID_Contrato });
                        match = catalogo.find(c => normalizeText(c.Descripcion) === descNorm);
                    }
                }

                // RECUPERACIÓN GENÉRICA DE HUÉRFANOS (Natural Key match con ID_Contrato vacío)
                if (!match && skName && registro[skName]) {
                    const rel = typeof RELACIONES_BD !== 'undefined' ? RELACIONES_BD[tabla] : null;
                    const rels = Array.isArray(rel) ? rel : [rel];
                    // Solo aplica si es hijo directo de Contratos con FK ID_Contrato
                    if (rels.some(r => r && r.padre === 'Contratos' && r.fk === 'ID_Contrato')) {
                        const resHuertano = dbSelect(tabla, { [skName]: registro[skName], ID_Contrato: "" });
                        if (resHuertano && resHuertano.length > 0) match = resHuertano[0];
                    }
                }

                // Caso especial Matriz (por ID_Concepto + Clave_Insumo)
                if (!match && tabla === 'Matriz_Insumos' && registro.ID_Concepto && registro.Clave_Insumo) {
                    const res = dbSelect(tabla, { ID_Concepto: registro.ID_Concepto, Clave_Insumo: registro.Clave_Insumo });
                    if (res && res.length > 0) match = res[0];
                }

                if (match) {
                    // Sincronizar contexto si encontramos el contrato en DB
                    if (tabla === 'Contratos') {
                        idContratoCtx = match.ID_Contrato;
                    }

                    // CORRECCIÓN: Permitir que la UI muestre los cambios financieros si vienen de la APU
                    const esActualizacionDesdeAPU = (tabla === 'Catalogo_Conceptos' && (tipo === 'MATRIZ_INSUMOS' || tipo === 'APU'));
                    const camposProtegidos = esActualizacionDesdeAPU 
                        ? ['Clave'] 
                        : (tabla === 'Catalogo_Conceptos' ? ['Clave', 'Descripcion', 'Unidad'] : []);

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

/**
 * Limpia y repara respuestas JSON de la IA que pueden traer errores de sintaxis
 * como comas finales, bloques de markdown o caracteres de control.
 */
function limpiarJSONIA(texto) {
    if (!texto) return "{}";

    let limpio = texto.trim();

    // 1. Extraer contenido de bloques markdown si existen
    const matchJson = limpio.match(/```json\s*([\s\S]*?)\s*```/) || limpio.match(/```\s*([\s\S]*?)\s*```/);
    if (matchJson) {
        limpio = matchJson[1].trim();
    }

    // 2. Eliminar comas finales (Trailing commas) - Causante #1 de SyntaxError
    // Busca una coma seguida de un cierre de objeto or array, ignorando espacios y saltos de línea
    limpio = limpio.replace(/,\s*(\]|\})/g, '$1');

    // 3. Eliminar caracteres de control (C0 y C1) que rompen JSON.parse
    limpio = limpio.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, "");

    return limpio;
}
