// Esquema Relacional de Base de Datos SGC
const ESQUEMA_BD = {
    'Usuarios_Sistema': ['ID_Usuario', 'Username', 'Nombre_Full', 'Rol', 'Email', 'Activo'],
    'Log_Actividad': ['ID_Log', 'ID_Usuario', 'Accion', 'Tabla_Afectada', 'Timestamp', 'Detalles'],
    'Parametros_Sistema': ['Clave_Parametro', 'Valor_Parametro', 'Descripcion'],
    'Estadisticas_Financieras': ['ID_Periodo', 'Año', 'Mes', 'Monto_Ejecutado', 'Monto_Proyectado'],
    'Conversaciones_IA': ['ID_Conversacion', 'Fecha_Hora', 'Usuario', 'Prompt', 'Respuesta'],
    'Convenios_Recurso': ['ID_Convenio', 'Numero_Acuerdo', 'Nombre_Fondo', 'Monto_Apoyo', 'Fecha_Firma', 'Vigencia_Fin', 'Objeto_Programa', 'Estado', 'Link_Sharepoint'],
    'Contratistas': ['ID_Contratista', 'Razon_Social', 'RFC', 'Domicilio_Fiscal', 'Representante_Legal', 'Telefono', 'Banco', 'Cuenta_Bancaria', 'Cuenta_CLABE'],
    'Contratos': [
        'ID_Contrato', 'Numero_Contrato', 'ID_Convenio_Vinculado', 'ID_Contratista',
        'Objeto_Contrato', 'Tipo_Contrato', 'Area_Responsable',
        'No_Concurso', 'Modalidad_Adjudicacion', 'Fecha_Adjudicacion',
        'Monto_Total_Sin_IVA', 'Monto_Total_Con_IVA',
        'Fecha_Firma', 'Fecha_Inicio_Obra', 'Fecha_Fin_Obra', 'Plazo_Ejecucion_Dias',
        'Porcentaje_Amortizacion_Anticipo', 'Porcentaje_Penas_Convencionales',
        'No_Fianza_Cumplimiento', 'Monto_Fianza_Cumplimiento', 'No_Fianza_Anticipo', 'Monto_Fianza_Anticipo',
        'No_Fianza_Garantia', 'Monto_Fianza_Garantia', 'No_Fianza_Vicios_Ocultos', 'Monto_Fianza_Vicios_Ocultos',
        'Estado', 'Retencion_Vigilancia_Pct', 'Retencion_Garantia_Pct', 'Otras_Retenciones_Pct',
        'Nombre_Residente_Dependencia', 'Link_Sharepoint',
        'Pct_Indirectos_Oficina', 'Pct_Indirectos_Campo', 'Pct_Indirectos_Totales', 'Pct_Financiamiento', 'Pct_Utilidad', 'Pct_Cargos_SFP', 'Pct_ISN'
    ],
    'Convenios_Modificatorios': ['ID_Convenio_Mod', 'ID_Contrato', 'Numero_Convenio_Mod', 'Tipo_Modificacion', 'Nuevo_Monto_Con_IVA', 'Nueva_Fecha_Fin', 'Motivo', 'Link_Sharepoint'],
    'Anticipos': ['ID_Anticipo', 'ID_Contrato', 'Porcentaje_Otorgado', 'Monto_Anticipo', 'Fecha_Pago', 'Monto_Amortizado_Acumulado', 'Saldo_Por_Amortizar'],
    'Catalogo_Conceptos': ['ID_Concepto', 'ID_Contrato', 'Clave', 'Descripcion', 'Unidad', 'Cantidad_Contratada', 'Precio_Unitario', 'Importe_Total_Sin_IVA', 'Orden', 'Costo_Directo'],
    'Programa': ['ID_Numero_Programa', 'ID_Contrato', 'Tipo_Programa', 'Programa', 'Fecha_Inicio', 'Fecha_Termino'],
    'Programa_Periodo': ['ID_Programa_Periodo', 'ID_Numero_Programa', 'Numero_Periodo', 'Periodo', 'Fecha_Inicio', 'Fecha_Termino'],
    'Programa_Ejecucion': ['ID_Programa', 'ID_Concepto', 'ID_Programa_Periodo', 'Fecha_Inicio', 'Fecha_Fin', 'Monto_Programado', 'Avance_Programado_Pct', 'Link_Sharepoint'],
    'Estimaciones': ['ID_Estimacion', 'ID_Contrato', 'No_Estimacion', 'Tipo_Estimacion', 'Periodo_Inicio', 'Periodo_Fin', 'Monto_Bruto_Estimado', 'Deduccion_Surv_05_Monto', 'Subtotal', 'IVA', 'Monto_Neto_A_Pagar', 'Avance_Acumulado_Anterior', 'Avance_Actual', 'Estado_Validacion', 'Link_Sharepoint', 'Avance_Anterior_Porcentaje', 'Avance_Actual_Porcentaje', 'Avance_Acumulado_Porcentaje'],
    'Detalle_Estimacion': ['ID_Detalle', 'ID_Estimacion', 'ID_Concepto', 'Cantidad_Estimada_Periodo', 'Precio_Unitario_Contrato', 'Importe_Este_Periodo', 'Avance_Acumulado_Porcentaje', 'Importe_Acumulado', 'Avance_Periodo_Porcentaje'],
    'Deducciones_Retenciones': ['ID_Deduccion', 'ID_Estimacion', 'Tipo_Deduccion', 'Monto_Deducido', 'Concepto_Deduccion'],
    'Facturas': ['ID_Factura', 'ID_Estimacion', 'Folio_Fiscal_UUID', 'No_Factura', 'Fecha_Emision', 'Monto_Facturado', 'Estatus_SAT', 'Link_Sharepoint'],
    'Pagos_Emitidos': ['ID_Pago', 'ID_Estimacion', 'Fecha_Pago', 'Monto_Pagado', 'Referencia_Bancaria', 'Estatus_Pago'],
    'Matriz_Insumos': ['ID_Matriz', 'ID_Concepto', 'Tipo_Insumo', 'Clave_Insumo', 'Descripcion', 'Unidad', 'Costo_Unitario', 'Rendimiento_Cantidad', 'Importe', 'Porcentaje_Incidencia'],
    'Validacion_Archivos': ['ID_Validacion', 'ID_Estimacion', 'Tipo_Archivo', 'Fecha_Carga', 'Estado_Validacion', 'Checklist_JSON', 'Observaciones_Resumen']
};

function doGet(e) {
    try {
        return HtmlService.createTemplateFromFile('index')
            .evaluate()
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setTitle('Sistema de Gestión de Contratos');
    } catch (error) {
        return ContentService.createTextOutput("Error cargando la aplicación: " + error.message);
    }
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generarRespuesta(exito, resultado, nombreFuncion = "") {
    if (exito) {
        return { success: true, data: sanitizeData(resultado) };
    } else {
        return { success: false, error: resultado, functionName: nombreFuncion };
    }
}

/**
 * Sanitiza recursivamente un objeto o array para asegurar que sea serializable por GAS.
 * Convierte objetos Date a strings ISO.
 */
function sanitizeData(data) {
    if (data === null || data === undefined) return data;
    if (data instanceof Date) return Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd'T12:00:00'");
    if (Array.isArray(data)) return data.map(item => sanitizeData(item));
    if (typeof data === 'object' && data !== null) {
        const sanitized = {};
        for (const key in data) {
            sanitized[key] = sanitizeData(data[key]);
        }
        return sanitized;
    }
    return data;
}

function getView(nombreVista) {
    try {
        const html = HtmlService.createTemplateFromFile(nombreVista).evaluate().getContent();
        return generarRespuesta(true, html);
    } catch (error) {
        return generarRespuesta(false, error.message, 'getView');
    }
}

function getContratosData() {
    const nombreFuncion = "getContratosData";
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Contratos');
        if (!sheet) throw new Error("No se encontró la hoja 'Contratos'");
        const data = getSafeData(sheet);
        if (data.length <= 1) return generarRespuesta(true, []);
        const cabeceras = data[0];
        const jsonData = [];
        for (let i = 1; i < data.length; i++) {
            const fila = data[i];
            let isEmptyRow = true;
            let obj = {};
            for (let j = 0; j < cabeceras.length; j++) {
                let cellValue = fila[j];
                if (cellValue instanceof Date) {
                    cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd'T12:00:00'");
                }
                if (cellValue !== "" && cellValue !== null && cellValue !== undefined) isEmptyRow = false;
                obj[cabeceras[j]] = cellValue;
            }
            if (!isEmptyRow) jsonData.push(obj);
        }
        return generarRespuesta(true, jsonData);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

function getDetalleContrato(idContrato) {
    const nombreFuncion = "getDetalleContrato";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Contratos');
        if (!sheet) throw new Error("No se encontró la hoja 'Contratos'");
        const data = getSafeData(sheet);
        if (data.length <= 1) return generarRespuesta(false, "No hay datos en la hoja Contratos", nombreFuncion);
        const cabeceras = data[0];
        const idxIdContrato = getColIndex(cabeceras, 'ID_Contrato');
        const idxIdContratista = getColIndex(cabeceras, 'ID_Contratista');

        if (idxIdContrato === -1) throw new Error("No se encontró la columna 'ID_Contrato' en la cabecera.");

        let contratoEncontrado = null;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdContrato]) === String(idContrato)) {
                let obj = {};
                for (let j = 0; j < cabeceras.length; j++) {
                    let cellValue = data[i][j];
                    if (cellValue instanceof Date) {
                        cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd'T12:00:00'");
                    }
                    obj[cabeceras[j]] = cellValue;
                }
                contratoEncontrado = obj;
                break;
            }
        }

        if (contratoEncontrado && idxIdContratista !== -1) {
            const idContratistaValue = contratoEncontrado['ID_Contratista'];
            const sheetC = ss.getSheetByName('Contratistas');
            if (sheetC) {
                const dataC = getSafeData(sheetC);
                const cabC = dataC[0];
                const idxBuscaId = getColIndex(cabC, 'ID_Contratista');
                const idxNombre = getColIndex(cabC, 'Razon_Social');
                if (idxBuscaId !== -1 && idxNombre !== -1) {
                    for (let k = 1; k < dataC.length; k++) {
                        if (String(dataC[k][idxBuscaId]) === String(idContratistaValue)) {
                            contratoEncontrado['Nombre_Contratista'] = dataC[k][idxNombre];
                            break;
                        }
                    }
                }
            }
        }

        return contratoEncontrado ? generarRespuesta(true, contratoEncontrado) : generarRespuesta(false, `No se encontró el contrato con ID: ${idContrato}`, nombreFuncion);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Actualiza los datos de un contrato existente en la hoja 'Contratos'.
 * @param {string} idContrato - El ID único del contrato a actualizar.
 * @param {Object} nuevosDatos - Objeto con los campos y valores a actualizar.
 * @returns {Object} Respuesta estandarizada.
 */
function actualizarDatosContrato(idContrato, nuevosDatos) {
    const nombreFuncion = "actualizarDatosContrato";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Contratos');
        if (!sheet) throw new Error("No se encontró la hoja 'Contratos'.");

        const data = getSafeData(sheet);
        const cabeceras = data[0];
        const idxId = getColIndex(cabeceras, 'ID_Contrato');

        if (idxId === -1) throw new Error("No se encontró la columna ID_Contrato.");

        let rowIdx = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxId]) === String(idContrato)) {
                rowIdx = i + 1; // 1-indexed for sheet
                break;
            }
        }

        if (rowIdx === -1) throw new Error(`No se encontró el contrato con ID: ${idContrato}`);

        // Mapear los datos a las columnas correspondientes
        for (let key in nuevosDatos) {
            let colIdx = getColIndex(cabeceras, key);
            if (colIdx !== -1) {
                let valor = nuevosDatos[key];

                // Conversión de datos según el tipo esperado
                if (key.includes('Fecha') && valor) {
                    valor = new Date(valor + "T12:00:00"); // Forzar mediodía para evitar desfases por zona horaria
                } else if (key.includes('Monto') || key.includes('Porcentaje') || key.includes('Plazo') || key.includes('Pct')) {
                    valor = parseFloat(valor) || 0;
                }

                sheet.getRange(rowIdx, colIdx + 1).setValue(valor);
            }
        }

        return generarRespuesta(true, { id: idContrato }, nombreFuncion);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

function getDashboardData() {
    const nombreFuncion = "getDashboardData";
    const sanitizeDate = (d) => {
        if (d instanceof Date) {
            try {
                return Utilities.formatDate(d, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd'T12:00:00'");
            } catch (e) { return null; }
        }
        return d;
    };

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetContratos = ss.getSheetByName('Contratos');
        const sheetEstimaciones = ss.getSheetByName('Estimaciones');
        const sheetPagos = ss.getSheetByName('Pagos_Emitidos');
        const sheetPrograma = ss.getSheetByName('Programa_Ejecucion');
        const sheetPeriodos = ss.getSheetByName('Programa_Periodo');

        let totalContratado = 0;
        let totalFiniquitado = 0;
        let totalPagado = 0;
        let contratosActivos = 0;
        let estimacionesPendientes = 0;
        let fianzasPorVencer = 0;

        const parseMonto = (v) => {
            if (typeof v === 'number') return v;
            if (!v) return 0;
            let n = parseFloat(String(v).replace(/[$,\s]/g, ''));
            return isNaN(n) ? 0 : n;
        };

        // 1. Datos de Contratos
        let contratosRecientes = [];
        const conteoEstados = {};
        try {
            const dataC = getSafeData(sheetContratos);
            if (dataC && dataC.length > 1) {
                const cabC = dataC[0];
                const idxMonto = getColIndex(cabC, 'Monto_Total_Con_IVA');
                const idxEstadoCol = getColIndex(cabC, 'Estado');
                const idxEstado = idxEstadoCol !== -1 ? idxEstadoCol : getColIndex(cabC, 'Situación');
                const idxId = getColIndex(cabC, 'ID_Contrato');
                const idxNum = getColIndex(cabC, 'Numero_Contrato');
                const idxFecha = getColIndex(cabC, 'Fecha_Firma');

                const todos = [];
                for (let i = 1; i < dataC.length; i++) {
                    const f = dataC[i];
                    const monto = parseMonto(f[idxMonto]);
                    const estado = String(f[idxEstado] || 'N/A').trim();
                    const estadoLow = estado.toLowerCase();

                    conteoEstados[estado] = (conteoEstados[estado] || 0) + 1;

                    if (['activo', 'vigente', 'en proceso'].includes(estadoLow) || estado === '') {
                        contratosActivos++;
                    }

                    if (estadoLow.includes('finiquit') || estadoLow.includes('cerrado') || estadoLow.includes('concluido')) {
                        totalFiniquitado += monto;
                    }

                    totalContratado += monto;

                    todos.push({
                        ID_Contrato: idxId !== -1 ? f[idxId] : '',
                        Numero_Contrato: idxNum !== -1 ? f[idxNum] : 'S/N',
                        Estado: estado,
                        Monto_Total_Con_IVA: monto,
                        Fecha_Firma_Date: (idxFecha !== -1 && f[idxFecha] instanceof Date) ? f[idxFecha] : new Date(0)
                    });
                }
                todos.sort((a, b) => b.Fecha_Firma_Date - a.Fecha_Firma_Date);
                contratosRecientes = todos.slice(0, 8).map(c => {
                    const sanitized = { ...c };
                    sanitized.Fecha_Firma = sanitizeDate(c.Fecha_Firma_Date);
                    delete sanitized.Fecha_Firma_Date; // Borrar objeto Date para evitar errores de serialización
                    return sanitized;
                });
            }
        } catch (e) {
            console.error("Error procesando contratos:", e);
        }

        // 2. Datos de Pagos
        try {
            if (sheetPagos) {
                const dataP = getSafeData(sheetPagos);
                if (dataP && dataP.length > 1) {
                    const cabP = dataP[0];
                    const idxMontoP = getColIndex(cabP, 'Monto_Pagado');
                    for (let i = 1; i < dataP.length; i++) {
                        totalPagado += parseMonto(dataP[i][idxMontoP]);
                    }
                }
            }
        } catch (e) { console.error("Error procesando pagos:", e); }

        // 3. Estimaciones Pendientes
        try {
            if (sheetEstimaciones) {
                const dataE = getSafeData(sheetEstimaciones);
                if (dataE && dataE.length > 1) {
                    const cabE = dataE[0];
                    const idxEst = getColIndex(cabE, 'Estado_Validacion');
                    if (idxEst !== -1) {
                        for (let i = 1; i < dataE.length; i++) {
                            const est = String(dataE[i][idxEst] || '').toLowerCase();
                            if (['pendiente', 'en revisión', 'en revision'].includes(est)) {
                                estimacionesPendientes++;
                            }
                        }
                    }
                }
            }
        } catch (e) { console.error("Error procesando estimaciones:", e); }

        // 4. Proyecciones
        let proyeccionesMensuales = {};
        let corteEjercicios = {};
        try {
            if (sheetPrograma && sheetPeriodos) {
                const dataProg = getSafeData(sheetPrograma);
                const dataPer = getSafeData(sheetPeriodos);

                if (dataProg && dataProg.length > 1 && dataPer && dataPer.length > 1) {
                    const cabProg = dataProg[0];
                    const cabPer = dataPer[0];
                    const idxMontoProg = getColIndex(cabProg, 'Monto_Programado');
                    const idxIdPerProg = getColIndex(cabProg, 'ID_Programa_Periodo');
                    const idxIdPer = getColIndex(cabPer, 'ID_Programa_Periodo');
                    const idxFechaPer = getColIndex(cabPer, 'Fecha_Inicio');

                    const mapaPeriodos = {};
                    for (let i = 1; i < dataPer.length; i++) {
                        mapaPeriodos[String(dataPer[i][idxIdPer])] = dataPer[i][idxFechaPer];
                    }

                    for (let i = 1; i < dataProg.length; i++) {
                        const idPer = String(dataProg[i][idxIdPerProg]);
                        const monto = parseMonto(dataProg[i][idxMontoProg]);
                        const fecha = mapaPeriodos[idPer];

                        if (fecha instanceof Date) {
                            const year = fecha.getFullYear();
                            const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
                            const keyMes = `${year}-${mes}`;
                            proyeccionesMensuales[keyMes] = (proyeccionesMensuales[keyMes] || 0) + monto;
                            corteEjercicios[year] = (corteEjercicios[year] || 0) + monto;
                        }
                    }
                }
            }
        } catch (e) {
            console.error("Error procesando proyecciones:", e);
            proyeccionesMensuales = {};
            corteEjercicios = {};
        }

        return generarRespuesta(true, {
            kpis: {
                totalContratado,
                totalPagado,
                totalFiniquitado,
                porEjercer: totalContratado - totalPagado,
                contratosActivos,
                estimacionesPendientes,
                fianzasPorVencer
            },
            contratosRecientes,
            conteoEstados,
            proyeccionesMensuales,
            corteEjercicios
        });
    } catch (error) {
        console.error("Error fatal en getDashboardData:", error);
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

function getConceptosContrato(idContrato) {
    const nombreFuncion = "getConceptosContrato";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheet) return generarRespuesta(true, []);
        const data = getSafeData(sheet);
        if (data.length <= 1) return generarRespuesta(true, []);
        const cabeceras = data[0];
        const idxIdContrato = getColIndex(cabeceras, 'ID_Contrato');
        if (idxIdContrato === -1) throw new Error("No se encontró la columna 'ID_Contrato' en Catalogo_Conceptos");
        const conceptos = [];
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdContrato]) === String(idContrato)) {
                let obj = {};
                for (let j = 0; j < cabeceras.length; j++) {
                    obj[cabeceras[j]] = data[i][j];
                }
                conceptos.push(obj);
            }
        }

        // Verificar cuáles conceptos tienen matriz asociada
        const sheetMatriz = ss.getSheetByName('Matriz_Insumos');
        const idsConMatriz = new Set();
        if (sheetMatriz) {
            const dataM = getSafeData(sheetMatriz);
            const idxMIdConc = getColIndex(dataM[0], 'ID_Concepto');
            if (idxMIdConc !== -1) {
                for (let k = 1; k < dataM.length; k++) {
                    idsConMatriz.add(String(dataM[k][idxMIdConc]));
                }
            }
        }

        conceptos.forEach(c => {
            c.Tiene_Matriz = idsConMatriz.has(String(c.ID_Concepto));
        });

        // Ordenar por la columna 'Orden' si existe
        conceptos.sort((a, b) => {
            const ordenA = parseFloat(a.Orden) || 999999;
            const ordenB = parseFloat(b.Orden) || 999999;
            return ordenA - ordenB;
        });

        return generarRespuesta(true, conceptos);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Obtiene el programa de ejecución en formato matricial (Conceptos vs Meses).
 * @param {string|number} idContrato - ID del contrato.
 */
/**
 * Obtiene el programa de ejecución en formato matricial (Conceptos vs Meses).
 * Implementa 3 niveles: Programa -> Programa_Periodo -> Programa_Ejecucion.
 * @param {string|number} idContrato - ID del contrato.
 */
function getProgramaEjecucion(idContrato) {
    const nombreFuncion = "getProgramaEjecucion";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 0. Inicializar hojas si no existen
        let sheetProg = ss.getSheetByName('Programa');
        let sheetPer = ss.getSheetByName('Programa_Periodo');
        let sheetEjec = ss.getSheetByName('Programa_Ejecucion');

        if (!sheetProg) {
            sheetProg = ss.insertSheet('Programa');
            sheetProg.appendRow(ESQUEMA_BD['Programa']);
        }
        if (!sheetPer) {
            sheetPer = ss.insertSheet('Programa_Periodo');
            sheetPer.appendRow(ESQUEMA_BD['Programa_Periodo']);
        }
        if (!sheetEjec) {
            sheetEjec = ss.insertSheet('Programa_Ejecucion');
            sheetEjec.appendRow(ESQUEMA_BD['Programa_Ejecucion']);
        }

        // 1. Obtener detalles del contrato para conocer el rango de fechas
        const detalleRes = getDetalleContrato(idContrato);
        if (!detalleRes.success) throw new Error(detalleRes.error);
        const contrato = detalleRes.data;

        if (!contrato.Fecha_Inicio_Obra || !contrato.Fecha_Fin_Obra) {
            return generarRespuesta(true, { meses: [], conceptosMatriz: [] });
        }

        const fInicio = parseSafeDate(contrato.Fecha_Inicio_Obra);
        const fFin = parseSafeDate(contrato.Fecha_Fin_Obra);

        if (!fInicio || !fFin) {
            return generarRespuesta(true, { meses: [], conceptosMatriz: [] });
        }

        // 2. Obtener o Crear registro en la tabla Programa (Nivel 1)
        const dataProg = getSafeData(sheetProg);
        const cabecerasProg = dataProg[0];
        const idxProgIdNum = getColIndex(cabecerasProg, 'ID_Numero_Programa');
        const idxProgIdCont = getColIndex(cabecerasProg, 'ID_Contrato');

        let programa = dataProg.find(r => String(r[idxProgIdCont]) === String(idContrato));
        let idNumeroPrograma;

        if (!programa) {
            idNumeroPrograma = dataProg.length > 1 ? (Math.max(...dataProg.slice(1).map(r => parseInt(r[idxProgIdNum]) || 0)) + 1) : 1;
            sheetProg.appendRow([
                idNumeroPrograma,
                idContrato,
                "MENSUAL", // Forzado a Mensual
                "ORIGINAL",
                fInicio,
                fFin
            ]);
        } else {
            idNumeroPrograma = programa[idxProgIdNum];
        }

        // 3. Gestionar Periodos en Programa_Periodo (Nivel 2)
        const dataPer = getSafeData(sheetPer);
        const cabecerasPer = dataPer[0];
        const idxPerId = getColIndex(cabecerasPer, 'ID_Programa_Periodo');
        const idxPProgNum = getColIndex(cabecerasPer, 'ID_Numero_Programa');
        const idxPerNum = getColIndex(cabecerasPer, 'Numero_Periodo');
        const idxPerLabel = getColIndex(cabecerasPer, 'Periodo');

        const periodosExistentes = dataPer.filter(r => String(r[idxPProgNum]) === String(idNumeroPrograma));
        const mesesLabels = [];
        let curr = new Date(fInicio.getFullYear(), fInicio.getMonth(), 1);
        const final = new Date(fFin.getFullYear(), fFin.getMonth(), 1);

        let contadorNumPer = 1;
        while (curr <= final) {
            let mesStart = new Date(curr.getFullYear(), curr.getMonth(), 1);
            let mesEnd = new Date(curr.getFullYear(), curr.getMonth() + 1, 0);

            if (curr.getFullYear() === fInicio.getFullYear() && curr.getMonth() === fInicio.getMonth()) mesStart = fInicio;
            if (curr.getFullYear() === fFin.getFullYear() && curr.getMonth() === fFin.getMonth()) mesEnd = fFin;

            const mesesMap = {
                'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo', 'April': 'Abril',
                'May': 'Mayo', 'June': 'Junio', 'July': 'Julio', 'August': 'Agosto',
                'September': 'Septiembre', 'October': 'Octubre', 'November': 'Noviembre', 'December': 'Diciembre'
            };
            const labelEn = Utilities.formatDate(curr, Session.getScriptTimeZone(), "MMMM yyyy");
            let label = labelEn;
            for (let en in mesesMap) {
                if (label.indexOf(en) !== -1) {
                    label = label.replace(en, mesesMap[en]);
                    break;
                }
            }
            label = label.toUpperCase();

            // BUSQUEDA ROBUSTA: Por número de periodo en lugar de solo label para evitar duplicados por idioma/formato
            let pExistente = periodosExistentes.find(p => parseInt(p[idxPerNum]) === contadorNumPer);
            let idPeriodo;

            if (!pExistente) {
                const fullDataPer = getSafeData(sheetPer);
                const lastId = fullDataPer.length > 1 ? Math.max(...fullDataPer.slice(1).map(r => parseInt(r[idxPerId]) || 0)) : 0;
                idPeriodo = lastId + 1;
                // Forzar formato texto con apostrofe para evitar auto-deteccion de fecha en Sheets
                sheetPer.appendRow([idPeriodo, idNumeroPrograma, contadorNumPer, "'" + label, mesStart, mesEnd]);
            } else {
                idPeriodo = pExistente[idxPerId];
                // Opcional: Actualizar el label si ha cambiado
                if (String(pExistente[idxPerLabel]) !== String(label) && !pExistente[idxPerLabel]) {
                    sheetPer.getRange(dataPer.indexOf(pExistente) + 1, idxPerLabel + 1).setValue("'" + label);
                }
            }

            mesesLabels.push({
                id: idPeriodo,
                numero: contadorNumPer,
                label: pExistente ? pExistente[idxPerLabel] : label,
                start: Utilities.formatDate(mesStart, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                end: Utilities.formatDate(mesEnd, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                isoStart: mesStart.toISOString(),
                isoEnd: mesEnd.toISOString(),
                year: curr.getFullYear(),
                month: curr.getMonth()
            });
            contadorNumPer++;
            curr.setMonth(curr.getMonth() + 1);
        }

        // 4. Obtener conceptos y Programa_Ejecucion (Nivel 3)
        const conceptosRes = getConceptosContrato(idContrato);
        const conceptos = conceptosRes.success ? conceptosRes.data : [];
        if (conceptos.length === 0) {
            return generarRespuesta(true, { meses: mesesLabels, conceptosMatriz: [] });
        }

        const dataEjec = getSafeData(sheetEjec);
        const cabecerasEjec = dataEjec[0];
        const idxEjProg = getColIndex(cabecerasEjec, 'ID_Programa');
        const idxEjConc = getColIndex(cabecerasEjec, 'ID_Concepto');
        const idxEjPerId = getColIndex(cabecerasEjec, 'ID_Programa_Periodo');
        const idxEjMonto = getColIndex(cabecerasEjec, 'Monto_Programado');

        // OPTIMIZACIÓN: Crear un mapa de ejecuciones para acceso O(1)
        const ejecuMap = {};
        if (dataEjec.length > 1) {
            dataEjec.slice(1).forEach(r => {
                if (String(r[idxEjProg]) === String(idNumeroPrograma)) {
                    const key = `${r[idxEjConc]}_${r[idxEjPerId]}`;
                    ejecuMap[key] = parseFloat(r[idxEjMonto]) || 0;
                }
            });
        }

        const conceptosMatriz = conceptos.map(c => {
            const montosMensuales = mesesLabels.map(mes => ejecuMap[`${c.ID_Concepto}_${mes.id}`] || 0);

            return {
                id: c.ID_Concepto,
                clave: c.Clave,
                descripcion: c.Descripcion,
                totalConcepto: parseFloat(c.Importe_Total_Sin_IVA) || 0,
                montos: montosMensuales
            };
        });

        return generarRespuesta(true, {
            idPrograma: idNumeroPrograma,
            meses: mesesLabels,
            conceptosMatriz: conceptosMatriz
        });

    } catch (error) {
        console.error(`Error en ${nombreFuncion}: ${error.stack}`);
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Inserta o actualiza un concepto en el catálogo.
 * Calcula automáticamente el importe (Cant * PU).
 */
function upsertConcepto(concepto) {
    const nombreFuncion = "upsertConcepto";
    try {
        if (!concepto.ID_Contrato) throw new Error("Falta ID_Contrato");
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheet) throw new Error("No se encontró la hoja 'Catalogo_Conceptos'");

        const data = getSafeData(sheet);
        const cabeceras = data[0];
        const idxIdConcepto = getColIndex(cabeceras, 'ID_Concepto');
        const idxIdContrato = getColIndex(cabeceras, 'ID_Contrato');
        const idxClave = getColIndex(cabeceras, 'Clave');
        const idxDesc = getColIndex(cabeceras, 'Descripcion');
        const idxUnidad = getColIndex(cabeceras, 'Unidad');
        const idxCant = getColIndex(cabeceras, 'Cantidad_Contratada');
        const idxPU = getColIndex(cabeceras, 'Precio_Unitario');
        const idxImporte = getColIndex(cabeceras, 'Importe_Total_Sin_IVA');

        // Calcular importe
        const cant = parseFloat(concepto.Cantidad_Contratada) || 0;
        const pu = parseFloat(concepto.Precio_Unitario) || 0;
        const importe = cant * pu;

        let filaAEditar = -1;
        if (concepto.ID_Concepto) {
            // Buscar fila para actualización
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][idxIdConcepto]) === String(concepto.ID_Concepto)) {
                    filaAEditar = i + 1;
                    break;
                }
            }
        }

        if (filaAEditar !== -1) {
            // Actualización
            sheet.getRange(filaAEditar, idxClave + 1).setValue(concepto.Clave);
            sheet.getRange(filaAEditar, idxDesc + 1).setValue(concepto.Descripcion);
            sheet.getRange(filaAEditar, idxUnidad + 1).setValue(concepto.Unidad);
            sheet.getRange(filaAEditar, idxCant + 1).setValue(cant);
            sheet.getRange(filaAEditar, idxPU + 1).setValue(pu);
            sheet.getRange(filaAEditar, idxImporte + 1).setValue(importe);

            // Sincronizar Programa si el importe cambió
            sincronizarProgramaPorMontoConcepto(String(concepto.ID_Contrato), String(concepto.ID_Concepto), importe);
        } else {
            // Inserción
            const nuevoId = "CON-" + Utilities.getUuid().substring(0, 8).toUpperCase();
            const nuevoOrden = data.length; // Posición al final
            sheet.appendRow([
                nuevoId,
                String(concepto.ID_Contrato),
                concepto.Clave,
                concepto.Descripcion,
                concepto.Unidad,
                cant,
                pu,
                importe,
                nuevoOrden
            ]);
        }

        const respuestaData = {
            ID_Concepto: concepto.ID_Concepto || nuevoId,
            ID_Contrato: concepto.ID_Contrato,
            Clave: concepto.Clave,
            Descripcion: concepto.Descripcion,
            Unidad: concepto.Unidad,
            Cantidad_Contratada: cant,
            Precio_Unitario: pu,
            Importe_Total_Sin_IVA: importe
        };

        return generarRespuesta(true, respuestaData);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Redistribuye los montos del programa proporcionalmente cuando cambia el total del concepto.
 */
/**
 * Redistribuye los montos del programa proporcionalmente cuando cambia el total del concepto en el catálogo.
 */
function sincronizarProgramaPorMontoConcepto(idContrato, idConcepto, nuevoMontoTotal) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetProg = ss.getSheetByName('Programa');
        const sheetEjec = ss.getSheetByName('Programa_Ejecucion');
        if (!sheetProg || !sheetEjec) return;

        // 0. Obtener subtotal contrato para recalcular %
        const conceptosRes = getConceptosContrato(idContrato);
        const conceptos = conceptosRes.data || [];
        const subtotalContrato = conceptos.reduce((acc, c) => acc + (parseFloat(c.Importe_Total_Sin_IVA) || 0), 0);

        // 1. Obtener ID_Programa (Nivel 1)
        const dataProg = getSafeData(sheetProg);
        if (dataProg.length <= 1) return;
        const cabecerasProg = dataProg[0];
        const idxProgIdNum = getColIndex(cabecerasProg, 'ID_Numero_Programa');
        const idxProgIdCont = getColIndex(cabecerasProg, 'ID_Contrato');

        const programa = dataProg.find(r => String(r[idxProgIdCont]) === String(idContrato));
        if (!programa) return;
        const idProgActivo = programa[idxProgIdNum];

        // 2. Obtener registros de ejecución para este concepto y programa (Nivel 3)
        const dataEjec = getSafeData(sheetEjec);
        if (dataEjec.length <= 1) return;
        const cabecerasEjec = dataEjec[0];
        const idxEjProg = getColIndex(cabecerasEjec, 'ID_Programa');
        const idxEjConc = getColIndex(cabecerasEjec, 'ID_Concepto');
        const idxEjMonto = getColIndex(cabecerasEjec, 'Monto_Programado');
        const idxEjPct = getColIndex(cabecerasEjec, 'Avance_Programado_Pct');

        const registros = [];
        let totalActualP = 0;

        for (let i = 1; i < dataEjec.length; i++) {
            if (String(dataEjec[i][idxEjProg]) == String(idProgActivo) && String(dataEjec[i][idxEjConc]) == String(idConcepto)) {
                const m = parseFloat(dataEjec[i][idxEjMonto]) || 0;
                totalActualP += m;
                registros.push({ fila: i + 1, monto: m });
            }
        }

        if (registros.length === 0 || totalActualP === 0) return;

        // 3. Redistribuir proporcionalmente
        const factorGlobal = nuevoMontoTotal / totalActualP;

        registros.forEach(reg => {
            const nuevoMontoP = reg.monto * factorGlobal;
            const nuevoPctP = subtotalContrato > 0 ? (nuevoMontoP / subtotalContrato) * 100 : 0;
            sheetEjec.getRange(reg.fila, idxEjMonto + 1).setValue(nuevoMontoP);
            sheetEjec.getRange(reg.fila, idxEjPct + 1).setValue(nuevoPctP);
        });

    } catch (e) {
        console.error("Error en sincronizarProgramaPorMontoConcepto:", e);
    }
}

/**
 * Actualiza o crea un registro en Programa_Ejecucion siguiendo la estructura de 3 niveles.
 */
function actualizarPeriodoPrograma(idContrato, idConcepto, idPrograma, idPeriodo, nuevoMonto) {
    const nombreFuncion = "actualizarPeriodoPrograma";
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheetEjec = ss.getSheetByName('Programa_Ejecucion');
        if (!sheetEjec) {
            sheetEjec = ss.insertSheet('Programa_Ejecucion');
            sheetEjec.appendRow(ESQUEMA_BD['Programa_Ejecucion']);
        }

        // 1. Obtener el catálogo para el tope y el cálculo de %
        const conceptosRes = getConceptosContrato(idContrato);
        const conceptos = conceptosRes.data || [];
        const concepto = conceptos.find(c => String(c.ID_Concepto) == String(idConcepto));
        if (!concepto) throw new Error("Concepto no encontrado en el catálogo.");

        const topeMaximo = parseFloat(concepto.Importe_Total_Sin_IVA) || 0;
        const subtotalContrato = conceptos.reduce((acc, c) => acc + (parseFloat(c.Importe_Total_Sin_IVA) || 0), 0);

        const dataEjec = getSafeData(sheetEjec);
        const cabecerasEjec = dataEjec[0];
        const idxEjProg = getColIndex(cabecerasEjec, 'ID_Programa');
        const idxEjConc = getColIndex(cabecerasEjec, 'ID_Concepto');
        const idxEjPerId = getColIndex(cabecerasEjec, 'ID_Programa_Periodo');
        const idxEjMonto = getColIndex(cabecerasEjec, 'Monto_Programado');
        const idxEjPct = getColIndex(cabecerasEjec, 'Avance_Programado_Pct');

        let filaEncontrada = -1;
        let sumaOtrosPeriodos = 0;

        for (let i = 1; i < dataEjec.length; i++) {
            if (String(dataEjec[i][idxEjProg]) === String(idPrograma) && String(dataEjec[i][idxEjConc]) === String(idConcepto)) {
                if (String(dataEjec[i][idxEjPerId]) === String(idPeriodo)) {
                    filaEncontrada = i + 1;
                } else {
                    sumaOtrosPeriodos += parseFloat(dataEjec[i][idxEjMonto]) || 0;
                }
            }
        }

        // 2. Validar Tope
        if ((sumaOtrosPeriodos + nuevoMonto) > (topeMaximo + 0.10)) {
            return generarRespuesta(false, `Excede tope: Programado $${(sumaOtrosPeriodos + nuevoMonto).toFixed(2)} vs Catálogo $${topeMaximo.toFixed(2)}`);
        }

        // 3. Calcular % de avance relativo al TOTAL DEL CONTRATO
        const pctAvance = subtotalContrato > 0 ? (nuevoMonto / subtotalContrato) * 100 : 0;

        // 4. Acción (Eliminar si 0, de lo contrario Upsert)
        if (nuevoMonto <= 0) {
            if (filaEncontrada !== -1) {
                sheetEjec.deleteRow(filaEncontrada);
                return generarRespuesta(true, { success: true, message: "Removido" });
            }
            return generarRespuesta(true, { success: true });
        }

        if (filaEncontrada !== -1) {
            sheetEjec.getRange(filaEncontrada, idxEjMonto + 1).setValue(nuevoMonto);
            sheetEjec.getRange(filaEncontrada, idxEjPct + 1).setValue(pctAvance);
        } else {
            // Obtener fechas desde Programa_Periodo
            const sheetPer = ss.getSheetByName('Programa_Periodo');
            const dataPer = getSafeData(sheetPer);
            const cabPer = dataPer[0];
            const p = dataPer.find(r => String(r[getColIndex(cabPer, 'ID_Programa_Periodo')]) === String(idPeriodo));
            const fIni = p ? p[getColIndex(cabPer, 'Fecha_Inicio')] : "";
            const fFin = p ? p[getColIndex(cabPer, 'Fecha_Termino')] : "";

            // Seguir estrictamente ESQUEMA_BD: ['ID_Programa', 'ID_Concepto', 'ID_Programa_Periodo', 'Fecha_Inicio', 'Fecha_Fin', 'Monto_Programado', 'Avance_Programado_Pct', 'Link_Sharepoint']
            sheetEjec.appendRow([
                idPrograma,
                idConcepto,
                idPeriodo,
                fIni,
                fFin,
                nuevoMonto,
                pctAvance,
                ""
            ]);
        }

        return generarRespuesta(true, { success: true, message: "Guardado" });
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Elimina un concepto por su ID y limpia su programa asociado (Borrado en cascada).
 */
function eliminarConcepto(idConcepto) {
    const nombreFuncion = "eliminarConcepto";
    try {
        if (!idConcepto) throw new Error("ID de concepto no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Eliminar del Catálogo
        const sheetC = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheetC) throw new Error("No hay hoja de conceptos.");
        const dataC = getSafeData(sheetC);
        const idxIdC = getColIndex(dataC[0], 'ID_Concepto');

        let borradoC = false;
        for (let i = 1; i < dataC.length; i++) {
            if (String(dataC[i][idxIdC]) === String(idConcepto)) {
                sheetC.deleteRow(i + 1);
                borradoC = true;
                break;
            }
        }

        // 2. Eliminar del Programa_Ejecucion (Cascade)
        const sheetP = ss.getSheetByName('Programa_Ejecucion');
        if (sheetP) {
            const dataP = getSafeData(sheetP);
            const idxIdCP = getColIndex(dataP[0], 'ID_Concepto');
            // De abajo hacia arriba para evitar problemas con los índices al borrar
            for (let i = dataP.length - 1; i >= 1; i--) {
                if (String(dataP[i][idxIdCP]) === String(idConcepto)) {
                    sheetP.deleteRow(i + 1);
                }
            }
        }

        if (!borradoC) return generarRespuesta(false, "Concepto no encontrado en el catálogo", nombreFuncion);
        return generarRespuesta(true, "Concepto y registros de programa eliminados correctamente");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Actualiza masivamente la columna 'Orden' para los conceptos de un contrato.
 * @param {string} idContrato
 * @param {Array<string>} listaIdsOrdenados - IDs en el nuevo orden.
 */
function actualizarOrdenConceptos(idContrato, listaIdsOrdenados) {
    const nombreFuncion = "actualizarOrdenConceptos";
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheet) throw new Error("No hay hoja de catálogo.");

        const range = sheet.getDataRange();
        const data = range.getValues();
        const cabeceras = data[0];
        const idxId = getColIndex(cabeceras, 'ID_Concepto');
        let idxOrden = getColIndex(cabeceras, 'Orden');

        // Si no existe la columna Orden, la creamos
        if (idxOrden === -1) {
            idxOrden = cabeceras.length;
            sheet.getRange(1, idxOrden + 1).setValue('Orden');
        }

        // Mapa para buscar filas por ID rápidamente
        const mapaFilas = {};
        for (let i = 1; i < data.length; i++) {
            mapaFilas[String(data[i][idxId])] = i + 1;
        }

        // Actualizar cada ID en la lista con su nueva posición
        listaIdsOrdenados.forEach((id, index) => {
            const numFila = mapaFilas[String(id)];
            if (numFila) {
                sheet.getRange(numFila, idxOrden + 1).setValue(index + 1);
            }
        });

        return generarRespuesta(true, "Orden de conceptos sincronizado");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

function getColIndex(headers, colName) {
    if (!headers || !colName) return -1;
    const target = colName.toLowerCase().replace(/\s+/g, '');
    return headers.findIndex(h => {
        if (!h) return false;
        return h.toString().toLowerCase().replace(/\s+/g, '') === target;
    });
}

/**
 * Parsea una cadena de fecha de forma segura evitando desfases de zona horaria.
 */
function parseSafeDate(fechaStr) {
    if (!fechaStr) return null;
    try {
        if (fechaStr instanceof Date) return fechaStr;
        if (typeof fechaStr === 'string' && fechaStr.includes('T')) {
            const parts = fechaStr.split('T')[0].split('-');
            const tPart = (fechaStr.split('T')[1] || "0:0:0").split(':');
            return new Date(
                parseInt(parts[0]),
                parseInt(parts[1]) - 1,
                parseInt(parts[2]),
                parseInt(tPart[0]) || 0,
                parseInt(tPart[1]) || 0,
                parseInt(tPart[2]) || 0
            );
        }
        const d = new Date(fechaStr);
        return isNaN(d.getTime()) ? null : d;
    } catch (e) {
        return null;
    }
}

/**
 * Obtiene la configuración de los niveles 1 y 2 para un contrato.
 */

/**
 * Aplica los factores de sobrecosto a todos los conceptos de un contrato.

/**
 * Aplica los factores de sobrecosto a todos los conceptos de un contrato.
 * Actualiza tanto los porcentajes en el contrato como los P.U. en el catálogo.
 */
function aplicarFactoresSobrecostoCatalogo(idContrato, factores) {
    const nombreFuncion = "aplicarFactoresSobrecostoCatalogo";
    try {
        if (!idContrato) throw new Error("ID de contrato ausente.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Actualizar porcentajes en la tabla Contratos
        actualizarDatosContrato(idContrato, factores);

        // 2. Obtener el catálogo y recalcular cada concepto
        const sheetCat = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheetCat) throw new Error("No se encontró la hoja Catalogo_Conceptos");

        const data = getSafeData(sheetCat);
        if (data.length <= 1) return generarRespuesta(true, "Catálogo vacío.");

        const cab = data[0];
        const idxIdCont = getColIndex(cab, 'ID_Contrato');
        const idxPU = getColIndex(cab, 'Precio_Unitario');
        const idxCD = getColIndex(cab, 'Costo_Directo');
        const idxCant = getColIndex(cab, 'Cantidad_Contratada');
        const idxTotal = getColIndex(cab, 'Importe_Total_Sin_IVA');

        if (idxCD === -1) throw new Error("La columna 'Costo_Directo' no existe en el catálogo. Ejecute la sincronización de esquema.");

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdCont]) === String(idContrato)) {
                let costoDirecto = parseFloat(data[i][idxCD]) || 0;

                // Si el Costo Directo es 0 pero ya hay un P.U., lo tomamos como base inicial (Migración)
                if (costoDirecto === 0 && parseFloat(data[i][idxPU]) > 0) {
                    costoDirecto = parseFloat(data[i][idxPU]);
                    sheetCat.getRange(i + 1, idxCD + 1).setValue(costoDirecto);
                }

                // Cálculo Cascada
                let pctInd = parseFloat(factores.Pct_Indirectos_Totales) || 0;
                if (pctInd === 0) {
                    pctInd = (parseFloat(factores.Pct_Indirectos_Oficina) || 0) + (parseFloat(factores.Pct_Indirectos_Campo) || 0);
                }

                const sub1 = costoDirecto * (1 + (pctInd / 100));
                const sub2 = sub1 * (1 + ((parseFloat(factores.Pct_Financiamiento) || 0) / 100));
                const sub3 = sub2 * (1 + ((parseFloat(factores.Pct_Utilidad) || 0) / 100));

                // Cargos Adicionales (SFP 5 al millar formula: Sub3 * 0.005 / 0.995)
                const montoSFP = sub3 * (0.005 / 0.995);
                const montoISN = costoDirecto * ((parseFloat(factores.Pct_ISN) || 0) / 100);

                const puFinal = sub3 + montoSFP + montoISN;
                const cantidad = parseFloat(data[i][idxCant]) || 0;
                const importeTotal = puFinal * cantidad;

                // Actualizar en la hoja
                sheetCat.getRange(i + 1, idxPU + 1).setValue(puFinal);
                sheetCat.getRange(i + 1, idxTotal + 1).setValue(importeTotal);
            }
        }

        return generarRespuesta(true, "Catálogo actualizado con éxito.");

    } catch (e) {
        console.error("Error en " + nombreFuncion + ": " + e.message);
        return generarRespuesta(false, e.message, nombreFuncion);
    }
}

/**
 * Obtiene el detalle de la matriz (insumos) para un concepto específico.
 * @param {string|number} idConcepto - ID del concepto.
 */
function getMatrizConcepto(idConcepto) {
    const nombreFuncion = "getMatrizConcepto";
    try {
        if (!idConcepto) throw new Error("ID de concepto no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Obtener el concepto para saber a qué contrato pertenece
        const sheetCat = ss.getSheetByName('Catalogo_Conceptos');
        if (!sheetCat) throw new Error("No se encontró la hoja Catalogo_Conceptos");

        const dataCat = getSafeData(sheetCat);
        const cabCat = dataCat[0];
        const idxIdC = getColIndex(cabCat, 'ID_Concepto');
        const idxIdContrato = getColIndex(cabCat, 'ID_Contrato');

        let idContrato = null;
        for (let i = 1; i < dataCat.length; i++) {
            if (String(dataCat[i][idxIdC]) === String(idConcepto)) {
                idContrato = dataCat[i][idxIdContrato];
                break;
            }
        }

        // 2. Obtener datos del contrato (para sobrecostos)
        let contrato = null;
        if (idContrato) {
            const resC = getDetalleContrato(idContrato);
            if (resC.success) contrato = resC.data;
        }

        // 3. Obtener los insumos de la matriz
        const sheetMat = ss.getSheetByName('Matriz_Insumos');
        const insumos = [];
        if (sheetMat) {
            const dataMat = getSafeData(sheetMat);
            if (dataMat.length > 1) {
                const cabMat = dataMat[0];
                const idxIdConceptoMat = getColIndex(cabMat, 'ID_Concepto');
                for (let i = 1; i < dataMat.length; i++) {
                    if (String(dataMat[i][idxIdConceptoMat]) === String(idConcepto)) {
                        let obj = {};
                        for (let j = 0; j < cabMat.length; j++) {
                            obj[cabMat[j]] = dataMat[i][j];
                        }
                        insumos.push(obj);
                    }
                }
            }
        }

        return generarRespuesta(true, {
            insumos: insumos,
            contrato: contrato
        });
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Elimina todos los insumos (matriz) de un concepto específico.
 */
function purgarMatrizConcepto(idConcepto) {
    const nombreFuncion = "purgarMatrizConcepto";
    try {
        if (!idConcepto) throw new Error("ID de concepto no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Matriz_Insumos');
        if (!sheet) return generarRespuesta(true, "No hay matriz que purgar.");

        const data = getSafeData(sheet);
        const cabeceras = data[0];
        const idxIdConc = getColIndex(cabeceras, 'ID_Concepto');

        // Eliminar filas de abajo hacia arriba para no alterar índices
        let eliminados = 0;
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][idxIdConc]) === String(idConcepto)) {
                sheet.deleteRow(i + 1);
                eliminados++;
            }
        }

        return generarRespuesta(true, `Matriz purgada: ${eliminados} insumos eliminados.`);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Obtiene el resumen financiero total del contrato (CD, Indirectos, etc.)
 */
function getResumenFinancieroContrato(idContrato) {
    const nombreFuncion = "getResumenFinancieroContrato";
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Obtener Factores del Contrato
        const resContrato = getDetalleContrato(idContrato);
        if (!resContrato.success) throw new Error("No se pudo obtener el detalle del contrato.");
        const factores = resContrato.data;

        // 2. Obtener Conceptos
        const resConceptos = getConceptosContrato(idContrato);
        if (!resConceptos.success) throw new Error("No se pudieron obtener los conceptos.");
        const conceptos = resConceptos.data;

        // 3. Calcular Sumatoria CD * Cantidad
        let totalCD = 0;
        conceptos.forEach(c => {
            const cant = parseFloat(c.Cantidad_Contratada) || 0;
            const cd = parseFloat(c.Costo_Directo) || 0;
            totalCD += (cant * cd);
        });

        // 4. Aplicar Cascada de Sobre costos (Similar a la reportada en matrices)
        let pInd = parseFloat(factores.Pct_Indirectos_Totales) || 0;
        if (pInd === 0) {
            pInd = (parseFloat(factores.Pct_Indirectos_Oficina) || 0) + (parseFloat(factores.Pct_Indirectos_Campo) || 0);
        }

        const pFin = parseFloat(factores.Pct_Financiamiento) || 0;
        const pUtil = parseFloat(factores.Pct_Utilidad) || 0;
        const pISN = parseFloat(factores.Pct_ISN) || 0;
        const pSFP = parseFloat(factores.Pct_Cargos_SFP) || 0;

        const montoInd = totalCD * (pInd / 100);
        const sub1 = totalCD + montoInd;

        const montoFin = sub1 * (pFin / 100);
        const sub2 = sub1 + montoFin;

        const montoUtil = sub2 * (pUtil / 100);
        const sub3 = sub2 + montoUtil;

        // Cargos Adicionales (SFP + ISN)
        // El SFP suele calcularse como Monto / 0.995 * 0.005 si es el 5 al millar
        let montoSFP = 0;
        if (pSFP > 0) {
            // Si es 0.5 (5 al millar), usamos la fórmula pro: Sub3 * (0.005 / 0.995)
            // Para otros valores, aplicamos el porcentaje directo.
            if (pSFP === 0.5) {
                montoSFP = sub3 * (0.005 / 0.995);
            } else {
                montoSFP = sub3 * (pSFP / 100);
            }
        }

        const montoISN = totalCD * (pISN / 100);
        const totalCargos = montoSFP + montoISN;
        const sub4 = sub3 + totalCargos;

        return generarRespuesta(true, {
            Costo_Directo: totalCD,
            Indirectos: montoInd,
            Financiamiento: montoFin,
            Utilidad: montoUtil,
            Cargos_Adicionales: totalCargos,
            Precio_Unitario_Total: sub4,
            Factores: {
                Pct_Indirectos: pInd,
                Pct_Financiamiento: pFin,
                Pct_Utilidad: pUtil,
                Pct_SFP: pSFP,
                Pct_ISN: pISN
            }
        });

    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Recupera todas las estimaciones registradas para un contrato.
 */
function getEstimacionesContrato(idContrato) {
    const nombreFuncion = "getEstimacionesContrato";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Estimaciones');
        if (!sheet) return generarRespuesta(true, []);

        const data = getSafeData(sheet);
        const cabeceras = data[0];
        const idxIdContrato = getColIndex(cabeceras, 'ID_Contrato');

        const estimaciones = [];
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdContrato]) === String(idContrato)) {
                let obj = {};
                for (let j = 0; j < cabeceras.length; j++) {
                    let cellValue = data[i][j];
                    if (cellValue instanceof Date) {
                        cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
                    }
                    obj[cabeceras[j]] = cellValue;
                }
                estimaciones.push(obj);
            }
        }

        // Ordenar por número de estimación
        const idxNoEst = getColIndex(cabeceras, 'No_Estimacion');
        estimaciones.sort((a, b) => (parseInt(a.No_Estimacion) || 0) - (parseInt(b.No_Estimacion) || 0));

        return generarRespuesta(true, estimaciones);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Genera la siguiente estimación de forma consecutiva según los periodos del programa.
 */
function generarProximaEstimacion(idContrato) {
    const nombreFuncion = "generarProximaEstimacion";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Obtener estimaciones actuales para saber el siguiente número
        const resEst = getEstimacionesContrato(idContrato);
        if (!resEst.success) throw new Error(resEst.error);
        const estimacionesActuales = resEst.data;
        const siguienteNo = estimacionesActuales.length + 1;

        // 2. Obtener el programa y el periodo correspondiente
        const sheetProg = ss.getSheetByName('Programa');
        const sheetPer = ss.getSheetByName('Programa_Periodo');
        if (!sheetProg || !sheetPer) throw new Error("El contrato no tiene un programa configurado.");

        // Encontrar ID_Numero_Programa para este contrato
        const dataProg = getSafeData(sheetProg);
        const idxIdContratoP = getColIndex(dataProg[0], 'ID_Contrato');
        const idxIdNumProg = getColIndex(dataProg[0], 'ID_Numero_Programa');
        const progRow = dataProg.find(r => String(r[idxIdContratoP]) === String(idContrato));
        if (!progRow) throw new Error("No se encontró programa para este contrato.");
        const idNumProg = progRow[idxIdNumProg];

        // Buscar el periodo N
        const dataPer = getSafeData(sheetPer);
        const idxProgNumPer = getColIndex(dataPer[0], 'ID_Numero_Programa');
        const idxNoPer = getColIndex(dataPer[0], 'Numero_Periodo');
        const idxPerInicio = getColIndex(dataPer[0], 'Fecha_Inicio');
        const idxPerFin = getColIndex(dataPer[0], 'Fecha_Termino');

        const periodoRow = dataPer.find(r => String(r[idxProgNumPer]) === String(idNumProg) && parseInt(r[idxNoPer]) === siguienteNo);
        if (!periodoRow) throw new Error(`No existe el Periodo ${siguienteNo} en el programa. Configure más periodos.`);

        // 3. Crear el registro de la Estimación
        const sheetEst = ss.getSheetByName('Estimaciones') || ss.insertSheet('Estimaciones');
        if (sheetEst.getLastRow() === 0) sheetEst.appendRow(ESQUEMA_BD['Estimaciones']);

        const dataEst = getSafeData(sheetEst);
        const idEstimacion = dataEst.length > 1 ? (Math.max(...dataEst.slice(1).map(r => parseInt(r[0]) || 0)) + 1) : 1;

        const nuevaFila = [];
        const cabEst = dataEst.length > 0 ? dataEst[0] : [];
        cabEst.forEach(col => {
            switch (col) {
                case 'ID_Estimacion': nuevaFila.push(idEstimacion); break;
                case 'ID_Contrato': nuevaFila.push(idContrato); break;
                case 'No_Estimacion': nuevaFila.push(siguienteNo); break;
                case 'Tipo_Estimacion': nuevaFila.push("NORMAL"); break;
                case 'Periodo_Inicio': nuevaFila.push(periodoRow[idxPerInicio]); break;
                case 'Periodo_Fin': nuevaFila.push(periodoRow[idxPerFin]); break;
                case 'Estado_Validacion': nuevaFila.push("BORRADOR"); break;
                default: nuevaFila.push(0); break;
            }
        });
        sheetEst.appendRow(nuevaFila);

        // 4. Crear Detalle de Estimación (Solo conceptos con programación para este periodo - DELTAS)
        const sheetEjec = ss.getSheetByName('Programa_Ejecucion');
        if (sheetEjec) {
            const dataEjec = getSafeData(sheetEjec);
            const cabEjec = dataEjec[0];
            const idxIdProgPerEjec = getColIndex(cabEjec, 'ID_Programa_Periodo');
            const idxIdConceptoEjec = getColIndex(cabEjec, 'ID_Concepto');
            const idxMontoProgEjec = getColIndex(cabEjec, 'Monto_Programado');

            const idProgPeriodo = periodoRow[getColIndex(dataPer[0], 'ID_Programa_Periodo')];

            // Obtener conceptos para PU
            const resConceptos = getConceptosContrato(idContrato);
            const conceptosMap = {};
            if (resConceptos.success) {
                resConceptos.data.forEach(c => conceptosMap[String(c.ID_Concepto)] = parseFloat(c.Precio_Unitario) || 0);
            }

            const sheetDet = ss.getSheetByName('Detalle_Estimacion') || ss.insertSheet('Detalle_Estimacion');
            if (sheetDet.getLastRow() === 0) sheetDet.appendRow(ESQUEMA_BD['Detalle_Estimacion']);

            const dataDetSnapshot = getSafeData(sheetDet);
            const cabDet = dataDetSnapshot.length > 0 ? dataDetSnapshot[0] : [];
            const idxIdDet = getColIndex(cabDet, 'ID_Detalle');
            let lastIdDet = dataDetSnapshot.length > 1 ? Math.max(...dataDetSnapshot.slice(1).map(r => parseInt(r[0]) || 0)) : 0;

            const filasDetalle = [];
            dataEjec.forEach((row, i) => {
                if (i === 0) return;
                if (String(row[idxIdProgPerEjec]) === String(idProgPeriodo)) {
                    const montoProg = parseFloat(row[idxMontoProgEjec]) || 0;
                    if (montoProg > 0) {
                        const idC = String(row[idxIdConceptoEjec]);
                        const pu = conceptosMap[idC] || 0;
                        const cantProg = pu > 0 ? montoProg / pu : 0;

                        lastIdDet++;
                        const fila = [];
                        cabDet.forEach(col => {
                            switch (col) {
                                case 'ID_Detalle': fila.push(lastIdDet); break;
                                case 'ID_Estimacion': fila.push(idEstimacion); break;
                                case 'ID_Concepto': fila.push(idC); break;
                                case 'Cantidad_Estimada_Periodo': fila.push(cantProg); break;
                                case 'Precio_Unitario_Contrato': fila.push(pu); break;
                                case 'Importe_Este_Periodo': fila.push(montoProg); break;
                                default: fila.push(0); break;
                            }
                        });
                        filasDetalle.push(fila);
                    }
                }
            });

            if (filasDetalle.length > 0) {
                sheetDet.getRange(sheetDet.getLastRow() + 1, 1, filasDetalle.length, cabDet.length).setValues(filasDetalle);
            }
        }

        return generarRespuesta(true, { id: idEstimacion, no: siguienteNo }, "Estimación generada correctamente.");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Obtiene el detalle completo de una estimación, incluyendo carátula y conceptos.
 */
function getDetalleEstimacion(idEstimacion) {
    const nombreFuncion = "getDetalleEstimacion";
    try {
        if (!idEstimacion) throw new Error("ID de estimación no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Datos de la estimación
        const sheetEst = ss.getSheetByName('Estimaciones');
        const dataEst = getSafeData(sheetEst);
        const cabEst = dataEst[0];
        const idxIdEst = getColIndex(cabEst, 'ID_Estimacion');
        const filaEst = dataEst.find(r => String(r[idxIdEst]) === String(idEstimacion));
        if (!filaEst) throw new Error("No se encontró el registro de la estimación.");

        let estimacion = {};
        cabEst.forEach((col, i) => {
            let val = filaEst[i];
            if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            estimacion[col] = val;
        });

        // 2. Datos del Contrato (para factores y encabezados)
        const resCont = getDetalleContrato(estimacion.ID_Contrato);
        const contrato = resCont.success ? resCont.data : {};

        // 3. Detalle de Conceptos (VISTA COMPLETA - Incluye los que no están en la BD de esta estimación)
        const sheetDet = ss.getSheetByName('Detalle_Estimacion');
        const resConceptos = getConceptosContrato(estimacion.ID_Contrato);
        if (!resConceptos.success) throw new Error(resConceptos.error);
        const catalogo = resConceptos.data;

        // Mapear los deltas guardados en la BD para esta estimación
        const deltasMap = {};
        if (sheetDet) {
            const dataDet = getSafeData(sheetDet);
            const cabDet = dataDet.length > 0 ? dataDet[0] : [];
            const idxIdEstDet = getColIndex(cabDet, 'ID_Estimacion');
            const idxIdConceptoDet = getColIndex(cabDet, 'ID_Concepto');

            for (let i = 1; i < dataDet.length; i++) {
                if (String(dataDet[i][idxIdEstDet]) === String(idEstimacion)) {
                    const idC = String(dataDet[i][idxIdConceptoDet]);
                    const obj = {};
                    cabDet.forEach((col, j) => obj[col] = dataDet[i][j]);
                    deltasMap[idC] = obj;
                }
            }
        }

        // Construir lista final mezclando catálogo + deltas
        const detalles = catalogo.map(c => {
            const d = deltasMap[String(c.ID_Concepto)] || {};
            return {
                ID_Detalle: d.ID_Detalle || null,
                ID_Estimacion: idEstimacion,
                ID_Concepto: c.ID_Concepto,
                Clave: c.Clave || '',
                Descripcion: c.Descripcion || '',
                Unidad: c.Unidad || '',
                Cantidad_Contratada: parseFloat(c.Cantidad_Contratada) || 0,
                Precio_Unitario_Contrato: parseFloat(c.Precio_Unitario) || 0,
                Importe_Total_Sin_IVA: parseFloat(c.Importe_Total_Sin_IVA) || 0,
                Cantidad_Estimada_Periodo: parseFloat(d.Cantidad_Estimada_Periodo) || 0,
                Importe_Este_Periodo: parseFloat(d.Importe_Este_Periodo) || 0,
                Avance_Periodo_Porcentaje: parseFloat(d.Avance_Periodo_Porcentaje) || 0,
                Avance_Acumulado_Porcentaje: parseFloat(d.Avance_Acumulado_Porcentaje) || 0,
                Importe_Acumulado: parseFloat(d.Importe_Acumulado) || 0
            };
        });

        return generarRespuesta(true, {
            estimacion: estimacion,
            contrato: contrato,
            conceptos: detalles
        });
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Actualiza las cantidades ejecutadas y recalcula los totales de la estimación.
 */
function actualizarDetalleEstimacion(idEstimacion, items) {
    const nombreFuncion = "actualizarDetalleEstimacion";
    try {
        if (!idEstimacion) throw new Error("ID de estimación no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 0. Obtener info de la estimación actual y el contrato
        const sheetEst = ss.getSheetByName('Estimaciones');
        const dataEst = getSafeData(sheetEst);
        const cabEst = dataEst[0];
        const idxIdEstCab = getColIndex(cabEst, 'ID_Estimacion');
        const rowIdxEst = dataEst.findIndex(r => String(r[idxIdEstCab]) === String(idEstimacion));
        if (rowIdxEst === -1) throw new Error("No se encontró la estimación.");

        const estRow = dataEst[rowIdxEst];
        const idContrato = estRow[getColIndex(cabEst, 'ID_Contrato')];
        const noEstActual = parseInt(estRow[getColIndex(cabEst, 'No_Estimacion')]) || 0;

        // 1. Obtener contrato y sus conceptos para importes totales
        const resContrato = getDetalleContrato(idContrato);
        if (!resContrato.success) throw new Error(resContrato.error);
        const contrato = resContrato.data;
        const pctAmort = parseFloat(contrato.Porcentaje_Amortizacion_Anticipo) || 0;
        const montoTotalContrato = parseFloat(contrato.Monto_Total_Sin_IVA) || 1;

        const resConceptos = getConceptosContrato(idContrato);
        const conceptosMap = {};
        if (resConceptos.success) {
            resConceptos.data.forEach(c => conceptosMap[String(c.ID_Concepto)] = parseFloat(c.Importe_Total_Sin_IVA) || 0);
        }

        // 2. Obtener TODAS las estimaciones previas del contrato para calcular acumulados
        const sheetDet = ss.getSheetByName('Detalle_Estimacion');
        const allDetalles = getSafeData(sheetDet);
        const cabDet = allDetalles[0];
        const idxIdEstDet = getColIndex(cabDet, 'ID_Estimacion');
        const idxIdConceptoDet = getColIndex(cabDet, 'ID_Concepto');
        const idxImporteDet = getColIndex(cabDet, 'Importe_Este_Periodo');

        // Mapear acumulados previos por concepto
        const acumuladosPrevios = {};
        const idsPrevios = dataEst.filter(r => String(r[getColIndex(cabEst, 'ID_Contrato')]) === String(idContrato)
            && parseInt(r[getColIndex(cabEst, 'No_Estimacion')]) < noEstActual)
            .map(r => String(r[idxIdEstCab]));

        allDetalles.forEach((row, i) => {
            if (i === 0) return;
            if (idsPrevios.includes(String(row[idxIdEstDet]))) {
                const idC = String(row[idxIdConceptoDet]);
                acumuladosPrevios[idC] = (acumuladosPrevios[idC] || 0) + (parseFloat(row[idxImporteDet]) || 0);
            }
        });

        // 3. Persistencia de Deltas (Eliminar anteriores e insertar solo los > 0 para esta estimación)
        const idxCantEst = getColIndex(cabDet, 'Cantidad_Estimada_Periodo');
        const idxPU = getColIndex(cabDet, 'Precio_Unitario_Contrato');
        const idxImporteActualCol = getColIndex(cabDet, 'Importe_Este_Periodo');
        const idxAcumuladoDet = getColIndex(cabDet, 'Importe_Acumulado');
        const idxPctAcumDet = getColIndex(cabDet, 'Avance_Acumulado_Porcentaje');

        // Determinar filas a eliminar de la estimación actual
        const filasABorrar = [];
        for (let i = allDetalles.length - 1; i >= 1; i--) {
            if (String(allDetalles[i][idxIdEstDet]) === String(idEstimacion)) {
                filasABorrar.push(i + 1);
            }
        }
        filasABorrar.forEach(row => sheetDet.deleteRow(row));

        let totalBrutoActual = 0;
        let totalAcumuladoAnterior = 0;

        // Sumar todos los acumulados previos para el total de la carátula
        Object.values(acumuladosPrevios).forEach(v => totalAcumuladoAnterior += v);

        // Obtener el PU de los conceptos para los nuevos cálculos
        const resCat = getConceptosContrato(idContrato);
        const catalogoMap = {};
        if (resCat.success) {
            resCat.data.forEach(c => catalogoMap[String(c.ID_Concepto)] = {
                pu: parseFloat(c.Precio_Unitario) || 0,
                total: parseFloat(c.Importe_Total_Sin_IVA) || 0
            });
        }

        const nuevasFilasDelta = [];
        const idxIdDet = getColIndex(cabDet, 'ID_Detalle');
        // Obtener el último ID global de la hoja después de borrar
        const dataVacia = getSafeData(sheetDet);
        let lastIdDet = dataVacia.length > 1 ? Math.max(...dataVacia.slice(1).map(r => parseInt(r[idxIdDet]) || 0)) : 0;

        items.forEach(it => {
            const cantActual = parseFloat(it.Cantidad_Estimada_Periodo) || 0;
            const idConcepto = String(it.ID_Concepto);
            const infoC = catalogoMap[idConcepto] || { pu: 0, total: 0 };
            const pu = infoC.pu;
            const importeActual = cantActual * pu;

            totalBrutoActual += importeActual;

            if (cantActual > 0) {
                const prev = acumuladosPrevios[idConcepto] || 0;
                const nuevoAcumulado = prev + importeActual;
                const pctAcumulado = (nuevoAcumulado / montoTotalContrato);
                const pctPeriodo = (importeActual / montoTotalContrato);

                lastIdDet++;
                const fila = [];
                cabDet.forEach(col => {
                    switch (col) {
                        case 'ID_Detalle': fila.push(lastIdDet); break;
                        case 'ID_Estimacion': fila.push(idEstimacion); break;
                        case 'ID_Concepto': fila.push(idConcepto); break;
                        case 'Cantidad_Estimada_Periodo': fila.push(cantActual); break;
                        case 'Precio_Unitario_Contrato': fila.push(pu); break;
                        case 'Importe_Este_Periodo': fila.push(importeActual); break;
                        case 'Importe_Acumulado': fila.push(nuevoAcumulado); break;
                        case 'Avance_Acumulado_Porcentaje': fila.push(pctAcumulado); break;
                        case 'Avance_Periodo_Porcentaje': fila.push(pctPeriodo); break;
                        default: fila.push(0); break;
                    }
                });
                nuevasFilasDelta.push(fila);
            }
        });

        if (nuevasFilasDelta.length > 0) {
            sheetDet.getRange(sheetDet.getLastRow() + 1, 1, nuevasFilasDelta.length, cabDet.length).setValues(nuevasFilasDelta);
        }

        // 4. Actualizar Tabla Estimaciones (Carátula)
        const iva = totalBrutoActual * 0.16;
        const subtotal = totalBrutoActual;
        const deduccionSFP = totalBrutoActual * 0.005; // 5 al millar
        const amortizacion = totalBrutoActual * (pctAmort / 100);

        const montoNeto = (totalBrutoActual + iva) - deduccionSFP - amortizacion;

        const updateMap = {
            'Monto_Bruto_Estimado': totalBrutoActual,
            'Deduccion_Surv_05_Monto': deduccionSFP,
            'Subtotal': subtotal,
            'IVA': iva,
            'Monto_Neto_A_Pagar': montoNeto,
            'Avance_Acumulado_Anterior': totalAcumuladoAnterior,
            'Avance_Actual': totalBrutoActual,
            'Estado_Validacion': 'VALIDADA',
            'Avance_Anterior_Porcentaje': (totalAcumuladoAnterior / montoTotalContrato),
            'Avance_Actual_Porcentaje': (totalBrutoActual / montoTotalContrato),
            'Avance_Acumulado_Porcentaje': ((totalAcumuladoAnterior + totalBrutoActual) / montoTotalContrato)
        };

        for (let key in updateMap) {
            let colIdx = getColIndex(cabEst, key);
            if (colIdx !== -1) {
                sheetEst.getRange(rowIdxEst + 1, colIdx + 1).setValue(updateMap[key]);
            }
        }

        return generarRespuesta(true, "Estimación actualizada con acumulados y retenciones.");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Obtiene la configuración del programa y sus periodos para un contrato.
 */
function getProgramaYPeriodos(idContrato) {
    const nombreFuncion = "getProgramaYPeriodos";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Obtener Programa (Nivel 1)
        const progs = dbSelect('Programa', { 'ID_Contrato': idContrato });
        let programa = progs.length > 0 ? progs[0] : null;

        // 2. Si no existe, crear uno por defecto
        if (!programa) {
            const detalle = getDetalleContrato(idContrato);
            programa = dbInsert('Programa', {
                'ID_Contrato': idContrato,
                'Tipo_Programa': 'MENSUAL',
                'Fecha_Inicio': detalle.success ? detalle.data.Fecha_Inicio_Obra : "",
                'Fecha_Termino': detalle.success ? detalle.data.Fecha_Fin_Obra : ""
            });
        }

        // 3. Obtener Periodos (Nivel 2)
        const periodos = dbSelect('Programa_Periodo', { 'ID_Numero_Programa': programa.ID_Numero_Programa });

        // 4. Verificar si tiene datos registrados en Nivel 3 (para bloqueo de edición)
        let hasData = false;
        const sheetEjec = ss.getSheetByName('Programa_Ejecucion');
        if (sheetEjec) {
            const dataEjec = getSafeData(sheetEjec);
            if (dataEjec.length > 1) {
                const cabEjec = dataEjec[0];
                const idxIdProgPer = getColIndex(cabEjec, 'ID_Programa_Periodo');
                const idxMonto = getColIndex(cabEjec, 'Monto_Programado');
                const perIds = new Set(periodos.map(p => String(p.ID_Programa_Periodo)));

                for (let i = 1; i < dataEjec.length; i++) {
                    if (idxIdProgPer !== -1 && perIds.has(String(dataEjec[i][idxIdProgPer]))) {
                        const monto = parseFloat(dataEjec[i][idxMonto]) || 0;
                        if (monto > 0) {
                            hasData = true;
                            break;
                        }
                    }
                }
            }
        }

        return generarRespuesta(true, {
            programa: programa,
            periodos: periodos.sort((a, b) => (parseInt(a.Numero_Periodo) || 0) - (parseInt(b.Numero_Periodo) || 0)),
            hasData: hasData
        });
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Guarda la configuración del programa y actualiza/reemplaza los periodos.
 */
function guardarConfiguracionPrograma(idContrato, config) {
    const nombreFuncion = "guardarConfiguracionPrograma";
    try {
        if (!idContrato) throw new Error("ID de contrato no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Actualizar Nivel 1 (Programa)
        const progs = dbSelect('Programa', { 'ID_Contrato': idContrato });
        if (progs.length === 0) throw new Error("No existe un programa base para este contrato.");
        const idProg = progs[0].ID_Numero_Programa;

        dbUpdate('Programa', {
            'Tipo_Programa': config.programa.tipo,
            'Fecha_Inicio': config.programa.fechaInicio,
            'Fecha_Termino': config.programa.fechaTermino
        }, { 'ID_Numero_Programa': idProg });

        // 2. Actualizar Nivel 2 (Periodos)
        // Por simplicidad en la sincronización, borramos los periodos que no tengan datos en Nivel 3 y reinsertamos.
        // Pero el requerimiento suele ser sobreescribir.

        // Borrar periodos actuales del programa (cascada controlada)
        dbDelete('Programa_Periodo', { 'ID_Numero_Programa': idProg });

        // Insertar nuevos periodos
        config.periodos.forEach((p, index) => {
            dbInsert('Programa_Periodo', {
                'ID_Numero_Programa': idProg,
                'Numero_Periodo': index + 1,
                'Periodo': "'" + p.label, // Forzar texto
                'Fecha_Inicio': p.fechaInicio,
                'Fecha_Termino': p.fechaTermino
            });
        });

        return generarRespuesta(true, "Configuración de programa guardada exitosamente.");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Elimina un periodo específico del programa, verificando que no tenga datos en Nivel 3.
 */
function eliminarPeriodo(idPeriodo) {
    const nombreFuncion = "eliminarPeriodo";
    try {
        if (!idPeriodo) throw new Error("ID de periodo no proporcionado.");
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Verificar si existen datos en Programa_Ejecucion para este periodo
        const sheetEjec = ss.getSheetByName('Programa_Ejecucion');
        if (sheetEjec) {
            const dataEjec = getSafeData(sheetEjec);
            if (dataEjec.length > 1) {
                const cabEjec = dataEjec[0];
                const idxIdProgPer = getColIndex(cabEjec, 'ID_Programa_Periodo');
                const idxMonto = getColIndex(cabEjec, 'Monto_Programado');

                for (let i = 1; i < dataEjec.length; i++) {
                    if (String(dataEjec[i][idxIdProgPer]) === String(idPeriodo)) {
                        const monto = parseFloat(dataEjec[i][idxMonto]) || 0;
                        if (monto > 0) {
                            throw new Error("No se puede eliminar un periodo que ya tiene montos programados en la matriz.");
                        }
                    }
                }
            }
        }

        // 2. Proceder a la eliminación en Programa_Periodo
        dbDelete('Programa_Periodo', { 'ID_Programa_Periodo': idPeriodo });

        return generarRespuesta(true, "Periodo eliminado correctamente.");
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Guarda o actualiza un registro de validación de archivo para una estimación.
 * Si ya existe un registro para el mismo ID_Estimacion + Tipo_Archivo, lo actualiza.
 * @param {string} idEstimacion - ID de la estimación
 * @param {string} tipoArchivo - 'CARATULA_GENERADOR' | 'FACTURA_BORRADOR' | 'FACTURA_TIMBRADA'
 * @param {string} estadoValidacion - 'APROBADO' | 'CON_ERRORES' | 'PENDIENTE'
 * @param {Array} checklistArray - Array de objetos {rubro, aprobado, observacion}
 * @param {string} observacionesResumen - Texto libre con resumen
 */
function guardarValidacionArchivo(idEstimacion, tipoArchivo, estadoValidacion, checklistArray, observacionesResumen) {
    const nombreFuncion = "guardarValidacionArchivo";
    try {
        if (!idEstimacion) throw new Error("ID de estimación no proporcionado.");
        if (!tipoArchivo) throw new Error("Tipo de archivo no proporcionado.");

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName('Validacion_Archivos');
        if (!sheet) {
            sheet = ss.insertSheet('Validacion_Archivos');
            sheet.appendRow(ESQUEMA_BD['Validacion_Archivos']);
        }

        const data = getSafeData(sheet);
        const cabeceras = data[0];
        const idxId = getColIndex(cabeceras, 'ID_Validacion');
        const idxIdEst = getColIndex(cabeceras, 'ID_Estimacion');
        const idxTipo = getColIndex(cabeceras, 'Tipo_Archivo');
        const idxFecha = getColIndex(cabeceras, 'Fecha_Carga');
        const idxEstado = getColIndex(cabeceras, 'Estado_Validacion');
        const idxChecklist = getColIndex(cabeceras, 'Checklist_JSON');
        const idxObs = getColIndex(cabeceras, 'Observaciones_Resumen');

        const checklistJSON = JSON.stringify(checklistArray || []);
        const ahora = new Date();

        // Buscar registro existente para esta estimación + tipo
        let filaExistente = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdEst]) === String(idEstimacion) && String(data[i][idxTipo]) === String(tipoArchivo)) {
                filaExistente = i + 1;
                break;
            }
        }

        if (filaExistente !== -1) {
            // Actualizar registro existente
            sheet.getRange(filaExistente, idxFecha + 1).setValue(ahora);
            sheet.getRange(filaExistente, idxEstado + 1).setValue(estadoValidacion);
            sheet.getRange(filaExistente, idxChecklist + 1).setValue(checklistJSON);
            sheet.getRange(filaExistente, idxObs + 1).setValue(observacionesResumen || '');
        } else {
            // Crear nuevo registro
            const nuevoId = "VAL-" + Utilities.getUuid().substring(0, 8).toUpperCase();
            sheet.appendRow([
                nuevoId,
                String(idEstimacion),
                tipoArchivo,
                ahora,
                estadoValidacion,
                checklistJSON,
                observacionesResumen || ''
            ]);
        }

        return generarRespuesta(true, { idEstimacion, tipoArchivo, estado: estadoValidacion });
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}

/**
 * Obtiene todos los registros de validación para una estimación.
 * @param {string} idEstimacion - ID de la estimación
 * @returns {Object} Array de registros de validación con checklist parseado
 */
function getValidacionesEstimacion(idEstimacion) {
    const nombreFuncion = "getValidacionesEstimacion";
    try {
        if (!idEstimacion) throw new Error("ID de estimación no proporcionado.");

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Validacion_Archivos');
        if (!sheet) return generarRespuesta(true, []);

        const data = getSafeData(sheet);
        if (data.length <= 1) return generarRespuesta(true, []);

        const cabeceras = data[0];
        const idxIdEst = getColIndex(cabeceras, 'ID_Estimacion');
        const resultados = [];

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxIdEst]) === String(idEstimacion)) {
                let obj = {};
                for (let j = 0; j < cabeceras.length; j++) {
                    let val = data[i][j];
                    if (val instanceof Date) {
                        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                    }
                    obj[cabeceras[j]] = val;
                }
                // Parsear Checklist_JSON a objeto
                try {
                    obj.Checklist = JSON.parse(obj.Checklist_JSON || '[]');
                } catch (e) {
                    obj.Checklist = [];
                }
                resultados.push(obj);
            }
        }

        return generarRespuesta(true, resultados);
    } catch (error) {
        return generarRespuesta(false, error.message, nombreFuncion);
    }
}
