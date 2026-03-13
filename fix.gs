/**
 * ARCHIVO DE MANTENIMIENTO Y EXTRAPOLACIÓN (fix.gs)
 * Este archivo contiene utilidades para migrar y normalizar datos al esquema de 3 niveles.
 */

/**
 * Inicializa las tablas del Programa de 3 Niveles si no existen.
 */
function fix_inicializarTablasPrograma() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tablas = ['Programa', 'Programa_Periodo', 'Programa_Ejecucion'];

    tablas.forEach(nombre => {
        let sheet = ss.getSheetByName(nombre);
        if (!sheet) {
            sheet = ss.insertSheet(nombre);
            sheet.appendRow(ESQUEMA_BD[nombre]);
            console.log(`Hoja creada: ${nombre}`);
        } else {
            console.log(`La hoja ${nombre} ya existe.`);
        }
    });
}

/**
 * EXTRAPOLACIÓN COMPLETA:
 * 1. Crea registros en 'Programa' (Nivel 1) basados en 'Contratos'.
 * 2. Analiza 'Programa_Ejecucion' para extraer periodos únicos y crear 'Programa_Periodo' (Nivel 2).
 * 3. Actualiza 'Programa_Ejecucion' (Nivel 3) vinculando los IDs de Programa y Periodo.
 */
function fix_extrapolarEstructuraProgramaCompleto() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombreFuncion = "fix_extrapolarEstructuraProgramaCompleto";

    try {
        // Asegurar que las tablas existan
        fix_inicializarTablasPrograma();

        const sheetContratos = ss.getSheetByName('Contratos');
        const sheetCatalogo = ss.getSheetByName('Catalogo_Conceptos');
        const sheetProg = ss.getSheetByName('Programa');
        const sheetPer = ss.getSheetByName('Programa_Periodo');
        const sheetEjec = ss.getSheetByName('Programa_Ejecucion');

        const dataContratos = getSafeData(sheetContratos);
        const cabC = dataContratos[0];
        const idxC_Id = getColIndex(cabC, 'ID_Contrato');
        const idxC_Ini = getColIndex(cabC, 'Fecha_Inicio_Obra');
        const idxC_Fin = getColIndex(cabC, 'Fecha_Fin_Obra');

        const dataCat = getSafeData(sheetCatalogo);
        const cabCat = dataCat[0];
        const idxCat_Id = getColIndex(cabCat, 'ID_Concepto');
        const idxCat_Cont = getColIndex(cabCat, 'ID_Contrato');

        const dataEjecRaw = getSafeData(sheetEjec);
        const cabEjec = dataEjecRaw[0];
        const idxEj_IdProg = getColIndex(cabEjec, 'ID_Programa');
        const idxEj_IdConc = getColIndex(cabEjec, 'ID_Concepto');
        const idxEj_FIni = getColIndex(cabEjec, 'Fecha_Inicio');
        const idxEj_FFin = getColIndex(cabEjec, 'Fecha_Fin');
        const idxEj_IdPer = getColIndex(cabEjec, 'ID_Programa_Periodo');

        // Contadores para IDs
        let nextProgId = sheetProg.getLastRow() > 1 ? (Math.max(...sheetProg.getRange(2, 1, sheetProg.getLastRow() - 1, 1).getValues().flat()) + 1) : 1;
        let nextPerId = sheetPer.getLastRow() > 1 ? (Math.max(...sheetPer.getRange(2, 1, sheetPer.getLastRow() - 1, 1).getValues().flat()) + 1) : 1;

        console.log("Iniciando análisis de contratos...");

        for (let i = 1; i < dataContratos.length; i++) {
            const idContrato = dataContratos[i][idxC_Id];
            if (!idContrato) continue;

            // A. Gestionar Programa (Nivel 1)
            let idProgActual = null;
            const dataProgCurrent = getSafeData(sheetProg);
            const cabProgCurrent = dataProgCurrent.length > 0 ? dataProgCurrent[0] : [];
            const progExistente = dataProgCurrent.find(r => String(r[getColIndex(cabProgCurrent, 'ID_Contrato')]) === String(idContrato));

            if (progExistente) {
                idProgActual = progExistente[0];
            } else {
                idProgActual = nextProgId++;
                sheetProg.appendRow([
                    idProgActual,
                    idContrato,
                    'ORIGINAL',
                    dataContratos[i][idxC_Ini],
                    dataContratos[i][idxC_Fin]
                ]);
                console.log(`Creado Programa ID ${idProgActual} para Contrato ${idContrato}`);
            }

            // B. Obtener conceptos de este contrato
            const idsConceptosContrato = dataCat
                .filter(r => String(r[idxCat_Cont]) === String(idContrato))
                .map(r => String(r[idxCat_Id]));

            // C. Analizar registros de ejecución existentes para estos conceptos
            const registrosRelacionados = dataEjecRaw.filter((r, idx) => {
                if (idx === 0) return false;
                return idsConceptosContrato.includes(String(r[idxEj_IdConc]));
            });

            if (registrosRelacionados.length === 0) {
                console.log(`Contrato ${idContrato} no tiene registros de ejecución previos.`);
                continue;
            }

            // D. Identificar periodos únicos en la ejecución actual
            const periodosUnicos = [];
            registrosRelacionados.forEach(reg => {
                const fIni = reg[idxEj_FIni];
                const fFin = reg[idxEj_FFin];
                if (!fIni || !fFin) return;

                const key = `${new Date(fIni).getTime()}_${new Date(fFin).getTime()}`;
                if (!periodosUnicos.find(p => p.key === key)) {
                    periodosUnicos.push({ key, fIni, fFin });
                }
            });

            // E. Crear registros en 'Programa_Periodo' (Nivel 2)
            const mapaPeriodoId = {};
            const dataPerCurrent = getSafeData(sheetPer);
            const cabPer = dataPerCurrent[0];
            let numPerContador = 1;

            periodosUnicos.forEach(pu => {
                const perExistente = dataPerCurrent.find(r =>
                    String(r[getColIndex(cabPer, 'ID_Numero_Programa')]) === String(idProgActual) &&
                    new Date(r[getColIndex(cabPer, 'Fecha_Inicio')]).getTime() === new Date(pu.fIni).getTime() &&
                    new Date(r[getColIndex(cabPer, 'Fecha_Termino')]).getTime() === new Date(pu.fFin).getTime()
                );

                let idPer;
                if (perExistente) {
                    idPer = perExistente[0];
                } else {
                    idPer = nextPerId++;
                    const mesNombre = Utilities.formatDate(new Date(pu.fIni), "GMT", "MMMM yyyy").toUpperCase();
                    sheetPer.appendRow([
                        idPer,
                        idProgActual,
                        numPerContador++,
                        mesNombre,
                        pu.fIni,
                        pu.fFin
                    ]);
                }
                mapaPeriodoId[pu.key] = idPer;
            });

            // F. Actualizar 'Programa_Ejecucion' con los nuevos IDs (Nivel 3)
            // Se realiza directo en la hoja para seguridad sobre los registros que coincidan
            for (let j = 1; j < dataEjecRaw.length; j++) {
                if (idsConceptosContrato.includes(String(dataEjecRaw[j][idxEj_IdConc]))) {
                    const fIni = dataEjecRaw[j][idxEj_FIni];
                    const fFin = dataEjecRaw[j][idxEj_FFin];
                    if (!fIni || !fFin) continue;

                    const key = `${new Date(fIni).getTime()}_${new Date(fFin).getTime()}`;
                    const idPeriodoEncontrado = mapaPeriodoId[key];

                    if (idPeriodoEncontrado) {
                        sheetEjec.getRange(j + 1, idxEj_IdProg + 1).setValue(idProgActual);
                        sheetEjec.getRange(j + 1, idxEj_IdPer + 1).setValue(idPeriodoEncontrado);
                    }
                }
            }
            console.log(`Extrapolación finalizada para Contrato ${idContrato}`);
        }

        console.log("PROCESO DE EXTRAPOLACIÓN FINALIZADO CON ÉXITO.");

    } catch (e) {
        console.error("Error en " + nombreFuncion + ": " + e.message);
    }
}

/**
 * SINCRONIZACIÓN UNIVERSAL DE ESQUEMA:
 * - Si la hoja no existe: La crea, pone encabezados, convierte en Tabla y renombra la Tabla.
 * - Si existe: Asegura que tenga todas las columnas y el formato profesional.
 */
function fix_sincronizarTodasLasTablas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Iniciando sincronización universal con formato de tablas oficiales...");

    const tablasEsquema = Object.keys(ESQUEMA_BD);

    tablasEsquema.forEach(nombreTabla => {
        let sheet = ss.getSheetByName(nombreTabla);

        // 1. Crear hoja si no existe
        if (!sheet) {
            console.log(`[PASO 1] Creando hoja: ${nombreTabla}`);
            sheet = ss.insertSheet(nombreTabla);

            // 2. Crear encabezados
            const headers = ESQUEMA_BD[nombreTabla];
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            console.log(`[PASO 2] Encabezados creados para ${nombreTabla}`);

            // 3. Convertir en Tabla y renombrar
            const ultimaCol = headers.length;
            const rangoInicial = sheet.getRange(1, 1, 2, ultimaCol); // 2 filas para tabla

            try {
                // Intentar usar la API de Tablas Oficial (Formatted Tables)
                if (typeof sheet.createTable === 'function') {
                    const nuevaTabla = sheet.createTable(rangoInicial, true);
                    nuevaTabla.setName(nombreTabla);
                    console.log(`[PASO 3] Convertida a Tabla Oficial y nombrada como: ${nombreTabla}`);
                } else {
                    // Fallback: Named Range + Filtro + Banding
                    ss.setNamedRange(nombreTabla, rangoInicial);
                    rangoInicial.createFilter();
                    rangoInicial.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
                    console.log(`[PASO 3 - Fallback] Tabla simulada con Named Range: ${nombreTabla}`);
                }
            } catch (e) {
                console.warn(`Aviso en creación de tabla ${nombreTabla}: ${e.message}`);
            }
        } else {
            // Sincronizar columnas para hojas ya existentes
            console.log(`[SINC] Verificando columnas para ${nombreTabla}`);
            const cabeceraObjetivo = ESQUEMA_BD[nombreTabla];
            let cabeceraActual = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];

            cabeceraObjetivo.forEach(col => {
                if (!cabeceraActual.includes(col)) {
                    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
                    console.log(`[${nombreTabla}] Columna añadida: ${col}`);
                }
            });
        }

        // 4. Estilos Finales (Siempre se aplican para mantener consistencia)
        const uCol = sheet.getLastColumn();
        const encabezado = sheet.getRange(1, 1, 1, uCol);

        encabezado.setBackground('#1e293b') // Slate-800
            .setFontColor('#f8fafc') // Slate-50
            .setFontWeight('bold')
            .setHorizontalAlignment('center');

        sheet.setFrozenRows(1);
        sheet.autoResizeColumns(1, uCol);
    });

    console.log("PROCESO DE SINCRONIZACIÓN FINALIZADO.");
}

/**
 * Utility to clean up records that were not properly linked to a Contract or Program.
 * Uses RELACIONES_BD to walk through the hierarchy and ensure integrity without deleting valid data.
 */
function depurarDatosHuerfanos() {
    const SS = SpreadsheetApp.getActiveSpreadsheet();
    let totalBorrados = 0;
    const reporte = [];

    try {
        // Cache valid IDs from parent tables to optimize lookups
        const cacheIds = {};
        const getValidIds = (tableName) => {
            if (cacheIds[tableName]) return cacheIds[tableName];

            const sheet = SS.getSheetByName(tableName);
            if (!sheet) return new Set();

            const data = getSafeData(sheet);
            if (data.length <= 1) return new Set();

            const pkName = ESQUEMA_BD[tableName][0];
            const idx = data[0].indexOf(pkName);
            if (idx === -1) return new Set();

            const ids = new Set();
            for (let i = 1; i < data.length; i++) {
                const val = String(data[i][idx]).trim();
                if (val && val !== "null" && val !== "undefined") {
                    ids.add(val);
                }
            }
            cacheIds[tableName] = ids;
            return ids;
        };

        // Order of cleanup: Level 1 -> Level 2 -> Level 3
        const niveles = [
            ['Catalogo_Conceptos', 'Programa', 'Estimaciones', 'Anticipos', 'Convenios_Modificatorios'],
            ['Matriz_Insumos', 'Programa_Periodo', 'Detalle_Estimacion', 'Deducciones_Retenciones', 'Facturas', 'Pagos_Emitidos', 'Validacion_Archivos'],
            ['Programa_Ejecucion']
        ];

        niveles.forEach((nivel, index) => {
            nivel.forEach(tabla => {
                const rel = RELACIONES_BD[tabla];
                if (!rel) return;

                const sheet = SS.getSheetByName(tabla);
                if (!sheet) return;

                const data = getSafeData(sheet);
                if (data.length <= 1) return;

                const headers = data[0];
                const reglas = Array.isArray(rel) ? rel : [rel];

                let count = 0;
                // Reverse iteration to allow row deletion
                for (let i = data.length - 1; i >= 1; i--) {
                    let esHuerfano = false;
                    let razon = "";

                    for (const regla of reglas) {
                        const fkIdx = headers.indexOf(regla.fk);
                        if (fkIdx === -1) continue;

                        const val = String(data[i][fkIdx]).trim();
                        const validParents = getValidIds(regla.padre);

                        // CRITICAL FIX: Only delete if the FK is present BUT the parent target is missing.
                        // If FK is empty, we only delete if it's a Level 1 link to Contract (mandatory).
                        if (!val || val === "null" || val === "undefined" || val === "0") {
                            if (index === 0) { // Nivel 1 mandatory links
                                esHuerfano = true;
                                razon = `FK ${regla.fk} vacía en Nivel 1`;
                            }
                        } else if (!validParents.has(val)) {
                            esHuerfano = true;
                            razon = `Padre ${val} no encontrado en ${regla.padre}`;
                        }

                        if (esHuerfano) break;
                    }

                    if (esHuerfano) {
                        sheet.deleteRow(i + 1);
                        count++;
                    }
                }

                if (count > 0) {
                    totalBorrados += count;
                    reporte.push(`${tabla}: ${count}`);
                    // IMPORTANT: If we deleted rows in a level that could be a parent, reset cache
                    delete cacheIds[tabla];
                }
            });
        });

        return generarRespuesta(true, {
            total: totalBorrados,
            detalle: reporte.join(", ") || "Base de datos integra"
        });

    } catch (e) {
        console.error("Error en depurarDatosHuerfanos:", e);
        return generarRespuesta(false, e.toString());
    }
}

/**
 * Fusiona registros duplicados en Convenios_Recurso basados en la normalización del Numero_Acuerdo.
 * Busca coincidencias ignorando guiones, espacios y puntuación, manteniendo el registro más completo.
 */
function fix_limpiarDuplicadosConvenios() {
    const SS = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SS.getSheetByName('Convenios_Recurso');
    if (!sheet) return "Hoja Convenios_Recurso no encontrada";

    const data = getSafeData(sheet);
    if (data.length <= 1) return "No hay datos para procesar";

    const headers = data[0];
    const idxNum = headers.indexOf('Numero_Acuerdo');
    if (idxNum === -1) return "Columna Numero_Acuerdo no encontrada";
    
    const visitados = new Map(); // normalized -> index en el array 'data'
    const filasABorrar = [];
    let fusionados = 0;

    for (let i = 1; i < data.length; i++) {
        const numOriginal = data[i][idxNum];
        if (!numOriginal) continue;

        const numNorm = normalizeText(numOriginal);
        
        if (visitados.has(numNorm)) {
            const indexOriginal = visitados.get(numNorm);
            const rowBase = data[indexOriginal];
            const rowDuplicada = data[i];
            
            // Fusionar datos: Si el registro base tiene un campo vacío y el duplicado no, lo llenamos.
            headers.forEach((h, colIdx) => {
                // No tocar IDs ni el Numero_Acuerdo (el base se queda como está)
                if (h.startsWith('ID_') || h === 'Numero_Acuerdo') return;
                
                const valBase = rowBase[colIdx];
                const valDup = rowDuplicada[colIdx];

                if ((valBase === "" || valBase === null || valBase === undefined) && (valDup !== "" && valDup !== null && valDup !== undefined)) {
                    rowBase[colIdx] = valDup;
                    sheet.getRange(indexOriginal + 1, colIdx + 1).setValue(valDup);
                }
            });
            
            filasABorrar.push(i + 1);
            fusionados++;
            console.log(`[FIX] Fusionando duplicado: ${numOriginal} (Fila ${i+1}) -> Base: ${rowBase[idxNum]}`);
        } else {
            visitados.set(numNorm, i);
        }
    }

    // Borrar de abajo hacia arriba para no alterar índices de fila
    filasABorrar.reverse().forEach(row => sheet.deleteRow(row));
    
    console.log(`Limpieza finalizada. Registros fusionados: ${fusionados}`);
    return `Limpieza exitosa: se fusionaron ${fusionados} registros duplicados en Convenios_Recurso.`;
}
