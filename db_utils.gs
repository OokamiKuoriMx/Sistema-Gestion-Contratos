/**
 * DB_UTILS.GS - Capa de Abstracción de Datos (SGC)
 * Proporciona funciones genéricas para interactuar con las hojas de cálculo.
 */

/**
 * Función de utilidad para obtener datos de una hoja con reintentos en caso de errores transitorios.
 */
function getSafeData(sheet) {
    let retries = 3;
    while (retries > 0) {
        try {
            const lastRow = sheet.getLastRow();
            if (lastRow === 0) return [];
            const lastCol = sheet.getLastColumn();
            if (lastCol === 0) return [];
            return sheet.getRange(1, 1, lastRow, lastCol).getValues();
        } catch (e) {
            retries--;
            if (retries === 0) throw e;
            Utilities.sleep(500); // Esperar medio segundo antes de reintentar
        }
    }
}

/**
 * Selecciona registros de una tabla basados en condiciones.
 * @param {string} tabla Nombre de la hoja.
 * @param {Object} condiciones Objeto con pares columna:valor para filtrar.
 * @return {Array<Object>} Lista de registros que coinciden.
 */
function dbSelect(tabla, condiciones = {}) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(tabla);
        if (!sheet) return [];

        const data = getSafeData(sheet);
        if (data.length <= 1) return [];

        const headers = data[0];
        const results = [];

        // Cachear índices de columnas para los filtros
        const fieldIndices = {};
        for (let key in condiciones) {
            fieldIndices[key] = getColIndex(headers, key);
        }

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const record = {};

            headers.forEach((h, idx) => {
                if (h) {
                    let val = row[idx];
                    // Asegurar que las fechas sean serializables como strings para evitar errores INTERNAL en el retorno
                    if (val instanceof Date) {
                        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T12:00:00'");
                    }
                    record[h] = val;
                }
            });

            // Aplicar filtros optimizados
            let match = true;
            for (let key in condiciones) {
                const idxCol = fieldIndices[key];
                if (idxCol !== -1) {
                    if (String(record[key]) !== String(condiciones[key])) {
                        match = false;
                        break;
                    }
                }
            }

            if (match) results.push(record);
        }
        return results;
    } catch (e) {
        console.error("Error en dbSelect (tabla: " + tabla + "):", e);
        return [];
    }
}

/**
 * Inserta un nuevo registro en una tabla.
 * @param {string} tabla Nombre de la hoja.
 * @param {Object} datos Objeto con los datos a insertar.
 * @return {Object|null} El registro insertado con su ID asignado.
 */
function dbInsert(tabla, datos) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(tabla);
        if (!sheet) {
            // Si no existe, intentar crearla usando el esquema global
            if (typeof ESQUEMA_BD !== 'undefined' && ESQUEMA_BD[tabla]) {
                sheet = ss.insertSheet(tabla);
                sheet.appendRow(ESQUEMA_BD[tabla]);
            } else {
                throw new Error("La tabla " + tabla + " no existe y no tiene esquema definido.");
            }
        }

        const data = getSafeData(sheet);
        const headers = data.length > 0 ? data[0] : ESQUEMA_BD[tabla];
        const pkName = headers[0];

        // Generar ID si falta (Consecutivo simple para IDs numéricos o UUID corto)
        if (!datos[pkName]) {
            const lastRow = data.length; // Usar data[0] + rows
            if (lastRow > 1) {
                const lastId = data[data.length - 1][0];
                if (typeof lastId === 'number') {
                    datos[pkName] = lastId + 1;
                } else {
                    datos[pkName] = (tabla.substring(0, 3).toUpperCase()) + "-" + Utilities.getUuid().substring(0, 8).toUpperCase();
                }
            } else {
                datos[pkName] = 1;
            }
        }

        const row = headers.map(h => datos[h] !== undefined ? datos[h] : "");
        sheet.appendRow(row);
        SpreadsheetApp.flush(); // Garantizar consistencia

        return datos;
    } catch (e) {
        console.error("Error en dbInsert (tabla: " + tabla + "):", e);
        throw e;
    }
}

/**
 * Actualiza registros existentes en una tabla.
 * @param {string} tabla Nombre de la hoja.
 * @param {Object} datos Datos para actualizar.
 * @param {Object} condiciones Filtros para identificar registros.
 */
function dbUpdate(tabla, datos, condiciones) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(tabla);
        if (!sheet) return;

        const allData = getSafeData(sheet);
        if (allData.length <= 1) return;
        const headers = allData[0];

        // Cachear índices de columnas para las condiciones
        const condIndices = {};
        for (let key in condiciones) {
            condIndices[key] = getColIndex(headers, key);
        }

        // Cachear índices para los datos a actualizar
        const updateIndices = {};
        for (let key in datos) {
            updateIndices[key] = getColIndex(headers, key);
        }

        for (let i = 1; i < allData.length; i++) {
            let match = true;
            for (let key in condiciones) {
                const idx = condIndices[key];
                if (idx !== -1 && String(allData[i][idx]) !== String(condiciones[key])) {
                    match = false;
                    break;
                }
            }

            if (match) {
                // Actualizar solo los campos proporcionados
                for (let key in datos) {
                    const idx = updateIndices[key];
                    if (idx !== -1) {
                        sheet.getRange(i + 1, idx + 1).setValue(datos[key]);
                    }
                }
            }
        }
        SpreadsheetApp.flush();
    } catch (e) {
        console.error("Error en dbUpdate (tabla: " + tabla + "):", e);
        throw e;
    }
}

/**
 * Elimina registros de una tabla basados en condiciones.
 * @param {string} tabla Nombre de la hoja.
 * @param {Object} condiciones Filtros para identificar registros a eliminar.
 */
function dbDelete(tabla, condiciones) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(tabla);
        if (!sheet) return;

        const data = getSafeData(sheet);
        if (data.length <= 1) return;
        const headers = data[0];

        const condIndices = {};
        for (let key in condiciones) {
            condIndices[key] = getColIndex(headers, key);
        }

        for (let i = data.length - 1; i >= 1; i--) {
            let match = true;
            for (let key in condiciones) {
                const idx = condIndices[key];
                if (idx !== -1 && String(data[i][idx]) !== String(condiciones[key])) {
                    match = false;
                    break;
                }
            }

            if (match) {
                sheet.deleteRow(i + 1);
            }
        }
        SpreadsheetApp.flush();
    } catch (e) {
        console.error("Error en dbDelete (tabla: " + tabla + "):", e);
        throw e;
    }
}
