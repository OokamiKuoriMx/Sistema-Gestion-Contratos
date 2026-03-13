# Instrucciones mejoradas de importación IA (SGC)

## Objetivo
Estandarizar la carga de documentos (`Contrato`, `CAF/Suficiencia`, `Programa`, `APU`) para que el sistema:

1. **Actualice cuando exista** (sin duplicar).
2. **Inserte cuando no exista**.
3. **Vincule correctamente IDs padre-hijo** (`ID_Contrato`, `ID_Contratista`, `ID_Numero_Programa`, `ID_Programa_Periodo`, `ID_Concepto`).
4. Evite registros vacíos y huérfanos.

---

## Regla principal (SIEMPRE)
Para cada entidad extraída por IA:

1. **Normalizar datos** (trim, mayúsculas en `Periodo`, fechas y números).
2. **Buscar existencia** con llave primaria o llave natural definida.
3. **Si existe:** `UPDATE` solo de campos con valor válido.
4. **Si no existe:** `INSERT` completo.
5. **Guardar ID real** para enlazar hijos.

> Nunca insertar por defecto sin hacer matching previo.

---

## Orden obligatorio de procesamiento (pipeline)

## 1) Contrato

### 1.1 Contratista (primero)
- Llave natural recomendada: `RFC`.
- Si existe RFC: actualizar datos del contratista.
- Si no existe: insertar contratista.
- Guardar `ID_Contratista` real.

### 1.2 Contrato (después)
- Llave natural recomendada: `Numero_Contrato`.
- Si existe contrato: actualizar columnas que cambiaron.
- Si no existe: insertar contrato.
- **Forzar vínculo:** `Contratos.ID_Contratista = ID_Contratista` resuelto en 1.1.
- No insertar si el registro viene vacío.

---

## 2) CAF / Suficiencia / Convenio de recurso
- Llave natural recomendada: `Numero_Acuerdo`.
- Si existe: actualizar.
- Si no existe: insertar.
- Tomar `ID_Convenio` y vincular en contrato: `Contratos.ID_Convenio_Vinculado`.

---

## 3) Programa

### 3.1 Programa (cabecera)
- Llave natural recomendada: `ID_Contrato` (y opcionalmente `Tipo_Programa`).
- Si existe: actualizar.
- Si no existe: insertar.
- Guardar `ID_Numero_Programa`.

### 3.2 Periodos (`Programa_Periodo`)
**Problema que se debe evitar:** duplicar periodos por nombre.

Usar matching en este orden:
1. `ID_Numero_Programa + Numero_Periodo` (principal).
2. `ID_Numero_Programa + Periodo` normalizado (fallback).

Reglas:
- `Periodo` debe guardarse en **MAYÚSCULAS** y sin comillas iniciales.
- Si existe: actualizar fechas y etiqueta.
- Si no existe: insertar.
- No borrar periodos con dependencias en `Programa_Ejecucion`.

### 3.3 Programa de ejecución (`Programa_Ejecucion`)
Matching por:
- `ID_Concepto + ID_Programa_Periodo`.

Reglas:
- Si existe: actualizar `Monto_Programado` / `%`.
- Si no existe: insertar.
- Evitar crear nuevos periodos si ya existe uno equivalente.

---

## 4) APU (Matriz de insumos)

### 4.1 Conceptos (`Catalogo_Conceptos`)
Matching recomendado:
1. `ID_Contrato + Clave`.
2. `ID_Contrato + Descripcion normalizada` (fallback).

### 4.2 Insumos (`Matriz_Insumos`)
Matching recomendado:
- `ID_Concepto + Clave_Insumo`.

Reglas:
- Resolver primero `ID_Concepto` del concepto padre.
- Si existe insumo: actualizar.
- Si no existe: insertar.

---

## Llaves naturales recomendadas por tabla
- `Contratistas`: `RFC`
- `Contratos`: `Numero_Contrato`
- `Convenios_Recurso`: `Numero_Acuerdo`
- `Programa`: `ID_Contrato` (+ `Tipo_Programa` opcional)
- `Programa_Periodo`: `ID_Numero_Programa + Numero_Periodo`
- `Programa_Ejecucion`: `ID_Concepto + ID_Programa_Periodo`
- `Catalogo_Conceptos`: `ID_Contrato + Clave`
- `Matriz_Insumos`: `ID_Concepto + Clave_Insumo`

---

## Reglas de calidad de datos
1. No procesar objetos vacíos.
2. No crear contratos automáticamente desde documentos que no sean de tipo `CONTRATO`.
3. Normalizar:
   - `Periodo` a mayúsculas.
   - Strings con `trim()`.
   - Números con `parseFloat`.
   - Fechas en formato consistente.
4. No sobreescribir con vacío cuando ya existe dato bueno.

---

## Manejo de documentos largos (2 pasadas)
Cuando el documento exceda tiempo de sesión:

### Pasada 1
- Extraer/guardar cabeceras y entidades padre.
- Guardar checkpoint: página o bloque procesado.

### Pasada 2
- Continuar desde checkpoint.
- Cargar contexto de BD (periodos/IDs/conceptos existentes) para hacer matching y evitar duplicados.

> Debe ser idempotente: correr dos veces no debe duplicar datos.

---

## Checklist operativo por corrida
- [ ] Se resolvió `ID_Contratista` antes de guardar `Contratos`.
- [ ] `Contratos.ID_Contratista` quedó vinculado.
- [ ] Se resolvió `ID_Numero_Programa` antes de periodos.
- [ ] `Programa_Periodo` hizo match por número y no duplicó.
- [ ] `Periodo` quedó en mayúsculas.
- [ ] `Programa_Ejecucion` actualizó por `ID_Concepto + ID_Programa_Periodo`.
- [ ] No hubo filas en blanco insertadas.
- [ ] No hubo hijos huérfanos.

---

## Resultado esperado
Con estas reglas, el flujo queda:

`Contrato -> CAF -> Programa -> Periodos -> Ejecución -> APU`

con **upsert real**, **vínculos correctos** y **sin duplicados**.
