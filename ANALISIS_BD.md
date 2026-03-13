# Análisis de la base de datos y relaciones (SGC)

## 1) Modelo de datos actual (fuente de verdad en código)

El modelo de datos está definido en `ESQUEMA_BD` dentro de `code.gs` y contiene **21 tablas** con formato de “tabla por hoja” en Google Sheets.

- Catálogos/sistema: `Usuarios_Sistema`, `Log_Actividad`, `Parametros_Sistema`, `Estadisticas_Financieras`, `Conversaciones_IA`.
- Núcleo contractual: `Convenios_Recurso`, `Contratistas`, `Contratos`, `Convenios_Modificatorios`, `Anticipos`.
- Planeación y ejecución: `Catalogo_Conceptos`, `Programa`, `Programa_Periodo`, `Programa_Ejecucion`, `Matriz_Insumos`.
- Estimaciones y pago: `Estimaciones`, `Detalle_Estimacion`, `Deducciones_Retenciones`, `Facturas`, `Pagos_Emitidos`, `Validacion_Archivos`.

## 2) Jerarquía relacional declarada

La jerarquía de relaciones está declarada en `RELACIONES_BD` con 3 niveles:

- **Nivel 1 (hijos de `Contratos`)**: `Catalogo_Conceptos`, `Programa`, `Estimaciones`, `Anticipos`, `Convenios_Modificatorios` (FK `ID_Contrato`).
- **Nivel 2**:
  - `Matriz_Insumos` (FK `ID_Concepto` -> `Catalogo_Conceptos`)
  - `Programa_Periodo` (FK `ID_Numero_Programa` -> `Programa`)
  - `Detalle_Estimacion`, `Deducciones_Retenciones`, `Facturas`, `Pagos_Emitidos`, `Validacion_Archivos` (FK `ID_Estimacion` -> `Estimaciones`)
- **Nivel 3**: `Programa_Ejecucion` con doble dependencia (`ID_Programa_Periodo` y `ID_Concepto`).

## 3) Cómo se implementa realmente la persistencia

En `db_utils.gs`, la capa CRUD (`dbSelect`, `dbInsert`, `dbUpdate`, `dbDelete`) opera directamente sobre hojas por nombre de tabla.

Aspectos clave:

- **No hay constraints nativos** (ni PK/FK reales de motor SQL): la integridad es lógica/aplicativa.
- `dbInsert` crea una hoja si no existe y usa `ESQUEMA_BD[tabla]` como cabecera.
- El ID se infiere desde la **primera columna** de la cabecera; puede autogenerar consecutivo o UUID corto según heurística.
- `dbUpdate` y `dbDelete` filtran por coincidencia de texto (`String(...)`) de condiciones.

## 4) Relaciones efectivamente usadas en la lógica

Sí hay validaciones y sincronizaciones por FK en procesos críticos, por ejemplo:

- `upsertConcepto` exige `ID_Contrato` para alta/edición de conceptos.
- `actualizarPeriodoPrograma` valida que el `ID_Concepto` exista en el catálogo del contrato (`getConceptosContrato(idContrato)`) antes de escribir `Programa_Ejecucion`.
- `eliminarConcepto` aplica borrado en cascada manual de `Programa_Ejecucion` por `ID_Concepto`.
- `getDetalleContrato` realiza join aplicativo `Contratos -> Contratistas` por `ID_Contratista`.

## 5) Hallazgos importantes

1. **`RELACIONES_BD` está declarada pero no se usa como motor genérico de integridad/cascada**.
   - En el código revisado, la cascada se implementa de forma puntual (ej. `eliminarConcepto`) y no mediante una función central basada en `RELACIONES_BD`.

2. **Riesgo de huérfanos por ausencia de validación FK universal en `dbInsert/dbUpdate`**.
   - Cualquier inserción directa a tablas hijas puede guardar IDs no existentes si la función de negocio no valida.

3. **Diferencias de contrato de esquema entre documentación y código, especialmente en `Contratos`**.
   - El `readme.md` y `ESQUEMA_BD` no son idénticos en nombres/campos de fianzas/retenciones (`Porcentaje_Amortizacion_Anticipo` vs `pct_Anticipo`, etc.).

4. **Generación de PK por heurística de primera columna**.
   - Funciona para la mayoría de tablas con `ID_*`, pero puede ser problemática en tablas con PK semántica (`Clave_Parametro`) porque puede autogenerar valor si llega vacío.

## 6) Recomendaciones técnicas priorizadas

1. **Centralizar integridad referencial en capa DB**:
   - Crear validación opcional/obligatoria en `dbInsert/dbUpdate` usando un mapa de FKs derivado de `RELACIONES_BD`.

2. **Estandarizar cascadas**:
   - Implementar un `dbDeleteCascade(tabla, condiciones)` que recorra `RELACIONES_BD` recursivamente y evite huérfanos.

3. **Unificar contrato de esquema**:
   - Tomar `ESQUEMA_BD` como fuente única y generar automáticamente el diccionario de datos para `readme.md`.

4. **Política de IDs explícita por tabla**:
   - Definir si cada PK es `numeric`, `uuid` o `natural key` para eliminar heurísticas ambiguas.

5. **Auditoría de consistencia**:
   - Script de diagnóstico periódico para detectar FKs sin padre (`ID_Contrato`, `ID_Estimacion`, `ID_Concepto`, `ID_Numero_Programa`, `ID_Programa_Periodo`).

## 7) Conclusión

El repositorio ya tiene un diseño relacional bastante claro y bien segmentado por dominios (contratos, programa, estimaciones, pagos). Sin embargo, la integridad hoy depende de la disciplina de las funciones de negocio; no existe enforcement global en la capa CRUD. La mayor oportunidad es convertir `RELACIONES_BD` en una pieza ejecutable (validación + cascada) para robustecer consistencia de datos y reducir riesgo de registros huérfanos.
