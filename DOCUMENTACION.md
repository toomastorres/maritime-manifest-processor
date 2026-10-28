# 🛠️ Guía Técnica — Procesador de Manifiestos Marítimos (v2.0)

> Documentación de arquitectura, modelo de datos, algoritmos de conciliación y puntos de extensión del paquete `manifest_engine` v2.0.
>
> *Nota: todos los nombres de compañía (`Atlantic Maritime Agency — AMA`), buques (`ATLANTIC STAR`), códigos, identificadores tributarios y valores tarifarios son 100 % sintéticos con fines demostrativos de portafolio.*

---

## 1. Visión arquitectónica y principios de diseño

El motor sigue una arquitectura en capas **estrictamente desacoplada** donde el parseo, la gobernanza de catálogo, el cálculo de reglas y la renderización en Excel no comparten estado mutable:

1. **Gobernanza externa por catálogo (`catalog.py`):** ninguna clasificación de cargos marítimos, familias de bulto o puertos regionales vive cableada en condicionales `if/elif`. El motor carga `Definiciones_reglas_comisiones.xlsx` en estructuras inmutables (`@dataclass(frozen=True)`).
2. **Principio *"Never Invent Values"*:** cuando los datos de entrada no permiten resolver un cálculo con certeza matemática única (moneda no convertible, puerto desconocido, flete global `1 AA` en carga mixta o combinación no única de cargos locales), el motor devuelve `None` y emite una alerta tipada en la hoja `REVISIÓN`.
3. ** Validación híbrida de gastos de línea (`line_costs.py`):** el motor **lee** los importes reales facturados en el manifiesto (sin sobrescribirlos) y los contrasta contra el tarifario esperado (`LineCostTariffs`), reportando desvíos de carga de origen sin alterar la contabilidad del documento.
4. **Fórmulas vivas en Excel (`reports/`):** las planillas generadas (`IMPO`, `EXPO`, `Comm. Inward`, `Comm. Outward`, `Facturación`) escriben fórmulas dinámicas `=SUM(...)`, `=MAX(...)` y referencias entre celdas para que el analista pueda auditar o completar casos en revisión viendo actualizarse los totales en tiempo real.

---

## 2. Evolución del Motor (v1.0 → v2.0): De script ETL a motor gobernado por catálogo

| Dimensión | Versión 1.0 | Versión 2.0 |
|---|---|---|
| **Clasificación de cargos y bultos** | Diccionarios y prefijos fijos en Python. | Gobernada por usuario en `Definiciones_reglas_comisiones.xlsx` con listas desplegables y auto-descubrimiento (`register_new_concepts`). |
| **Comisiones sobre flete** | Porcentaje directo por categoría primaria del B/L. | Condición `ALL IN` (80 %), excepción asimétrica `THC (POL)` en IMPO, tasa diferencial a Brasil, mínimo por contenedor y Control Equipment Fees (`CEF`). |
| **Esquema de rodados** | Porcentaje fijo único sobre flete. | Tres modos explícitos: `AMA` (1 % del flete para cierres de la agencia), `UNIDAD` (tarifa fija por `CAR` / `VAN` / `RORO`) y `ESTIMADO` (con aviso en `REVISIÓN`). |
| **B/L con cargas mixtas** | Una fila única clasificada por prioridad fija. | Partición determinística (`split.py`) en sub-filas independientes asignando flete y `TOLL` por medidas exactas, `Tariff ID` o descarte; fila `MIXTO` protegida si hay ambigüedad. |
| **Liquidación para Caja** | No disponible; lectura simple de totales. | Nuevo libro `Facturación.xlsx` con extracción de `CUIT` validada por **Módulo 11**, soporte `Buenos Aires - Partner`, cuadratura de totales y búsqueda combinatoria de líneas `Prepaid`. |
| **Auditoría (`REVISIÓN`)** | Alertas genéricas con falsos positivos operativos. | **8 categorías accionables** depuradas de ruido operativo (transbordos sin flete o VIN parciales omitidos por diseño). |

---

## 3. Detalle por módulo

### 3.1 `manifest_engine/catalog.py` — Catálogo gobernado y auto-descubrimiento
Lee tres hojas de `Definiciones_reglas_comisiones.xlsx` (o la ruta definida en la variable de entorno `MANIFEST_DEFINICIONES`):
- **`CONCEPTOS DE CARGO`:** mapea cada concepto a `ConceptRule(name, category, effect)`, donde `category in {FLETE, RECARGO, GASTO, DESCUENTO}` y `effect in {SUMA, RESTA, NO AFECTA}`.
- **`TIPOS DE BULTO`:** mapea cada descripción exacta (`New Car(s)`, `40 ft. High Cube`, `Mafi 40`, etc.) a `KindRule(name, family, min_ctr, cef)`. Detecta automáticamente equipos de apoyo (`is_equipment`, como `Mafi` o `Bolster`) para que no inflen conteos de bultos ni disparen mínimos de contenedor cuando transportan carga `BBULK`.
- **`PUERTOS EXPO`:** mapea cada puerto de descarga a `PortRule(name, country, brasil)`.
- **`register_new_concepts(found, origen, path)`:** si durante una corrida aparecen cargos ausentes en el catálogo, abre el libro Excel copiando estilos y validaciones de datos de la última fila, inserta los conceptos nuevos como `GASTO FUERA DEL FLETE` / `NO AFECTA`, actualiza el rango de `AutoFilter` y la fórmula `=COUNTBLANK(...)`, e invalida la caché en memoria (`_cache`).

### 3.2 `manifest_engine/parsing/manifest.py` — Parser unificado y CUIT Módulo 11
Procesa archivos CSV que encapsulan registros `pipe-delimited` de ancho fijo (`2,  |col1|...|col7|,0;;`):
- **Filtrado documental:** procesa únicamente conocimientos comerciales `S3\d{8}` (incluyendo transbordos `[T]`) y descarta bloques `P3\d{8}` de reposicionamiento de equipo vacío.
- **Reconstrucción de cargos de dos líneas (`_read_charge_column` + `concept_name`):** el sistema de origen suele partir nombres largos en dos renglones (`Toll` arriba y `S/C(POD) 3.000 40 ...` abajo, o `Over Hght` arriba y `S/C ...` abajo). El parser conserva el encabezado pendiente por B/L y reconstruye el nombre canónico exacto sin confundir `Over Hght S/C` con `Toll S/C`.
- **Asociación de medidas por ítem (`_read_measures`):** vincula toneladas (`tn`), volumen (`cbm`) y metros lineales (`lm`) al `CargoItem` activo para alimentar la partición de cargas mixtas.
- **Extracción y validación Módulo 11 de CUIT (`extract_party_id`, `cuit_ok`):** aísla el bloque de texto correspondiente al rol tributario local (`CN` en `IMPO`, `SH` en `EXPO`), recompone números partidos entre renglones y verifica el dígito verificador con pesos `(5, 4, 3, 2, 7, 6, 5, 4, 3, 2)` en módulo 11 (rechazando restos `dv == 10` o dígitos inválidos).

### 3.3 `manifest_engine/rules/commission.py` — Evaluación y asignación de flete
Evalúa cada `BLRecord` produciendo un `CommissionResult` con trazabilidad completa:
- **Determinación de `ALL IN`:** un B/L es `ALL IN` (se comisiona sobre el **80 %** del `Basic Frt.`) sólo si no contiene ningún cargo clasificado como `RECARGO`, ningún adicional que sume al flete distinto de `Basic Frt.` (ej. `Open Top S/C`, `Over Hght S/C`), ni la excepción asimétrica de `THC (POL)` en `IMPO`.
- **Algoritmo de asignación en B/L mixtos (`_assign`):**
  1. **Coincidencia métrica (`_metric_candidates`):** compara el `factor` y `basis` de cada línea comisionable (`20`/`40` → `CONTENEDOR`; `MT`/`WMM`/`ME`/`PU` → ítem cuya medida coincide dentro de `TOL = 0.011`).
  2. **Descarte con contenedor:** en combinaciones binarias `CONTENEDOR + otra familia`, lo facturado por `20'`/`40'` se adjudica al contenedor y el resto a la carga acompañante.
  3. **Evidencia por cotización (`_tariff_evidence`):** si el manifiesto informa `Tariff ID` y alguna línea auxiliar de esa cotización (BAF, TOLL, THC) coincide unívocamente con las medidas de una sola familia, toda la cotización se adjudica a esa familia.
  4. **Cobertura completa (`_covered`):** si todos los bultos de una familia ya tienen flete propio asignado, las líneas remanentes se asignan a la única familia pendiente.
  5. **Propagación de descuentos porcentuales (`x% of y`):** los ajustes como `FAC` o `Commercial Adj. (x% of y)` heredan la familia de la línea cuyo importe coincide con su base `pct_of`.

### 3.4 `manifest_engine/rules/split.py` — Sub-filas para Biblias
Transforma cada `(BLRecord, CommissionResult)` en una lista de objetos `Fila`:
- Si el B/L tiene una sola familia o es imposible separar sin ambigüedad (`not res.split_ok` o `_toll_groups(...) is None`), devuelve una única fila (`TIPO` o `MIXTO`) y `split_alerts()` genera la advertencia para `REVISIÓN`.
- Si la separación es exacta, emite una `Fila` por familia (orden `CONTENEDOR → RODADO → BBULK`): los importes globales del B/L (`Collect` local, `BAF`) se escriben **una sola vez** en `first=True`, `THC` y `Sweeping` van en la fila con `container_costs=True`, y cada fila recibe exclusivamente sus unidades, su `TOLL` y su porción comisionable.

### 3.5 `manifest_engine/rules/billing.py` — Conciliación para Caja
Resuelve qué cobra efectivamente la agencia en Buenos Aires:
- **`is_local(place)`:** reconoce tanto `Total[Buenos Aires]` como `Total[Buenos Aires - Partner]` (bookings de subagente local cobrados en ventanilla local).
- **`local_lines(bl, letra, moneda)`:** cuando en `EXPO` existen múltiples totales `Prepaid` en la misma moneda (ej. `Total[Buenos Aires] 585.00P USD` y `Total[Head Office] 2,550.00P USD`), explora mediante `itertools.combinations` (hasta `MAX_LINEAS_COMBINAR = 18`) todos los subconjuntos de cargos `P USD` que suman `585.00 ± 0.005`. Si existe **exactamente una firma única** de conceptos e importes, devuelve esas líneas para detallar la factura de Caja; si existen cero o dos o más combinaciones empatadas, devuelve `None` evitando adjudicar conceptos por azar.
- **`total_mismatches(bl)`:** verifica para cada par `(P/C, moneda)` que `sum(charge_lines) == sum(total_lines)` dentro de `TOL_TOTAL = 0.10`.

---

## 4. Ejecución, CLI y pruebas

### Interfaz de línea de comandos (`manifest_engine/__main__.py`)
```bash
# Procesamiento interactivo o por lotes de una carpeta
python -m manifest_engine <carpeta_entrada> <carpeta_salida>

# Generación de modelos vacíos de comisiones con fórmulas pre-armadas
python -m manifest_engine --modelos <carpeta_modelos>
```

En uso programático (como lo invoca `streamlit_app.py` o `pytest`), `process_directory()` acepta parámetros explícitos sin bloquear por `stdin`:
```python
from manifest_engine.__main__ import process_directory

process_directory(
    input_dir="samples",
    output_dir="out",
    roe=1.08,
    agency_override=frozenset({"S300000103"}),
    register_concepts=False,  # evita mutar el catálogo versionado durante demos/tests
)
```

### Pruebas automatizadas (`tests/test_engine.py`)
Ejecutar `pytest -v` corre las **22 pruebas** que cubren desde funciones puras (`cuit_ok`, `local_lines`, `evaluate`, `filas_del_bl`, `register_new_concepts`) hasta la generación completa de los 9 libros Excel y verificación estructural de sus hojas con `openpyxl`.
