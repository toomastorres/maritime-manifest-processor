# Motor de Procesamiento de Manifiestos — Documentación

> Reescritura modular de `manifest_processor_v3.py` (monolito de 1609 líneas) al
> paquete `manifest_engine/`. Estado: **motor completo y ejecutable** — los 5
> reportes (planilla IMPO/EXPO, comisiones IMPO/EXPO, ENS) se generan vía CLI.

## Cómo correr (primera instancia)

```
python -m manifest_engine [carpeta_entrada] [carpeta_salida]
```

Busca los CSV (`ATL0426_IMPO_1.CSV`, …), los agrupa por buque-viaje y genera por
cada uno los 5 reportes en la carpeta de salida (`<COD>_<VOY>-Impo.xlsx`,
`-MontocomZte.xlsx`, `-Expo.xlsx`, `-MontoExpoZte.xlsx`, `-ENS.xlsx`). Al iniciar
pide el **ROE** (EUR→USD) y el **TOLL suelto/Tn** (20 actual / 18 viajes
viejos). Cada workbook trae una hoja **REVISIÓN** con los controles.

Precisión vs verdad exacta (medida por B/L, evaluando fórmulas): IMPO ~95%, ENS
~96%, comisiones 100% clasificación / 91% montos, **EXPO datos ~97.6%** (la lógica
prepaid/Nápoles/abroad implementa la regla del negocio; las diffs restantes son
retoques manuales de la planilla + carga de conexión/transbordo y rolada, que son
operativas → van a REVISIÓN). El núcleo de extracción financiera ronda 95–100%.

> Nota sobre EXPO: la biblia NO es 1:1 con el manifiesto — el operador agrega carga
> de conexión de otros buques (referida como "Ex S3-..." en otros manifiestos) y
> quita carga rolada. El motor procesa fielmente el manifiesto del buque; esas
> diferencias operativas no son errores del motor.

---

## 1. Objetivo

Extraer, depurar y procesar manifiestos marítimos IMPO/EXPO (CSV) de una agencia
marítima y generar las "Biblias" (reportes Excel) calculando gastos de línea,
comisiones y la hoja TOLL con precisión, más una hoja **REVISIÓN** que señala lo
que requiere control humano.

**Verdad exacta:** los `.xlsx` en las carpetas `<COD> V. <MMYY>/` (IMPO, Expo,
MontocomZte, MontoExpoZte, ENS) están hechos a mano por el usuario y son la
referencia de comparación. Pueden contener ajustes/typos manuales → el motor
produce la versión canónica correcta y marca las divergencias en REVISIÓN.

---

## 2. Arquitectura (paquete `manifest_engine/`)

```
manifest_engine/
├── config.py          Parámetros de negocio "modificables" (tarifas, rates, mapeos)
├── domain/
│   ├── models.py      BLRecord, Charge, CargoData, ChargeLine, TotalLine (dataclasses)
│   └── cargo.py       Clasificación de carga + extracción de marca/agencia
├── parsing/
│   └── manifest.py    ManifestParser: CSV -> List[BLRecord] (+ consolidación de duplicados)
├── rules/
│   ├── line_costs.py  THC/TOLL/SWEEPING: validación vs tarifa fija
│   ├── commission.py  Monto Comisión, split por basis, max(%,30·ctr) por BL
│   └── biblia.py      Prepaid/Collect y formato de celdas (IMPO/EXPO)
├── reports/
│   ├── styles.py      Constantes de estilo Excel + abreviar_puerto
│   ├── impo.py        Planilla IMPO (hoja por puerto + TOLL) ✅
│   └── revision.py    Hoja REVISIÓN
├── validation.py      Capa REVISIÓN: consolida alertas (BL/tipo/detalle)
└── verification.py    Harness de diff celda-por-celda (diff_sheet/auto_offset)
```

Flujo: `parse_manifest(csv,'IMPO')` → `List[BLRecord]` → `generar_planilla_impo(bls, ..., cfg)` que internamente usa `rules/*` y agrega la hoja REVISIÓN vía `validation.revisar_bls`.

---

## 3. Reglas de negocio implementadas

### 3.1 Monto Comisión (base comisionable)
- **Comisionable** = `Basic Frt.` + `Open Top S/C` + `Over Hght S/C` (y variantes
  Over Height/Weight) − descuentos (`adj`, `rebate`, `fac`, `baf decrease`).
- **Nunca negativo**: si el neto ≤ 0 → 0 (se muestra `0,00`).
- **Split por tipo de carga** según el `BASIS` de cada línea de cargo:
  `20`/`40` → CONTENEDORES; cualquier otro (`PU/MT/ME/WMM/PS/AA/CM`) → GENERAL/AUTOS.
  Un BL mixto se reparte en ambas secciones (verificado contra la verdad exacta).
- **Rates de comisión** (config): Rolling 1%, General Cargo 2% (IMPO) / 4% (EXPO),
  Contenedores 2%/4% **Y** 25 USD/ctr → el final es el **mayor de ambos, por BL**.
- **EUR→USD** vía fórmula Excel con el ROE (no se convierte en el motor). GBP/BRL/SEK
  son informativos (sin conversión).

### 3.2 Gastos de línea (tarifa FIJA, modo validación)
- THC 100 (20') / 120 (40') por contenedor; S/C (TOLL) 150/contenedor y 20/Tn
  en carga suelta; SWEEPING 10/contenedor.
- Histórico: el S/C suelto era **18/Tn** (viajes ~2025, ej. ATL 0426). Para
  reproducir esos: `LineCostTariffs(sc_loose_per_ton=18)`.
- Cualquier importe del manifiesto que **no** coincida con la tarifa = **error de
  carga de origen** → REVISIÓN (no se recalcula sobre el dato malo).
- La columna PESO de la planilla IMPO = `=TOLL / tarifa_fija`.

### 3.3 Clasificación de carga y marca
- Carga: contenedores (Tank/Dry Cargo/High Cube/Open Top), vehículos (New/Used
  Car/Van/RoRo), general (Pallet/Unit/Package/Case/Bundle/Crate/Bag/Piece/...).
- **Marca real**: se extrae de la **descripción de la mercadería** (no del consignee)
  vía `config.MAKE_KEYWORDS` (DELTAR, MB, CORVEX, FENIX, BIONDA, AVALON, LYRA,
  MISTRAL, NOVAX, ORBIS, GRIFON, PEGASO, QUASAR, ...). Sufijo de agencia (OCEANLINK/DELCAR/
  PREMIUMCARS/EXIMPORT) desde el consignee.
- **El motor suele ser más preciso que la verdad** (ej. descripción "MARCA DELTAR /
  VIN ZZDLTR..." pero la planilla a mano escribió "MISTRAL"). Decisión del usuario:
  el motor escribe la marca precisa/canónica y marca la divergencia en REVISIÓN.

### 3.4 Biblia IMPO (columnas de la planilla)
Layout (hoja por puerto de descarga, que en IMPO siempre es ZÁRATE):
`A=B/L, C=CARGA, D=PESO, E=DTN, F=TOLL, G/H=THC 20'/40', I=SWP, J/K=DOLARES
PREPAID/COLLECT, L/M=EURO PREPAID/COLLECT, N=BAF, O=COMM`.
- Grupos por **puerto de carga** ordenados **alfabéticamente**; cada grupo cierra
  con fila `TOTALES` (`=SUM`).
- Celdas con **fórmulas-componentes** (`=2767.6-69.19`, `=Total-BAF`, `=F/factor`)
  como la verdad exacta.
- PREPAID/COLLECT según la condición (P/C) del **`Basic Frt.`** del BL.

---

## 4. Detalles del parser (IMPO)
- Lectura por líneas (latin-1), pipe-delimitado. Detecta BL `S3########`.
- `P3########` = delimitador lógico: cierra el BL anterior, no se imprime.
- Captura `charge_lines` estructuradas (DESCRIPTION/FACTOR/BASIS/RATE/TOTAL/cond/CURR),
  incluyendo descripciones **envueltas** en 2 líneas (ej. `Over Hght` / `S/C ...`).
- **Consolida BLs duplicados** (BL partido en varias páginas) en un único registro.
- Captura `description_lines` (col 4) para extraer la marca.
- POD en IMPO siempre ZÁRATE; lo dinámico es el **Port of Loading**.

---

## 5. Capa REVISIÓN
Hoja `REVISIÓN` embebida en cada workbook (lista simple, columnas BL/TIPO/DETALLE).
Tipos de alerta:
| TIPO | Detecta |
|---|---|
| GASTO LINEA | TOLL/THC/SWEEPING fuera de tarifa fija (error de origen) |
| COMISION | BL sin Basic Frt., o comisionable clampeado a 0 |
| MONEDA | Basic Frt. en GBP/BRL/SEK (no convertible) |
| CARGA | Hay Basic Frt. pero no se clasificó la carga |
| CANTIDAD | Unidades del Basic Frt. ≠ carga declarada (basis discreto 20/40/PU) |
| ENS | ENS del manifiesto ≠ 15 (forzado) |

---

## 6. Estado y precisión (IMPO)

- **Paridad 100%** del parser y las reglas vs el script original (22 CSV / 1362 BLs).
- **Escritor IMPO vs verdad exacta ATL 0426: ~95%** celda-por-celda.
- **Panorama por B/L (12 viajes, 997 BLs, métrica inmune al posicionamiento):**
  extracción ~94% global.
  - Financiero **excelente**: THC40 100%, SWP 99%, prepaid/collect 97–99%,
    comisionable (COMM) 95.4%, TOLL 92%.
  - **CARGA ~72%**: gran parte es el motor corrigiendo grafías manuales
    inconsistentes (DELTAR↔MISTRAL, MB↔EVORA) → la métrica subestima.
- El ~4% restante en ATL 0426 son artefactos hechos a mano (ajustes/typos) → REVISIÓN.

---

## 7. Cómo correr y verificar

```python
from manifest_engine import config
from manifest_engine.parsing import parse_manifest
from manifest_engine.reports.impo import generar_planilla_impo
from manifest_engine.verification import diff_workbooks, auto_offset, diff_sheet

# Viaje viejo (tarifa suelta 18); viajes nuevos usan 20
cfg = config.EngineConfig(tariffs=config.LineCostTariffs(sc_loose_per_ton=18))
bls = parse_manifest('ATL 0426_IMPO_1.CSV', 'IMPO')
generar_planilla_impo(bls, 'ATLANTIC STAR', '0426', 'salida.xlsx', cfg=cfg)

# Comparar contra la verdad exacta
diff_workbooks('salida.xlsx', 'GBA V. 0625/IMPO ATL 0426.xlsx', tol=0.5)
```
Para un panorama por-BL (extracción pura, inmune a filas), comparar cada BL por su
número usando `bl_rows` + evaluación de fórmulas aritméticas (ver historial de la
sesión; se puede formalizar como comando en la etapa CLI).

---

## 8. Decisiones del usuario (confirmadas)
1. Verdad exacta = los xlsx hechos a mano (referencia de diff).
2. Gastos de línea: leer y **validar** vs tarifa fija (mismatch = error de origen).
3. `max(%, 30·ctr)` de contenedores: **por BL**.
4. Parser: reforzar solo si hace falta.
5. Ajustes manuales (TOLL en origen, comisionables a 0): el motor calcula limpio y
   **marca** en REVISIÓN; no los inventa.
6. Celdas con **fórmulas-componentes**.
7. ROE (EUR→USD) en Excel; GBP/BRL/SEK informativos sin conversión.
8. `Basic Frt.` Collect entra al comisionable igual que Prepaid.
9. Marca real desde la **descripción**; el motor escribe la canónica y marca divergencias.
10. CARGA: mostrar siempre la general en combos RORO+general.
11. REVISIÓN: hoja embebida, lista simple, monedas raras informativas, auditoría
    de cantidades solo sobre conteo discreto.

---

## 9. Pendiente / preguntas abiertas
- Patrón exacto de **TOLL pagado en origen** para auto-calcular el `+887.5` (hoy se flaggea).
- Si un `Basic Frt.` viene en GBP/BRL/SEK: ¿convertir con TC o dejar en REVISIÓN?
- Conteos dispares (ej. 30 vs 6 CORVEX): ¿retoque manual o el motor cuenta siempre lo de la descripción?
- Era de tarifa TOLL por viaje (18 vs 20): hoy se pasa por config; idealmente auto-detectar.

---

## 10. Próximo: EXPORTACIÓN (EXPO)
Replicar el patrón ya probado en IMPO:
1. **Escritor EXPO** (planilla ZÁRATE + TOLL): layout de 19 columnas (DESTINO, B/L,
   CARGA, THC 20/40, SWP, ENS, DOLARES INCL PREPAID/COLLECT/ABROAD, MONTO COMISIÓN,
   %, cálculos de comisión, CTRS, AUTO, TOLL). Diffear contra `Expo <COD> <VOY>.xlsx`.
2. **MontoExpoZte** (comisiones EXPO) y **MontocomZte** (comisiones IMPO) — el
   comisionable ya está validado al 95–98%.
3. **ENS** (debit note, monto fijo 15 USD por BL).
4. Reglas EXPO específicas: THC/SWEEPING tipo P normal y tipo C → mensaje; "USD INCL
   PREPAID" = Total BA USD − THC(P) − SWEEPING(P) (si negativo → MATRIZ); COLLECT =
   "Total [País] [Moneda]"; ABROAD = "Total Matriz USD".
5. La capa REVISIÓN ya es reutilizable (pasar `op_type='EXPO'`).

> El parser ya soporta EXPO (dict por BL, ENS, totales Matriz/abroad). Lo que falta
> es el **escritor EXPO** y los de comisiones/ENS, más el ajuste fino contra la verdad.
