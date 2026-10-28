# 📋 Reglas de Negocio — Liquidación de Manifiestos Marítimos (v2.0)

> Especificación funcional de las reglas operativas, comisionables, de facturación en Caja y de auditoría implementadas por `manifest_engine` v2.0 para la agencia ficticia **Atlantic Maritime Agency (AMA)**.
>
> ⚠️ **Nota importante sobre tarifas y alícuotas:** todos los porcentajes, mínimos, aranceles por unidad y montos fijos documentados abajo son **valores sintéticos e ilustrativos de portafolio** definidos en `manifest_engine/config.py`. La estructura lógica reproduce la complejidad real de una agencia marítima multipropósito (*Containers, Ro-Ro & Break-Bulk*).

---

## 1. Gobernanza por Catálogo (`Definiciones_reglas_comisiones.xlsx`)

El motor no adivina ni infiere reglas sobre textos desconocidos. Toda decisión parte de tres tablas maestras mantenidas en Excel:

1. **`CONCEPTOS DE CARGO`:** define para cada concepto del bloque `Charge Information`:
   - **Rol frente a `ALL IN`:** `FLETE BASE`, `RECARGO INCLUIDO EN EL FLETE (rompe ALL IN)`, `GASTO FUERA DEL FLETE (no afecta ALL IN)` o `DESCUENTO DEL FLETE`.
   - **Efecto en la base comisionable:** `SUMA`, `RESTA` o `NO AFECTA`.
   - **Conceptos nuevos:** si aparece un cargo inédito en un manifiesto, no altera la comisión de esa corrida y se registra automáticamente al final de la hoja como gasto neutro con aviso para auditoría.
2. **`TIPOS DE BULTO`:** asigna cada tipo de bulto del manifiesto a una de las **5 familias operativas** (`CONTENEDOR`, `BBULK`, `CAR`, `VAN`, `RORO`) e indica si computa para el **mínimo por contenedor** y para **Control Equipment Fees (`CEF`)**.
   - **Regla de equipos de apoyo (`Mafi` / `Bolster`):** los *roll-trailers* (`Mafi 40`, `60 ft. Mafi...`) son equipos de línea utilizados para estibar carga `BBULK`. No tributan mínimo de contenedor ni `CEF`, y cuando el B/L detalla los bultos cargados encima (ej. `2-CASE(S)`), el equipo no duplica unidades comerciales.
3. **`PUERTOS EXPO`:** determina explícitamente en qué puertos de descarga de exportación rige la alícuota regional de **Brasil** (`SANTOS`, `RIO DE JANEIRO`, `PARANAGUA`, `ITAJAI`) y en cuáles la alícuota general de ultramar. Un puerto ausente en la tabla envía el B/L a `REVISIÓN`.

---

## 2. Alcance Documental y Monedas

- **Conocimientos válidos (`S3...`):** sólo se liquidan *Bills of Lading* comerciales cuyo número inicia con `S3` (8 dígitos, incluyendo transbordos marcados `[T]`). Los documentos `P3...` corresponden a posicionamiento interno de equipo vacío y se descartan al parsear.
- **Conversión multimoneda (`ROE`):** las comisiones se liquidan en **USD**. Si la base comisionable está expresada en **EUR** (o combina líneas en `USD` y `EUR`), la porción en euros se convierte mediante el tipo de cambio del viaje (`ROE`, por defecto ilustrativo `1 EUR = 1,08 USD`).
- **Monedas no convertibles:** si existe base comisionable positiva en otra divisa (ej. `GBP`), el motor **no aplica tipos de cambio implícitos**: deja la comisión en blanco y emite alerta `MONEDA` / `COMISIÓN - REVISAR`.

---

## 3. Reglas de Comisión a Casa Matriz (Valores Ilustrativos)

### 3.1 Base comisionable y condición `ALL IN` (80 %)
- **Base bruta y neta:** suman los cargos marcados `SUMA` (`Basic Frt.`, `Open Top S/C`, `Over Hght S/C`, `Over Wght S/C`) y restan los descuentos negativos marcados `RESTA` (`FAC`, `Rebate`, `Commercial Adj.`). El neto nunca puede ser menor que cero.
- **Evaluación `ALL IN` (por B/L completo):** un B/L califica como `ALL IN` únicamente cuando su flete consiste de forma pura en `Basic Frt.` (pudiendo incluir gastos locales ajenos al flete como `Toll S/C`).
  - Cualquier recargo marítimo (`BAF`, `ETS+FUEL EU Surcharge`, `BAF Decrease`, `Sweeping`, `ENS`, `ISPS`, etc.) o adicional de sobredimensión (`Open Top S/C`, `Over Hght S/C`) **rompe la condición `ALL IN`** para todo el B/L.
  - **Asimetría operativa de `THC`:** en **IMPO**, `THC (POD)` es el gasto local en Argentina (no rompe `ALL IN`), mientras que `THC (POL)` es un cargo de origen externo (sí rompe `ALL IN`). En **EXPO** ocurre exactamente a la inversa (`THC (POL)` es local y no rompe `ALL IN`; `THC (POD)` rompe `ALL IN`).
- **Efecto `ALL IN`:** cuando el B/L es `ALL IN`, la alícuota porcentual se aplica sobre el **80 %** del `Basic Frt.` neto de descuentos.

### 3.2 Esquema tarifario ilustrativo por familia de carga

| Familia / Concepto | Importación (`IMPO`) | Exportación General (`EXPO`) | Exportación Brasil (`EXPO`) | Regla especial / Mínimo |
|---|---|---|---|---|
| **Contenedores (`CONTENEDOR`)** | `2,0 %` | `4,0 %` | `3,5 %` | **Mínimo `25,00 USD` por contenedor:** se toma `max(porcentaje_ajustado, N_ctr × 25 USD)`. |
| **Carga Suelta (`BBULK`)** | `2,0 %` | `4,0 %` | `3,5 %` | Sin mínimo por unidad (incluye carga sobre `Mafi`). |
| **Rodados cerrados por AMA (`MODE_AGENCY`)** | `1,0 %` | `1,0 %` | `1,0 %` | Aplica sobre la base comisionable (con ajuste `ALL IN` si corresponde) para los B/L informados por Comercial. |
| **Rodados no cerrados por AMA (`MODE_UNIT`)** | Por unidad | Por unidad | Por unidad | **`2,50 USD` / `CAR`** · **`3,50 USD` / `VAN`** (*Small* y *Big Van*) · **`12,00 USD` / `RORO`** (*New/Used LM RoRo*), incluso en transbordos sin `Basic Frt.`. |
| **Control Equipment Fees (`CEF`)** | `4,00 USD / ctr` | `4,00 USD / ctr` | `4,00 USD / ctr` | Suma al `GRAND TOTAL` por cada contenedor elegible (excluye `Mafi`). |

### 3.3 Lista de Comercial para Rodados
Al procesar un viaje existen tres estados posibles respecto de la nómina de bookings cerrados por la agencia **AMA**:
1. **Lista con B/L (`S300000103, ...`):** esos conocimientos liquidan al `1,0 %` (`VEHICLES BOOKED BY AMA`) y el resto de los rodados liquida por unidad (`MODE_UNIT`). Si se informa un B/L inexistente en el viaje o sin rodados, se alerta en `LISTA COMERCIAL`.
2. **`NINGUNO`:** confirmación explícita de Comercial de que no hubo cierres locales; todos los rodados liquidan en firme por unidad (`MODE_UNIT`).
3. **Vacío (`ESTIMACIÓN`):** aún no se recibió la lista comercial; el motor calcula preventivamente por unidad (`MODE_ESTIMADO`) y encabeza la planilla y la hoja `REVISIÓN` advirtiendo que los importes de rodados son estimados.

---

## 4. Cargas Mixtas y Separación de Filas (`split.py`)

Cuando un mismo B/L transporta más de una familia con distinta regla (por ejemplo `1 New LM RoRo` + `4 CASE(S)`), el motor intenta adjudicar cada línea de flete y de `TOLL` de forma determinística:
1. **Por coincidencia exacta de medidas:** factor facturado igual al peso (`MT`), volumen (`CBM`), metros lineales (`LM`), tamaño de contenedor (`20`/`40`) o unidades (`PU`) de una única familia.
2. **Por descarte frente a contenedor:** en un B/L de `CONTENEDOR + otra carga`, las líneas en base `20`/`40` pertenecen al contenedor y las restantes a la otra familia.
3. **Por `Tariff ID` o cobertura completa:** si una cotización apunta unívocamente a las medidas de una sola carga, o una familia ya cubrió todos sus bultos, el remanente se adjudica a la otra.

- **Si la adjudicación es 100 % unívoca:** la Biblia abre una sub-fila por familia (sin duplicar cobros globales `Collect` ni `BAF`) y la planilla de comisiones liquida cada porción en su sección respectiva.
- **Si persiste ambigüedad (ej. flete global `1 AA` entre `RoRo` y `BBULK` cuando ambos requieren flete):** el B/L **no se inventa ni prorratea**; permanece en una fila `MIXTO`, su comisión queda en blanco y se documenta en `REVISIÓN`.

---

## 5. Gastos de Línea, ENS y Facturación para Caja

### 5.1 Tarifario ilustrativo de control de gastos de línea (`line_costs.py`)
El motor respeta en las Biblias el monto impreso en el manifiesto y audita contra los valores esperados:
- **`THC` Contenedor `20'`:** `100,00 USD` por unidad (debe ser múltiplo exacto).
- **`THC` Contenedor `40'` / `45'`:** `120,00 USD` por unidad (debe ser múltiplo exacto).
- **`TOLL` (`S/C`) Contenedor:** `150,00 USD` por contenedor.
- **`TOLL` (`S/C`) Carga Suelta / Rodados:** tarifa por tonelada (`MT`) autodetectada de la moda del viaje (default ilustrativo `20,00 USD/Tn`).
- **`Sweeping`:** `10,00 USD` por contenedor (`20'` y `40'` por igual).
- **`ENS` (sólo `EXPO`):** cada B/L de exportación que incluye el cargo `ENS` liquida un monto fijo ilustrativo de **`10,00 USD`** en la Biblia EXPO y genera una línea en `<VIAJE> - ENS.xlsx`.

### 5.2 Reglas del libro `Facturación.xlsx` (Caja)
- **Identificación tributaria (`CUIT`):** se extrae del bloque `CN:` en `IMPO` y `SH:` en `EXPO`, validando su dígito verificador **Módulo 11**. Si no figura o el dígito es inválido, la celda queda vacía para carga manual.
- **Alcance local (`Buenos Aires` y `Buenos Aires - Partner`):** todo `Total[Buenos Aires...]` `Collect` en `IMPO` o `Prepaid` en `EXPO` constituye cobranza local de la agencia.
- **Desglose de conceptos locales en `EXPO`:** si el B/L sólo tiene `Total[Buenos Aires]` en esa moneda, todos los cargos `Prepaid` de esa moneda integran la factura local. Si coexiste con `Total[Head Office]` / exterior, el motor busca la **combinación exacta y única** de cargos que sume el total de Buenos Aires; si hubiera más de una combinación posible con igual suma, informa el total general sin arriesgar un desglose equívoco.

---

## 6. Las 8 Categorías Accionables de la Hoja `REVISIÓN`

Cada libro incluye una hoja `REVISIÓN` libre de ruido operativo (se omiten deliberadamente situaciones normales como *Service B/L* de transbordo sin flete o ausencias parciales de VIN). Sólo se reportan desvíos accionables:

| Categoría | Origen | Significado operativo |
|---|---|---|
| `COMISIÓN - REVISAR` | `commission.py` | B/L cuya comisión quedó **en blanco** (puerto no catalogado, concepto indefinido, moneda no convertible o flete mixto ambiguo). |
| `CARGA MIXTA` | `split.py` | B/L multifamilia que no pudo abrirse en sub-filas en la Biblia por ambigüedad en flete o `TOLL`. |
| `LISTA COMERCIAL` | `__main__.py` | Corrida en modo `ESTIMACIÓN` (sin lista de Comercial) o B/L informado por Comercial que no existe / no trae rodados. |
| `GASTO LINEA` | `line_costs.py` | Diferencia entre `THC`, `TOLL` o `Sweeping` del manifiesto y el tarifario fijo esperado. |
| `MONEDA` | `validation.py` | Base comisionable positiva en divisa distinta de `USD` y `EUR`. |
| `CARGA` | `validation.py` | El B/L posee `Basic Frt.` pero ningún bulto pudo clasificarse en el catálogo. |
| `TOTAL MANIFIESTO` | `billing.py` | La suma de líneas de cargo `(P/C, moneda)` no cuadra con los renglones `Total[...]` del propio manifiesto. |
| `IMPORTES` | `billing.py` | Anomalía de cobranza local (ej. `Prepaid` local en una `IMPO`, `Collect` neto negativo o moneda local fuera de columnas). |
