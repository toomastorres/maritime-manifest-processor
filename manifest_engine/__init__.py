"""
manifest_engine - Motor de procesamiento de manifiestos marítimos (Atlantic Maritime Agency — AMA).

- config:      parámetros de negocio modificables (tarifas ilustrativas, tasas, mapeos)
- catalog:     lectura de Definiciones_reglas_comisiones.xlsx (conceptos, bultos, puertos)
- domain:      modelo de datos (BLRecord, ChargeLine, CargoItem...) y clasificación
- parsing:     lectura del CSV de manifiesto -> List[BLRecord]
- rules:       reglas de negocio (comisión, reparto por carga, importes, gastos de línea)
- reports:     planillas Excel (biblias, comisiones, facturación, ENS, extras)
- validation:  alertas accionables para la hoja REVISIÓN
"""

__version__ = "2.0.0"
