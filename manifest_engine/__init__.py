"""
manifest_engine - Motor de procesamiento de manifiestos marítimos.

Reescritura modular de manifest_processor_v3.py:
- config:      parámetros de negocio modificables (tarifas, rates, mapeos)
- domain:      modelo de datos tipado (BLRecord, Charge, CargoData) + clasificación
- parsing:     lectura del CSV de manifiesto -> List[BLRecord]
- rules:       reglas de negocio (gastos de línea, comisiones, biblias)
- reports:     escritores Excel (planilla IMPO/EXPO, comisiones, ENS)
- validation:  acumulación de discrepancias para revisión humana
"""

__version__ = "3.1.0"
