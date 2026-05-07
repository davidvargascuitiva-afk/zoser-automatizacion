from dataclasses import dataclass
from typing import Optional
import pandas as pd


@dataclass
class Sensor:
    serial: str
    nombre: str
    descripcion: str
    posicion: int           # 1-9, determines column in template
    tiene_temperatura: bool = True
    tiene_humedad: bool = False
    col_idx_temp: Optional[int] = None   # 1-based column in source Excel
    col_idx_hum: Optional[int] = None
    datos: Optional[pd.DataFrame] = None  # columns: timestamp, temperatura, humedad
    simulado_temp: bool = False
    simulado_hum: bool = False

    def tiene_datos_temp(self) -> bool:
        if self.datos is None:
            return False
        return 'temperatura' in self.datos.columns and self.datos['temperatura'].notna().any()

    def tiene_datos_hum(self) -> bool:
        if self.datos is None:
            return False
        return 'humedad' in self.datos.columns and self.datos['humedad'].notna().any()
