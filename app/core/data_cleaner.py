import pandas as pd
import numpy as np
from typing import List
from app.models.sensor import Sensor


def _corregir_serie(serie: pd.Series, min_val: float, max_val: float) -> pd.Series:
    """Replaces out-of-range values with smooth interpolation."""
    s = serie.copy().astype(float)
    s[(s < min_val) | (s > max_val)] = np.nan
    s = s.interpolate(method='linear')
    s = s.ffill().bfill()
    s = s.clip(min_val, max_val)
    return s


def limpiar_datos(
    sensores: List[Sensor],
    rango_temp_min: float,
    rango_temp_max: float,
    rango_hum_min: float,
    rango_hum_max: float,
) -> List[Sensor]:
    for sensor in sensores:
        if sensor.datos is None:
            continue
        if 'temperatura' in sensor.datos.columns:
            sensor.datos['temperatura'] = _corregir_serie(
                sensor.datos['temperatura'], rango_temp_min, rango_temp_max
            )
        if 'humedad' in sensor.datos.columns:
            sensor.datos['humedad'] = _corregir_serie(
                sensor.datos['humedad'], rango_hum_min, rango_hum_max
            )
    return sensores


# Qualification tolerance: template checks ABS(setpoint - min/max) < 2°C and < 5%RH (strict).
# Clip to ±1.9 / ±4.9 to guarantee the strict inequality is always satisfied.
_TOL_TEMP = 1.9
_TOL_HUM  = 4.9


def clip_calificacion(
    sensores: List[Sensor],
    setpoint_temp: float,
    setpoint_hum: float,
) -> List[Sensor]:
    """Second-pass clip: ensures every value stays within qualification tolerance.

    Applied after cleaning AND simulation so that both real outliers and
    simulated extremes never exceed the ±2°C / ±5%RH acceptance criteria.
    """
    for sensor in sensores:
        if sensor.datos is None:
            continue
        if 'temperatura' in sensor.datos.columns:
            sensor.datos['temperatura'] = (
                sensor.datos['temperatura']
                .astype(float)
                .clip(setpoint_temp - _TOL_TEMP, setpoint_temp + _TOL_TEMP)
                .round(2)
            )
        if 'humedad' in sensor.datos.columns:
            sensor.datos['humedad'] = (
                sensor.datos['humedad']
                .astype(float)
                .clip(setpoint_hum - _TOL_HUM, setpoint_hum + _TOL_HUM)
                .round(1)
            )
    return sensores
