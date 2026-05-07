import pandas as pd
import numpy as np
from typing import List
from app.models.sensor import Sensor


def _serie_referencia(sensores: List[Sensor], campo: str) -> pd.Series:
    """Returns the column-wise mean of real (non-simulated) sensors."""
    flag = f'simulado_{"temp" if campo == "temperatura" else "hum"}'
    series = [
        s.datos[campo]
        for s in sensores
        if s.datos is not None
        and campo in s.datos.columns
        and s.datos[campo].notna().any()
        and not getattr(s, flag)
    ]
    if series:
        return pd.concat(series, axis=1).mean(axis=1).reset_index(drop=True)
    return pd.Series(dtype=float)


def _simular_serie(
    referencia: pd.Series,
    min_val: float,
    max_val: float,
    setpoint: float,
    n: int,
    sensor_seed: int = 0,
    es_humedad: bool = False,
) -> np.ndarray:
    rng = np.random.default_rng(seed=sensor_seed)

    # Build base signal
    if referencia.notna().sum() > 5:
        base = referencia.interpolate().ffill().bfill().values[:n]
        if len(base) < n:
            base = np.pad(base, (0, n - len(base)), mode='edge')
    else:
        # No reference: simulate realistic drift around setpoint
        # Slow sinusoidal drift (compressor cycle ~20 samples = 100 min).
        # Minimum absolute amplitude so round(v,1) / round(v,2) never produces
        # a flat constant line regardless of range width.
        t = np.linspace(0, 4 * np.pi, n)
        amp = max((max_val - min_val) * 0.04, 0.6 if es_humedad else 0.18)
        drift = amp * np.sin(t + rng.uniform(0, 2 * np.pi))
        # Secondary slow trend
        drift2 = amp * 0.5 * np.sin(t * 0.3 + rng.uniform(0, 2 * np.pi))
        base = np.full(n, setpoint) + drift + drift2

    # Per-sensor systematic offset (position effect)
    span = max_val - min_val
    if es_humedad:
        bias = rng.uniform(-max(span * 0.03, 0.5), max(span * 0.03, 0.5))
        # Minimum absolute noise so round(v,1) never produces a flat constant line.
        # span*0.015 fails for tight ranges (e.g. span=4 → noise_std=0.06 → all 75.0).
        noise_std = max(span * 0.015, 0.35)
    else:
        bias = rng.uniform(-max(span * 0.06, 0.15), max(span * 0.06, 0.15))
        noise_std = max(span * 0.025, 0.10)

    # High-frequency noise
    noise = rng.normal(0, noise_std, n)
    # Smooth slightly (rolling 3) — keeps realistic short-term variation
    noise = pd.Series(noise).rolling(3, min_periods=1, center=True).mean().values

    # Slow oscillation (compressor cycle) unique per sensor
    osc_period = rng.integers(15, 30)  # 15-30 samples period
    # Minimum absolute osc amplitude so visible even for narrow ranges
    osc_amp = max(noise_std * 1.5, 0.3 if es_humedad else 0.08)
    osc = osc_amp * np.sin(np.linspace(0, n / osc_period * 2 * np.pi, n)
                           + rng.uniform(0, 2 * np.pi))

    result = np.clip(base + bias + noise + osc, min_val, max_val)
    return result


def simular_sensores_faltantes(
    sensores: List[Sensor],
    timestamps: pd.Series,
    rango_temp_min: float,
    rango_temp_max: float,
    rango_hum_min: float,
    rango_hum_max: float,
    num_sensores: int = 9,
    setpoint_temp: float = None,
    setpoint_hum: float = None,
) -> List[Sensor]:
    n = len(timestamps)

    # Default setpoints to midpoint if not provided
    if setpoint_temp is None:
        setpoint_temp = (rango_temp_min + rango_temp_max) / 2
    if setpoint_hum is None:
        setpoint_hum = (rango_hum_min + rango_hum_max) / 2

    # Create placeholder sensors for positions without real data
    posiciones_existentes = {s.posicion for s in sensores}
    for pos in range(1, num_sensores + 1):
        if pos not in posiciones_existentes:
            sensor_nuevo = Sensor(
                serial=f'SIM-{pos:02d}',
                nombre=f'Sensor Simulado {pos:02d}',
                descripcion='Simulado',
                posicion=pos,
                tiene_temperatura=True,
                tiene_humedad=True,
                simulado_temp=True,
                simulado_hum=True,
            )
            sensor_nuevo.datos = pd.DataFrame({
                'timestamp': timestamps.values,
                'temperatura': np.nan,
                'humedad': np.nan,
            })
            sensores.append(sensor_nuevo)

    # ── Detectar y marcar sensores con datos sospechosamente planos ─────────────
    # Un sensor real siempre presenta variación sobre 24h. Si std < umbral, el dato
    # es inválido (sensor bloqueado, export erróneo, columna equivocada) y se fuerza
    # simulación. Se hace ANTES de calcular la referencia para no contaminarla.
    _FLAT_HUM  = 0.15   # %RH — variación mínima realista en 24h
    _FLAT_TEMP = 0.05   # °C
    for _s in sensores:
        if _s.datos is None or _s.simulado_hum:
            continue
        if 'humedad' in _s.datos.columns:
            _hv = _s.datos['humedad'].dropna()
            if len(_hv) > 10 and _hv.std() < _FLAT_HUM:
                _s.datos['humedad'] = pd.Series(np.nan, index=_s.datos.index)
        if 'temperatura' in _s.datos.columns and not _s.simulado_temp:
            _tv = _s.datos['temperatura'].dropna()
            if len(_tv) > 10 and _tv.std() < _FLAT_TEMP:
                _s.datos['temperatura'] = pd.Series(np.nan, index=_s.datos.index)

    ref_temp = _serie_referencia(sensores, 'temperatura')
    ref_hum  = _serie_referencia(sensores, 'humedad')

    for sensor in sensores:
        if sensor.datos is None:
            sensor.datos = pd.DataFrame({
                'timestamp': timestamps.values,
                'temperatura': np.nan,
                'humedad': np.nan,
            })

        if sensor.tiene_temperatura and not sensor.tiene_datos_temp():
            sensor.datos['temperatura'] = _simular_serie(
                ref_temp, rango_temp_min, rango_temp_max, setpoint_temp, n,
                sensor_seed=sensor.posicion * 100 + 1,
                es_humedad=False,
            )
            sensor.simulado_temp = True

        if sensor.tiene_humedad and not sensor.tiene_datos_hum():
            sensor.datos['humedad'] = _simular_serie(
                ref_hum, rango_hum_min, rango_hum_max, setpoint_hum, n,
                sensor_seed=sensor.posicion * 100 + 2,
                es_humedad=True,
            )
            sensor.simulado_hum = True

    # Sort by position
    sensores.sort(key=lambda s: s.posicion)
    return sensores
