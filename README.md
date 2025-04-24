# Análisis Técnico de Mercados

Este programa realiza un análisis técnico completo de varios ETFs y acciones del mercado, calculando diferentes indicadores técnicos y generando señales de trading.

## Indicadores Calculados

### 1. ROC (Rate of Change)
- **Cálculo**: `((Precio_Actual - Precio_Anterior) / Precio_Anterior) * 100`
- **Período**: 26 días
- **Interpretación**: 
  - Positivo: Tendencia alcista
  - Negativo: Tendencia bajista

### 2. MACD Trimestral (La data aprovechamos la mensual)
- **Parámetros**: (36, 78, 21)
- **Cálculo**:
  - EMA 36: Media móvil exponencial de 36 períodos
  - EMA 78: Media móvil exponencial de 78 períodos
  - MACD = EMA 36 - EMA 78
  - Señal = EMA 21 del MACD
- **Coloración**:
  - Verde: EMA 36 > Señal
  - Rosa: EMA 36 < Señal

### 3. MACD Mensual
- **Parámetros**: (12, 26, 9)
- **Cálculo**:
  - EMA 12: Media móvil exponencial de 12 períodos
  - EMA 26: Media móvil exponencial de 26 períodos
  - MACD = EMA 12 - EMA 26
  - Señal = EMA 9 del MACD
- **Coloración**:
  - Verde: MACD > (Estocástico SK>%D o SK> 85)
  - Amarillo: Si no es Verde ni Rosa
  - Rosa: MACD < Señal y (Estocástico SK<%D o SK> 85)

### 4. MACD Semanal
- **Parámetros**: (12, 26, 9)
- **Cálculo**: Igual que MACD Mensual pero con datos semanales
- **Coloración**:
  - Verde: MACD > Señal
  - Rosa: MACD < Señal

### 5. Señal de Cruce
- **Parámetros**: (12, 9)
- **Cálculo**:
  - EMA 12: Media móvil exponencial de 12 períodos
  - EMA 9: Media móvil exponencial de 9 períodos
- **Señales**:
  - Azul: EMA12 (penúltima DATA) < Señal Y EMA12 (última DATA) > Señal
  - Naranja: EMA12 (penúltima DATA) > Señal Y EMA12 (última DATA) < Señal

### 6. Estocástico
- **Parámetros**: (89, 3)
- **Cálculo**:
  - %K = 100 * ((Cierre - Mínimo) / (Máximo - Mínimo))
  - %D = Media móvil de 3 períodos del %K
- **Uso**: Confirmación de señales mensuales

## Salida del Programa

El programa genera un archivo Excel (`resultados.xlsx`) con las siguientes columnas:

1. **Ticker**: Símbolo del activo
2. **ROC**: Rate of Change
3. **Trimestral**: Señal MACD trimestral (verde/rosa)
4. **Mensual**: Señal MACD mensual (verde/amarillo/rosa)
5. **Semanal**: Señal MACD semanal (verde/rosa)
6. **Señal**: Cruce de EMAs (azul/naranja)

## Requisitos

- Python 3.x
- Bibliotecas:
  - yfinance
  - pandas
  - numpy
  - openpyxl
  - ta (Technical Analysis)

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

El programa descargará los datos históricos, calculará los indicadores y generará el archivo Excel con los resultados.

