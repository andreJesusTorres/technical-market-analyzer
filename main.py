import yfinance as yf
import pandas as pd
import numpy as np
from ta.trend import MACD
from ta.momentum import ROCIndicator, StochasticOscillator
import openpyxl
import os
import subprocess
from datetime import datetime, timedelta
import sys

# Configuración
TICKERS = [
    'SPY', 'TLT', 'QQQ', 'SLV', 'GLD', 'USO', 'XLE', 'XLRE', 'XLI', 
    'XLF', 'XLB', 'XLY', 'XLK', 'XLP', 'XLV', 'XLU', 'UUP', 'MTUM', 
    'MOAT', 'SPYV', 'SPYG', 'RSP', 'IWO', 'IWN', 'GMF', 'IBIT', 'FXI'
]

# Constantes
ROC_WINDOW = 26
MACD_FAST = 12
MACD_SLOW = 26
MACD_SIGNAL = 9
STOCH_WINDOW = 89
STOCH_SMOOTH = 3 
EXCEL_FILE = 'resultados.xlsx'

# Fechas para descarga de datos
END_DATE = datetime.now()
START_DATE = END_DATE - timedelta(days=3650)  # 10 años

def print_header():
    """Imprime un encabezado atractivo para la aplicación"""
    print("\n" + "=" * 60)
    print("ANÁLISIS TÉCNICO DE MERCADOS".center(60))
    print("=" * 60)
    print(f"Fecha: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print(f"Tickers analizados: {len(TICKERS)}")
    print(f"Período: {START_DATE.strftime('%d-%m-%Y')} - {END_DATE.strftime('%d-%m-%Y')}")
    print("=" * 60 + "\n")

def download_data(ticker, period):
    """Descarga datos históricos para un ticker específico"""
    try:
        data = yf.download(
            ticker, 
            start=START_DATE, 
            end=END_DATE, 
            interval=period, 
            auto_adjust=True,
            progress=False
        )
        return data
    except Exception as e:
        print(f"Error descargando datos para {ticker}: {str(e)}")
        return pd.DataFrame()

def calculate_indicators(df):
    """Calcula los indicadores técnicos para el análisis"""
    # ROC (Rate of Change)
    df['ROC'] = ((df['Close'] - df['Close'].shift(ROC_WINDOW)) / df['Close'].shift(ROC_WINDOW)) * 100
    
    # MACD (Moving Average Convergence Divergence)
    exp1 = df['Close'].ewm(span=MACD_FAST, adjust=False).mean()
    exp2 = df['Close'].ewm(span=MACD_SLOW, adjust=False).mean()
    df['MACD'] = exp1 - exp2
    df['MACD_Signal'] = df['MACD'].ewm(span=MACD_SIGNAL, adjust=False).mean()
    df['MACD_Hist'] = df['MACD'] - df['MACD_Signal']
    
    # Stochastic Oscillator
    low_min = df['Low'].rolling(window=STOCH_WINDOW).min()
    high_max = df['High'].rolling(window=STOCH_WINDOW).max()
    df['Stochastic'] = 100 * ((df['Close'] - low_min) / (high_max - low_min))
    df['Stochastic'] = df['Stochastic'].rolling(window=STOCH_SMOOTH).mean()
    
    return df

def get_macd_signal(df):
    """Determina si MACD está cortado al alza o a la baja"""
    return 'alza' if df['MACD_Hist'].iloc[-1] > 0 else 'baja'

def get_stoch_signal(df):
    """Determina si el Estocástico está al alza o por encima de 85"""
    current_stoch = df['Stochastic'].iloc[-1]
    prev_stoch = df['Stochastic'].iloc[-2]
    return current_stoch > 85 or current_stoch > prev_stoch

def get_resource_path(relative_path):
    """Obtiene la ruta absoluta para recursos empaquetados"""
    try:
        # PyInstaller crea un directorio temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def export_to_excel(df):
    """Exporta los resultados a Excel con formato de color"""
    print("Generando archivo Excel con resultados...")
    
    # Obtener la ruta del directorio del ejecutable
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable
        application_path = os.path.dirname(sys.executable)
    else:
        # Si es script Python
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    # Definir la ruta completa del archivo Excel
    excel_path = os.path.join(application_path, EXCEL_FILE)
    
    # Verificar si el archivo ya existe
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        ws = wb.create_sheet('Sheet')
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
    
    # Estilo para encabezados
    header_style = openpyxl.styles.PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    font_style = openpyxl.styles.Font(color="FFFFFF", bold=True)
    
    # Escribir los encabezados
    headers = ['Ticker', 'ROC', 'Mensual', 'Semanal']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_style
        cell.font = font_style
    
    # Definir los colores para los indicadores
    verde = openpyxl.styles.PatternFill(
        start_color="90EE90", end_color="90EE90", fill_type="solid"
    )  # Verde claro
    amarillo = openpyxl.styles.PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )  # Amarillo
    rosa = openpyxl.styles.PatternFill(
        start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"
    )  # Rosa claro
    
    # Escribir los datos y aplicar colores
    for row, (_, data) in enumerate(df.iterrows(), 2):
        # Ticker
        ws.cell(row=row, column=1, value=data['Ticker'])
        # ROC
        roc_cell = ws.cell(row=row, column=2, value=data['ROC'])
        
        # Aplicar formato condicional al ROC
        if data['ROC'] > 0:
            roc_cell.font = openpyxl.styles.Font(color="006100")  # Verde oscuro
        else:
            roc_cell.font = openpyxl.styles.Font(color="9C0006")  # Rojo oscuro
        
        # Mensual (aplicar color)
        cell_mensual = ws.cell(row=row, column=3)
        if data['Mensual'] == 'verde':
            cell_mensual.fill = verde
        elif data['Mensual'] == 'amarillo':
            cell_mensual.fill = amarillo
        else:
            cell_mensual.fill = rosa
            
        # Semanal (aplicar color)
        cell_semanal = ws.cell(row=row, column=4)
        if data['Semanal'] == 'verde':
            cell_semanal.fill = verde
        else:
            cell_semanal.fill = rosa
    
    # Ajustar el ancho de las columnas
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = adjusted_width
    
    # Agregar bordes a todas las celdas
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin'),
        right=openpyxl.styles.Side(style='thin'),
        top=openpyxl.styles.Side(style='thin'),
        bottom=openpyxl.styles.Side(style='thin')
    )
    
    for row in ws.iter_rows(min_row=1, max_row=len(df) + 1, min_col=1, max_col=4):
        for cell in row:
            cell.border = thin_border
    
    # Guardar el archivo en la ruta correcta
    wb.save(excel_path)
    
    # Abrir el archivo Excel automáticamente
    try:
        print("Abriendo archivo Excel...")
        if os.name == 'nt':  # Windows
            os.startfile(excel_path)
        elif os.name == 'posix':  # macOS o Linux
            if os.uname().sysname == 'Darwin':  # macOS
                subprocess.call(['open', excel_path])
            else:  # Linux
                subprocess.call(['xdg-open', excel_path])
        print("✓ Archivo Excel abierto correctamente")
    except Exception as e:
        print(f"No se pudo abrir el archivo Excel automáticamente: {str(e)}")
        print(f"El archivo se ha guardado en: {os.path.abspath(excel_path)}")

def main():
    """Función principal del programa"""
    # Mostrar encabezado
    print_header()
    
    # Inicializar lista de resultados
    results = []
    
    # Procesar cada ticker
    print("Iniciando análisis de tickers...")
    
    total_tickers = len(TICKERS)
    for i, ticker in enumerate(TICKERS, 1):
        try:
            # Mostrar progreso
            print(f"Procesando {ticker} ({i}/{total_tickers})...")
            
            # Descargar datos semanales y mensuales
            weekly_data = download_data(ticker, '1wk')
            monthly_data = download_data(ticker, '1mo')
            
            if weekly_data.empty or monthly_data.empty:
                print(f"No se pudieron obtener datos para {ticker}")
                continue
            
            # Calcular indicadores
            weekly_data = calculate_indicators(weekly_data)
            monthly_data = calculate_indicators(monthly_data)
            
            # Obtener último ROC semanal
            roc_value = weekly_data['ROC'].iloc[-1]
            
            # Determinar colores según condiciones
            macd_monthly = get_macd_signal(monthly_data)
            stoch_condition = get_stoch_signal(monthly_data)
            macd_weekly = get_macd_signal(weekly_data)
            
            # Determinar color mensual
            if macd_monthly == 'alza' and stoch_condition:
                monthly_color = 'verde'
            elif macd_monthly == 'alza' or stoch_condition:
                monthly_color = 'amarillo'
            else:
                monthly_color = 'rosa'
                
            # Determinar color semanal
            weekly_color = 'verde' if macd_weekly == 'alza' else 'rosa'
            
            # Agregar resultados
            results.append({
                'Ticker': ticker,
                'ROC': round(roc_value, 2) if not np.isnan(roc_value) else 0,
                'Mensual': monthly_color,
                'Semanal': weekly_color
            })
            
        except Exception as e:
            print(f"Error procesando {ticker}: {str(e)}")
            continue
    
    # Verificar si hay resultados
    if not results:
        print("No se pudieron procesar datos para ningún ticker")
        return
    
    # Crear DataFrame y ordenar por ROC
    df_results = pd.DataFrame(results)
    df_results = df_results.sort_values('ROC', ascending=False)
    
    # Mostrar resumen de resultados
    print("\nResumen de resultados:")
    print(f"✓ Tickers procesados: {len(df_results)}/{len(TICKERS)}")
    print(f"✓ Mayor ROC: {df_results['ROC'].max():.2f}% ({df_results.iloc[0]['Ticker']})")
    print(f"✓ Menor ROC: {df_results['ROC'].min():.2f}% ({df_results.iloc[-1]['Ticker']})")
    
    # Exportar a Excel con formato
    try:
        export_to_excel(df_results)
        print(f"\n✓ Resultados guardados en '{EXCEL_FILE}'")
    except Exception as e:
        print(f"Error guardando el archivo Excel: {str(e)}")
    
    print("\n" + "¡ANÁLISIS COMPLETADO!".center(60))
    print("=" * 60 + "\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario")
    except Exception as e:
        print(f"\nError inesperado: {str(e)}")
    finally:
        print("Fin del programa") 