# 🎯 Technical Market Analyzer

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Pandas](https://img.shields.io/badge/Pandas-2.2.3-green.svg)](https://pandas.pydata.org/)
[![yfinance](https://img.shields.io/badge/yfinance-0.2.54-orange.svg)](https://github.com/ranaroussi/yfinance)
[![Technical Analysis](https://img.shields.io/badge/TA-0.11.0-purple.svg)](https://github.com/bukosabino/ta)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)](LICENSE)

> Advanced technical analysis tool for financial markets that calculates multiple indicators including ROC, MACD (multiple timeframes), and Stochastic oscillators to generate trading signals. **This project is part of my professional portfolio to demonstrate my development skills and practices.**

## 📋 Table of Contents

- [✨ Features](#-features)
- [🛠️ Technologies](#️-technologies)
- [📦 Installation](#-installation)
- [🎮 Usage](#-usage)
- [🏗️ Project Structure](#️-project-structure)
- [📊 Technical Indicators](#-technical-indicators)
- [🧪 Testing](#-testing)
- [📄 License](#-license)

## ✨ Features

### 🎯 Core Functionality
- **Multi-timeframe Analysis**: Weekly, monthly, and quarterly MACD calculations
- **Rate of Change (ROC)**: 26-period momentum indicator with color-coded output
- **MACD Signals**: Multiple timeframe MACD with customizable parameters
- **Stochastic Oscillator**: 89-period stochastic with %K and %D calculations
- **Crossover Detection**: Real-time EMA crossover signal detection
- **Excel Export**: Automated Excel report generation with color-coded formatting
- **Batch Processing**: Efficient processing of 27+ financial instruments
- **Error Handling**: Robust retry mechanism with exponential backoff

### 🎨 User Experience
- **Color-coded Output**: Terminal output with ANSI color formatting
- **Progress Tracking**: Real-time progress display with ticker count
- **Auto-open Results**: Automatic Excel file opening after completion
- **Comprehensive Statistics**: Summary statistics including mean, volatility, and positive percentage
- **Professional Header**: Attractive console header with timestamp and configuration

## 🛠️ Technologies

### Backend
| Technology | Version | Purpose |
|------------|---------|---------|
| [Python](https://www.python.org/) | 3.8+ | Core programming language |
| [Pandas](https://pandas.pydata.org/) | 2.2.3 | Data manipulation and analysis |
| [NumPy](https://numpy.org/) | 2.0.2 | Numerical computations |
| [yfinance](https://github.com/ranaroussi/yfinance) | 0.2.54 | Yahoo Finance data retrieval |
| [TA](https://github.com/bukosabino/ta) | 0.11.0 | Technical analysis indicators |
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.5 | Excel file generation |

### Data Sources
| Technology | Purpose |
|------------|---------|
| [Yahoo Finance API](https://finance.yahoo.com/) | Real-time and historical market data |
| [Alpha Vantage](https://www.alphavantage.co/) | Alternative financial data source |

### Development Tools
- **PyInstaller**: Executable generation for distribution
- **Rich**: Enhanced terminal output formatting
- **Requests**: HTTP library for API calls

## 📦 Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager
- Internet connection for data retrieval

### Quick Start

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/technical-market-analyzer.git
   cd technical-market-analyzer
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python main.py
   ```

4. **Access the results**
   - Excel file: `resultados.xlsx` (auto-opens after completion)
   - Console output: Real-time analysis with color-coded indicators

## 🎮 Usage

### Getting Started
1. **Execute the script**: Run `python main.py` in your terminal
2. **Monitor progress**: Watch real-time analysis of 27 financial instruments
3. **Review results**: Check the generated Excel file with color-coded indicators
4. **Analyze signals**: Use the technical indicators for trading decisions

### Key Features Usage

#### Batch Processing
```python
# The application processes tickers in batches for efficiency
BATCH_SIZE = 27  # Process all tickers at once
INTER_BATCH_DELAY = (0, 0)  # No pause between batches
```

#### Technical Indicator Calculation
```python
# Example of MACD calculation with custom parameters
def calculate_indicators(df):
    # ROC calculation
    last_price = df['Close'].iloc[-1]
    prev_price = df['Close'].iloc[-ROC_WINDOW-1]
    roc = ((last_price - prev_price) / prev_price) * 100
    
    # MACD calculation
    exp1 = df['Close'].ewm(span=MACD_FAST, adjust=False).mean()
    exp2 = df['Close'].ewm(span=MACD_SLOW, adjust=False).mean()
    df['MACD'] = exp1 - exp2
    return df
```

#### Excel Export with Formatting
```python
# Color-coded Excel output
verde = openpyxl.styles.PatternFill(
    start_color="90EE90", end_color="90EE90", fill_type="solid"
)  # Green for positive signals
rosa = openpyxl.styles.PatternFill(
    start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"
)  # Pink for negative signals
```

## 🏗️ Project Structure

```
technical-market-analyzer/
├── 📄 main.py                 # Main application logic
├── 📋 requirements.txt        # Python dependencies
├── 📖 README.md              # Project documentation
├── 🖼️ screener-ico.png       # Application icon
├── 📊 main.spec              # PyInstaller specification
└── 📈 resultados.xlsx        # Generated analysis results
```

## 📊 Technical Indicators

### Rate of Change (ROC)
- **Period**: 26 days
- **Formula**: `((Current_Price - Previous_Price) / Previous_Price) * 100`
- **Interpretation**: 
  - 🟢 Positive: Bullish trend
  - 🔴 Negative: Bearish trend

### MACD Trimestral
- **Parameters**: (36, 78, 21)
- **Calculation**: EMA 36 - EMA 78 with 21-period signal
- **Signals**:
  - 🟢 Green: EMA 36 > Signal
  - 🔴 Pink: EMA 36 < Signal

### MACD Mensual
- **Parameters**: (12, 26, 9)
- **Calculation**: EMA 12 - EMA 26 with 9-period signal
- **Signals**:
  - 🟢 Green: MACD > Signal
  - 🟡 Yellow: Neutral condition
  - 🔴 Pink: MACD < Signal

### MACD Semanal
- **Parameters**: (12, 26, 9)
- **Calculation**: Same as monthly but with weekly data
- **Signals**:
  - 🟢 Green: MACD > Signal
  - 🔴 Pink: MACD < Signal

### Crossover Signals
- **Parameters**: (12, 9)
- **Calculation**: EMA 12 vs EMA 9 crossover detection
- **Signals**:
  - 🔵 Blue: Bullish crossover
  - 🟠 Orange: Bearish crossover

### Stochastic Oscillator
- **Parameters**: (89, 3)
- **Calculation**: %K and %D with 3-period smoothing
- **Usage**: Confirmation for monthly signals

## 🧪 Testing

### Running Tests
```bash
# Currently manual testing through execution
python main.py

# Expected output:
# - Console progress with color-coded indicators
# - Excel file generation with proper formatting
# - Error handling for missing data
```

### Test Coverage
- ✅ Data download and validation
- ✅ Technical indicator calculations
- ✅ Excel export functionality
- ✅ Error handling and retry mechanisms
- ✅ Color-coded output formatting

## 📄 License

This project is proprietary software. All rights reserved. This code is made publicly available solely for portfolio demonstration purposes. See the [LICENSE](LICENSE) file for full terms and restrictions.

---

<div align="center">
  <p>
    <a href="#-technical-market-analyzer">Back to top</a>
  </p>
</div>

