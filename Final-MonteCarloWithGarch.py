import pandas as pd
import numpy as np
from arch import arch_model
from openpyxl import load_workbook

# Load data from Excel
file_path = 'C:/Users/mikes/Downloads/NVDA.xlsx'
sheet_name = 'NVDA'
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Extract columns
dates = pd.to_datetime(df['Date'])
close_prices = df['Close']
open_prices = df['Open']
high_prices = df['High']
low_prices = df['Low']
adj_close_prices = df['Adj Close']
volume = df['Volume']

# Calculate daily returns for each column
close_returns = close_prices.pct_change().dropna()
open_returns = open_prices.pct_change().dropna()
high_returns = high_prices.pct_change().dropna()
low_returns = low_prices.pct_change().dropna()
adj_close_returns = adj_close_prices.pct_change().dropna()
volume_returns = volume.pct_change().dropna()

# Define the number of simulations
num_simulations = 20
forecast_days = 30

# Apply rescaling based on the warnings
scaling_factors = {
    'Close': forecast_days * 10,
    'Open': forecast_days * 10,
    'High': forecast_days * 10,
    'Low': forecast_days * 10,
    'Adj Close': forecast_days * 10,
    'Volume': forecast_days * 10
}

# Rescale returns
scaled_returns = {
    'Close': close_returns * scaling_factors['Close'],
    'Open': open_returns * scaling_factors['Open'],
    'High': high_returns * scaling_factors['High'],
    'Low': low_returns * scaling_factors['Low'],
    'Adj Close': adj_close_returns * scaling_factors['Adj Close'],
    'Volume': volume_returns * scaling_factors['Volume']
}

# Fit GARCH(1,1) models to the rescaled returns for each column
garch_fits = {}
garch_forecasts = {}

for key, returns in scaled_returns.items():
    garch_model = arch_model(returns, vol='Garch', p=1, q=1)
    garch_fit = garch_model.fit(disp="off")
    garch_forecast = garch_fit.forecast(horizon=forecast_days)
    garch_fits[key] = garch_fit
    garch_forecasts[key] = garch_forecast.variance.values[-1, :] ** 0.5


# List to store the results
all_simulations = {key: [] for key in garch_fits.keys()}

# Monte Carlo Simulation for each column
for sim_num in range(num_simulations):
    last_prices = {
        'Close': close_prices.iloc[-1],
        'Open': open_prices.iloc[-1],
        'High': high_prices.iloc[-1],
        'Low': low_prices.iloc[-1],
        'Adj Close': adj_close_prices.iloc[-1],
        'Volume': volume.iloc[-1]
    }
    simulated_prices = {key: [last_prices[key]] for key in last_prices.keys()}
    
    for day in range(forecast_days):
        for key in garch_fits.keys():
            daily_volatility = garch_forecasts[key][day] / scaling_factors[key]  # Adjust back by dividing by scaling factor
            daily_return = np.random.normal(scaled_returns[key].mean() / scaling_factors[key], daily_volatility)
            next_price = simulated_prices[key][-1] * (1 + daily_return)
            simulated_prices[key].append(next_price)
        
        # Enforce the constraint: Open, Close, and Adj Close must be between High and Low
        simulated_high = simulated_prices['High'][-1]
        simulated_low = simulated_prices['Low'][-1]
        
        for key in ['Open', 'Close', 'Adj Close']:
            simulated_prices[key][-1] = np.clip(simulated_prices[key][-1], simulated_low, simulated_high)
    
    # Create DataFrames for each simulation and store them
    simulation_dates = [dates.max() + pd.DateOffset(days=i) for i in range(1, forecast_days + 1)]
    for key in garch_fits.keys():
        sim_df = pd.DataFrame({
            'Simulation Number': sim_num + 1,
            'Simulated Date': simulation_dates,
            f'Simulation Result {key}': simulated_prices[key][1:]  # Exclude the initial price
        })
        all_simulations[key].append(sim_df)

# Combine all simulations into a single DataFrame for each column
final_simulation_dfs = {key: pd.concat(all_simulations[key], ignore_index=True) for key in all_simulations.keys()}

# Merge all DataFrames together on 'Simulation Number' and 'Simulated Date'
final_result_df = final_simulation_dfs['Close']
for key in ['Open', 'High', 'Low', 'Adj Close', 'Volume']:
    final_result_df = pd.merge(final_result_df, final_simulation_dfs[key], on=['Simulation Number', 'Simulated Date'])

# Write results to Excel
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    final_result_df.to_excel(writer, sheet_name=sheet_name, startcol=12, index=False)
    col_names = ['Simulation Number', 'Simulated Date', 'Simulation Result Close', 'Simulation Result Open', 
                 'Simulation Result High', 'Simulation Result Low', 'Simulation Result Adj Close', 'Simulation Result Volume']
    for idx, col_name in enumerate(col_names, start=13):
        writer.sheets[sheet_name].cell(row=1, column=idx, value=col_name)

print("Monte Carlo simulation results with constraints and proper scaling incorporated written to Excel.")

