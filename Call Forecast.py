import pandas as pd
from prophet import Prophet
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load CSV
df = pd.read_csv(
    r"C:\Users\bj.aguirre\Downloads\GoodCharlie Dash Data(New Inbound).csv",
    encoding="latin-1"
)

# 2. Create AHT column
df['aht'] = df['Hold Time'] + df['Talk Time'] + df['Wrap up time']

# 3. Convert date and group weekly
df['Call start'] = pd.to_datetime(df['Call start'])
weekly = df.groupby(pd.Grouper(key='Call start', freq='W')).agg(
    calls=('Call start', 'count'),
    aht=('aht', 'mean')
).reset_index().rename(columns={'Call start': 'ds'})

# 4. Forecast Calls (1 year = 52 weeks)
calls_df = weekly[['ds', 'calls']].rename(columns={'ds':'ds','calls':'y'})
calls_model = Prophet()
calls_model.fit(calls_df)
future_calls = calls_model.make_future_dataframe(periods=52, freq='W')
forecast_calls = calls_model.predict(future_calls)

# 5. Forecast AHT (1 year = 52 weeks)
aht_df = weekly[['ds', 'aht']].rename(columns={'ds':'ds','aht':'y'})
aht_model = Prophet()
aht_model.fit(aht_df)
future_aht = aht_model.make_future_dataframe(periods=52, freq='W')
forecast_aht = aht_model.predict(future_aht)

# 6. Visualization style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (14, 6)
plt.rcParams['axes.titlesize'] = 18
plt.rcParams['axes.labelsize'] = 14

# Last historical dates
last_hist_date_calls = calls_df['ds'].max()
last_hist_date_aht = aht_df['ds'].max()

# Filter forecast only future
forecast_future_calls = forecast_calls[forecast_calls['ds'] > last_hist_date_calls]
forecast_future_aht = forecast_aht[forecast_aht['ds'] > last_hist_date_aht]

# --- Calls Forecast ---
plt.figure(figsize=(14,6))
plt.plot(calls_df['ds'], calls_df['y'], color="black", linestyle="-", linewidth=2, label="Historical Data")
plt.plot(forecast_future_calls['ds'], forecast_future_calls['yhat'], color="blue", linestyle="--", linewidth=2, label="Forecast")
plt.fill_between(forecast_future_calls['ds'], forecast_future_calls['yhat_lower'], forecast_future_calls['yhat_upper'], 
                 color="lightblue", alpha=0.4, label="Confidence Interval")
plt.axvline(x=last_hist_date_calls, color="red", linestyle=":", linewidth=2, label="Forecast Start")
plt.title("Weekly Call Volume Forecast - Next Year", fontsize=20, fontweight='bold')
plt.xlabel("Date")
plt.ylabel("Number of Calls")
plt.legend()
plt.show()

# --- AHT Forecast ---
plt.figure(figsize=(14,6))
plt.plot(aht_df['ds'], aht_df['y'], color="black", linestyle="-", linewidth=2, label="Historical Data")
plt.plot(forecast_future_aht['ds'], forecast_future_aht['yhat'], color="green", linestyle="--", linewidth=2, label="Forecast")
plt.fill_between(forecast_future_aht['ds'], forecast_future_aht['yhat_lower'], forecast_future_aht['yhat_upper'], 
                 color="lightgreen", alpha=0.4, label="Confidence Interval")
plt.axvline(x=last_hist_date_aht, color="red", linestyle=":", linewidth=2, label="Forecast Start")
plt.title("Weekly AHT Forecast - Next Year", fontsize=20, fontweight='bold')
plt.xlabel("Date")
plt.ylabel("Average Handling Time (seconds)")
plt.legend()
plt.show()
