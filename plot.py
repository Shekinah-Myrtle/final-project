import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
from openpyxl import Workbook

#PLOTTING

file_path = 'FinalDataSet.xlsx'
sheet_name = 'Material DQI'
df=pd.read_excel(file_path, sheet_name=sheet_name)



# Clean the data: Remove rows where 'Embodied Carbon' or 'DQI Score' are not numeric
df_clean = df[pd.to_numeric(df['Embodied Carbon - kgCO2e/kg'], errors='coerce').notna()]
df_clean = df_clean[pd.to_numeric(df['DQI Score'], errors='coerce').notna()]

# Convert the columns to numeric
df_clean['Embodied Carbon - kgCO2e/kg'] = pd.to_numeric(df_clean['Embodied Carbon - kgCO2e/kg'])
df_clean['DQI Score'] = pd.to_numeric(df_clean['DQI Score'])

# Set Seaborn style
sns.set(style="whitegrid")

# Create a comparison bar plot
plt.figure(figsize=(14, 8))

# Reshape the data for plotting
df_melted = df_clean.melt(id_vars=['Material'], value_vars=['Embodied Carbon - kgCO2e/kg', 'DQI Score'],
                          var_name='Metric', value_name='Value')

# Create the bar plot
sns.barplot(x='Material', y='Value', hue='Metric', data=df_melted, palette="muted")

# Customize plot
plt.title('Comparison of Embodied Carbon and DQI Score for Different Materials', fontsize=16)
plt.xlabel('Material', fontsize=14)
plt.ylabel('Value', fontsize=14)
plt.xticks(rotation=45, ha="right", fontsize=12)
plt.yticks(fontsize=12)
plt.legend(title='Metrics', fontsize=12)

# Display the plot
plt.tight_layout()
plt.show()