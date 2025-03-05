import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px

abre_arquivo_csv_com_pandas = pd.read_csv('IAeML\\archive\\credit_risk_dataset.csv')

print(abre_arquivo_csv_com_pandas)                                                                  # Mostra os primeiros e ultimos dados da tabela

#print(abre_arquivo_csv_com_pandas[abre_arquivo_csv_com_pandas['person_income'] >= 58000])          # Retorna as pessoas com person_income

print(abre_arquivo_csv_com_pandas[abre_arquivo_csv_com_pandas['loan_percent_income'] <= 0.00])      # Retorna as pessoas com loan_percent_income

print(np.unique(abre_arquivo_csv_com_pandas['cb_person_default_on_file']))