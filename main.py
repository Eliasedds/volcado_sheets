import pandas as pd
import gspread
from key import get_credentials
from functions import open_dataset, letras, google_sheet, add_data


def run():
    #DDM-CantTk
    credentials = get_credentials()
    df = open_dataset('C:/Users/elias/OneDrive/Desktop/Prune/automatizacion_tareas/DDM_CantTk.xlsx')
    week = df['Semana'].unique()
    tupla = df.shape
    worksheet = google_sheet(credentials, 'test', week, tupla)
    letters = letras()
    add_data(letters, worksheet, df, tupla)

    #DDM-CantTk - Region
    df = open_dataset('C:/Users/elias/OneDrive/Desktop/Prune/automatizacion_tareas/DDM_CantTk - Region.xlsx')
    week = df['Semana'].unique()
    tupla = df.shape
    worksheet = google_sheet(credentials, 'test_1', week, tupla)
    letters = letras()
    add_data(letters, worksheet, df, tupla)

    #DDM-CantTk - con Importe
    df = open_dataset('C:/Users/elias/OneDrive/Desktop/Prune/automatizacion_tareas/DDM_CantTk - con Importe.xlsx')
    week = df['Semana'].unique()
    tupla = df.shape
    worksheet = google_sheet(credentials, 'test_2.2', week, tupla, to_gpp=False)
    letters = letras()
    add_data(letters, worksheet, df, tupla)


if __name__ == '__main__':
    run()