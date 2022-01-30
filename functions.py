def open_dataset(file_name:str):
    import pandas as pd
    df = pd.read_excel(file_name, engine='openpyxl')
    # df = pd.read_csv(file_name)
    return df

def letras() -> list:
    import string
    return list(string.ascii_lowercase)


def google_sheet(credentials:dict, sheet_name:str, week:int ,matrix_shape:tuple, to_gpp:bool=True):
    import gspread
    import time
    gc = gspread.service_account_from_dict(credentials)
    if to_gpp == True:
        sh = gc.open(sheet_name)
        sh.share('some_email@gmail.com', perm_type='user', role='writer')
        worksheet = sh.add_worksheet(title=f'Sem{week[0]}', rows=(matrix_shape[0] + 1), cols=matrix_shape[1])
    else:
        new_sh = gc.create(sheet_name)
        # time.sleep(30)
        # sh = gc.open(sheet_name)
        new_sh.share('some_email@gmail.com', perm_type='user', role='writer')
        new_sh.share('some_email@gmail.com', perm_type='user', role='writer')
        worksheet = new_sh.add_worksheet(title='Hoja 1', rows=(matrix_shape[0] + 1), cols=matrix_shape[1])
    return worksheet

def add_data(letter_list:list, worksheet, df, matrix_shape):
    letras_usadas = letter_list[0:matrix_shape[1]]
    sheet_columns = [letra.upper() for letra in letras_usadas]
    from_column, to_column = sheet_columns[0], sheet_columns[len(sheet_columns) - 1]
    columns = list(df.columns)

    for n in range(2):
        if n == 0:
            cell_list = worksheet.range(f'{from_column}1:{to_column}{matrix_shape[1]}')  

            for i, val in enumerate(columns):  #gives us a tuple of an index and value
                cell_list[i].value = val    #use the index on cell_list and the val from cell_values

            worksheet.update_cells(cell_list)

        elif n == 1:
            aux_dict = dict(zip(sheet_columns, columns))

            for key, value in aux_dict.items():
                data = list(df[value])
                cell_list = worksheet.range(f'{key}2:{key}{matrix_shape[0] + 1}')
                cell_values = data

                for i, val in enumerate(cell_values):  #gives us a tuple of an index and value
                        cell_list[i].value = val    #use the index on cell_list and the val from cell_values

                worksheet.update_cells(cell_list)
    print('Done!!!')

if __name__ == '__main__':
    add_data(letras())