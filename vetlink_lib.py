import numpy as np
import pandas as pd

def read_file(file_path):

    data_file = pd.read_excel(file_path);

    return data_file;

def calculate(data_file, type_class):
    """
    Lê um arquivo Excel e calcula as somas das entradas e saídas de dinheiro.

    Args:
    file_path (str): Caminho para o arquivo Excel.

    Returns:
    tuple: Soma das entradas e soma das saídas.
    """

    data_file_filtered = data_file[data_file["Classe"]==type_class]

    column_accounting = data_file_filtered["Valor"].sum();

    processed_data = {type_class: column_accounting};

    return processed_data;

def export_to_execel(data_frame, file_name):
    try:
        data_frame.to_excel(file_name, index=False)
        return True
    except Exception as e:
        return False