import vetlink_lib as vet  # Importa a biblioteca vetlink_lib como vet
import PySimpleGUI as sg  # Importa a biblioteca PySimpleGUI como sg
import os  # Importa o módulo os para lidar com funcionalidades de sistema operacional

# Define o layout da janela usando PySimpleGUI
layout = [
    [sg.Combo(sorted(sg.user_settings_get_entry('-filenames-', [])), default_value=sg.user_settings_get_entry('-last filename-', ''), size=(50, 1), key='-FILENAME-'), sg.FileBrowse(), sg.B('Clear History')],
    [sg.Button('Ok', bind_return_key=True), sg.Button('Cancel')]
]

# Cria a janela principal com título 'Aplicativo Finaceiro VetLink' e o layout definido
window = sg.Window('Aplicativo Finaceiro VetLink', layout)

# Define uma lista padrão de tipos de classes e inicializa variáveis para cálculos
list_type_class_default = ["Realização Consulta", "Venda Vacinas", "Venda Testes Rápidos", "Realização Cirurgia", "Realização Exames"]
list_element_calculated = []
total_cost = 0
total_profit = 0
default_file_name = "Análise Completa.xlsx"

# Loop principal da aplicação para interação com o usuário
while True:
    event, values = window.read()  # Aguarda interação do usuário e retorna o evento e valores associados

    if event in (sg.WIN_CLOSED, 'Cancel'):  # Verifica se o usuário fechou a janela ou clicou em 'Cancel'
        break  # Encerra o loop e fecha a janela

    if event == 'Ok':  # Se o evento for 'Ok' (quando o botão Ok é pressionado)
        selected_file = values['-FILENAME-']  # Obtém o arquivo selecionado pelo usuário
        sg.user_settings_set_entry('-filenames-', list(set(sg.user_settings_get_entry('-filenames-', []) + [selected_file, ])))  # Atualiza o histórico de arquivos usados
        sg.user_settings_set_entry('-last filename-', selected_file)  # Define o último arquivo utilizado como o selecionado agora

        data_file = vet.read_file(selected_file)  # Lê o arquivo de dados usando uma função da biblioteca vetlink_lib
        list_type_class = list(set(data_file["Classe"]))  # Obtém uma lista de classes únicas a partir dos dados lidos

        for type in list_type_class:
            element_calculated = vet.calculate(data_file, type)  # Calcula elementos específicos com base nas classes
            list_element_calculated.append(element_calculated)  # Adiciona os elementos calculados a uma lista

            if type in list_type_class_default:
                total_profit = element_calculated[type] + total_profit  # Calcula o lucro total
            else:
                total_cost = element_calculated[type] + total_cost  # Calcula o custo total

        list_element_calculated.append({"Gastos Totais": total_cost})  # Adiciona o total de custos à lista de elementos calculados
        list_element_calculated.append({"Lucros Totais": total_profit})  # Adiciona o total de lucros à lista de elementos calculados

        # Prepara os dados para serem exportados para um arquivo Excel
        data = {}
        for element in list_element_calculated:
            for key, value in element.items():
                data[key] = [value]

        original_source_path = selected_file
        path = os.path.dirname(original_source_path)  # Obtém o diretório do arquivo selecionado
        new_path = os.path.join(path, default_file_name)  # Define o caminho onde o novo arquivo será salvo

        data_frame = vet.pd.DataFrame(data)  # Converte os dados para um DataFrame usando a biblioteca pandas

        validation_test = vet.export_to_excel(data_frame, new_path)  # Exporta os dados para um arquivo Excel usando uma função da biblioteca vetlink_lib

        if validation_test:
            message = f"Arquivo salvo na pasta '{path}'"  # Exibe uma mensagem de sucesso caso a exportação seja bem-sucedida
            sg.popup(message)
            break  # Encerra o loop e fecha a janela após a exportação

        else:
            sg.popup("Falha ao realizar o cálculo. Tente novamente")  # Exibe uma mensagem de erro se houver problemas na exportação

    elif event == 'Clear History':  # Se o evento for 'Clear History' (limpar histórico)
        sg.user_settings_set_entry('-filenames-', [])  # Limpa o histórico de arquivos utilizados
        sg.user_settings_set_entry('-last filename-', '')  # Define o último arquivo utilizado como vazio
        window['-FILENAME-'].update(values=[], value='')  # Atualiza o valor do campo de seleção de arquivo na interface

window.close()  # Fecha a janela ao final do loop
