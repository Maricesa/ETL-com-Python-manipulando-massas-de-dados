import pandas as pd
import glob
import os

#caminho dos arquivos excel
folder_path = 'make_data_files\\src\\data\\raw'

#lista todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder_path,'*.xlsx'))

if not excel_files:
    print("Nenhum arquivo compatível encontrado!")

else:
    #dataframe = tabela na memória para guardar os arquivos
    dfs = []

    for excel_file in excel_files:
        try:
            #leio o arquivo do excel
            df_temp = pd.read_excel(excel_file)
           
           #pego o nome do arquivo
            file_name = os.path.basename(excel_file)

            #crio uma nova coluna chamada location

            if 'brasil' in file_name.lower():
               df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it' 
            

            #crio uma nova coluna chamada campanha

            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            print(df_temp)

            #guarda os dados tratados dentro de uma dataframe comum
            dfs.append(df_temp)

        except Exception as e:
          print(f"Erro ao ler o arquivo {excel_file}: {e}")
    
    if dfs:

        #concatena todas as tabelas salvas no dfs em uma única tabela
        result = pd.concat(dfs,ignore_index=True)

        #caminho de saída
        output_file = os.path.join('make_data_files','src','data','ready','clean.xlsx')

        #configura o motor de escrita
        writer = pd.ExcelWriter(output_file,engine='xlsxwriter')

        #leva os dados do resultado a serem escritos em um motor de excel configurado
        result.to_excel(writer, index=False)

        #salva o arquivo excel
        writer._save()
    else:
        print("Nenhum dado a ser salvo")
