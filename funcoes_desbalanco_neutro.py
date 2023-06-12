import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill





def gera_matriz_interna(nr_lin_int=1, nr_col_int=1, cap_interna=999.9, des_int=0):
    matriz_interna = cap_interna * np.random.uniform(1 - des_int, 1 + des_int, (nr_lin_int, nr_col_int))
    return matriz_interna





#
#
#
# def matriz_capacitancias_externas(nr_lin_ext=1, nr_col_ext=1, cap_interna=999.9, des_int=0, nr_lin_int=1, nr_col_int=1):
#     matriz_externa = np.zeros((nr_lin_ext, nr_col_ext), dtype=object)
#     for i_ext in range(nr_lin_ext):
#         for j_ext in range(nr_col_ext):
#             matriz_interna = gera_matriz_interna(nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, cap_interna=cap_interna, des_int=des_int)
#             matriz_externa[i_ext, j_ext] = matriz_interna
#
#     return matriz_externa
#
# def remove_arquivo(file_name='meu_arquivo.xlsx'):
#     # Verifica se o arquivo existe antes de excluí-lo
#     if os.path.exists(file_name):
#         try:
#             # Tente abrir o arquivo no modo de atualização
#             with open(file_name, 'r+'):
#                 pass
#         except IOError:
#             print("O arquivo está aberto. Feche-o antes de continuar.")
#             return
#
#         os.remove(file_name)
#         print("Arquivo excluído com sucesso.")
#     else:
#         print("O arquivo especificado não existe.")
#
#
# def prenche_valores(matriz_a1=np.ones((4,4)), nr_lin_ext=2, nr_col_ext=2, cap_interna=999.9, des_int=2, nr_lin_int=2, nr_col_int=2, fase='erro', ramo=99, filename='meu_arquivo.xlsx'):
#
#     dataframes = {}
#     sheet_name = f'{fase}_{ramo}'
#
#     # Verifique se o arquivo existe
#     if not os.path.exists(filename):
#         wb = Workbook()
#         wb.save(filename)
#
#     with pd.ExcelWriter('meu_arquivo.xlsx', engine='openpyxl', mode='a') as writer:
#
#         for i_ext in range(nr_lin_ext):
#             for j_ext in range(nr_col_ext):
#                 df_matriz_interna = pd.DataFrame(matriz_a1[i_ext][j_ext])
#                 dataframes[f'df_matriz_interna_{fase}_{ramo}'] = df_matriz_interna
#                 startrow = i_ext * nr_lin_int + i_ext
#                 startcol = j_ext * nr_col_int + j_ext
#                 df_matriz_interna.to_excel(writer, header=False, index=False, startrow=startrow, startcol=startcol, sheet_name=f'{fase}_{ramo}_{i_ext}-{j_ext}')
#
#     wb = load_workbook(filename)
#     # Crie uma nova planilha
#     sheet_final = wb.create_sheet(title=sheet_name)
#
#     # Acesse todas as planilhas
#     sheets = wb.sheetnames
#
#     # Para cada planilha...
#     for sheet in sheets:
#         # Se o nome da planilha começa com 'fase_ramo'...
#         if sheet.startswith(f'{fase}_{ramo}'):
#             # Acesse essa planilha
#             current_sheet = wb[sheet]
#             # Para cada linha na planilha...
#             for row in range(1, current_sheet.max_row + 1):
#                 # Para cada coluna na planilha...
#                 for col in range(1, current_sheet.max_column + 1):
#                     # Se a célula na planilha final ainda não contém um valor, defina-o como 0
#                     if sheet_final.cell(row=row, column=col).value is None:
#                         sheet_final.cell(row=row, column=col).value = 0
#                     # Adicione o valor da célula atual ao valor correspondente na planilha final
#                     if current_sheet.cell(row=row, column=col).value is not None:
#                         sheet_final.cell(row=row, column=col).value = sheet_final.cell(row=row, column=col).value + current_sheet.cell(row=row, column=col).value
#                         print(sheet_final.cell)
#                         current_sheet.sheet_state = 'hidden'
#
#
#     wb.save(filename)
#
#
#     wb = load_workbook('meu_arquivo.xlsx')
#
#
#     ws1 = wb[f'{fase}_{ramo}']
#     wsABC = wb.create_sheet(f'{fase}_{ramo}')
#
#
#     for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, min_col=1, max_col=ws1.max_column):
#         for cell in row:
#             wsABC.cell(row=cell.row, column=ramo*cell.column, value=cell.value)
#
#
#
#     wb.save(filename)
#     wb = load_workbook(filename)
#     ws = wb[f'{fase}_{ramo}']
#     # Definir o padrão de preenchimento para preto
#     black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
#
#     # Percorrer todas as células na planilha
#     for row in ws.iter_rows():
#         for cell in row:
#             # Se a célula contém o valor 0, altere a cor de fundo para preto
#             if cell.value == 0:
#                 cell.fill = black_fill
#
#     # Salvar a planilha
#     wb.save(filename)
#
#
# nr_lin_ext=4
# nr_col_ext=4
# nr_lin_int=2
# nr_col_int=2
# cap_interna=10.0
# des_int=1.1 #fator a multpilicar
#
# remove_arquivo(file_name='meu_arquivo.xlsx')
#
# matriz_a1 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
# matriz_a2 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
#
# matriz_b1 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
# matriz_b2 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
#
# matriz_c1 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
# matriz_c2 = matriz_capacitancias_externas(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int)
#
# prenche_valores(matriz_a1=matriz_a1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='a', ramo=1, filename='meu_arquivo.xlsx')
# prenche_valores(matriz_a1=matriz_a2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='a', ramo=2, filename='meu_arquivo.xlsx')
#
# prenche_valores(matriz_a1=matriz_b1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='b', ramo=1, filename='meu_arquivo.xlsx')
# prenche_valores(matriz_a1=matriz_b2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='b', ramo=2, filename='meu_arquivo.xlsx')
#
# prenche_valores(matriz_a1=matriz_c1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='c', ramo=1, filename='meu_arquivo.xlsx')
# prenche_valores(matriz_a1=matriz_c2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, cap_interna=cap_interna, des_int=des_int, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, fase='c', ramo=2, filename='meu_arquivo.xlsx')
# # %%
# def calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=9, fase='fase_erro', file_name = 'meu_arquivo.xlsx'):
#     wb = load_workbook(file_name)
#     ws = wb[fase]
#     # Definir o intervalo de células usando números
#     start_row = 1 + nr_lin_int+i_ext
#     start_col = 1 + nr_col_int*j_ext
#     end_row = 1 + nr_lin_int+i_ext + nr_lin_int
#     end_col = 1 + nr_col_int*j_ext + nr_col_int
#
#     my_list_int = []
#     for row in range(start_row, end_row + 1):
#         for col in range(start_col, end_col + 1):
#             my_list_int.append(ws.cell(row=row, column=col).value)
#
#     matriz_interna_lida = np.array(my_list_int)
#     paralelos_matriz_interna = matriz_interna_lida.sum(axis=0)
#     series_matriz_interna = 1.0 / np.sum(np.reciprocal(paralelos_matriz_interna))
#     return series_matriz_interna
#
#
# def calcula_capacitancias_internas_equivalentes():
#     matriz_externa_calculada_fase_a_ramo_1 = np.ones((nr_lin_ext, nr_col_ext))
#     matriz_externa_calculada_fase_a_ramo_2 = np.ones((nr_lin_ext, nr_col_ext))
#
#     matriz_externa_calculada_fase_b_ramo_1 = np.ones((nr_lin_ext, nr_col_ext))
#     matriz_externa_calculada_fase_b_ramo_2 = np.ones((nr_lin_ext, nr_col_ext))
#
#     matriz_externa_calculada_fase_c_ramo_1 = np.ones((nr_lin_ext, nr_col_ext))
#     matriz_externa_calculada_fase_c_ramo_2 = np.ones((nr_lin_ext, nr_col_ext))
#
#     for i_ext in range(nr_lin_ext):
#         for j_ext in range(nr_col_ext):
#             matriz_externa_calculada_fase_a_ramo_1[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=1, fase='fase_a')
#             matriz_externa_calculada_fase_a_ramo_2[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=2, fase='fase_a')
#
#             matriz_externa_calculada_fase_b_ramo_1[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=1, fase='fase_b')
#             matriz_externa_calculada_fase_b_ramo_2[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=2, fase='fase_b')
#
#             matriz_externa_calculada_fase_c_ramo_1[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=1, fase='fase_c')
#             matriz_externa_calculada_fase_c_ramo_2[i_ext, j_ext] = calcula_capacitancias_equivalentes(i_ext, j_ext, ramo=2, fase='fase_c')
#
#
#     df_externa_calculada_fase_a_ramo_1 = pd.DataFrame(matriz_externa_calculada_fase_a_ramo_1)
#     df_externa_calculada_fase_a_ramo_2 = pd.DataFrame(matriz_externa_calculada_fase_a_ramo_2)
#
#     df_externa_calculada_fase_b_ramo_1 = pd.DataFrame(matriz_externa_calculada_fase_b_ramo_1)
#     df_externa_calculada_fase_b_ramo_2 = pd.DataFrame(matriz_externa_calculada_fase_b_ramo_2)
#
#     df_externa_calculada_fase_c_ramo_1 = pd.DataFrame(matriz_externa_calculada_fase_c_ramo_1)
#     df_externa_calculada_fase_c_ramo_2 = pd.DataFrame(matriz_externa_calculada_fase_c_ramo_2)
#
#
#     with pd.ExcelWriter('capacitancias_internas_equivalentes.xlsx') as writer:
#         df_externa_calculada_fase_a_ramo_1.to_excel(writer, sheet_name='fase_a_ramo_1', index=False, header=False, startrow=0, startcol=0)
#         df_externa_calculada_fase_a_ramo_2.to_excel(writer, sheet_name='fase_a_ramo_2', index=False, header=False, startrow=0, startcol=nr_col_ext + 1)
#         df_externa_calculada_fase_b_ramo_1.to_excel(writer, sheet_name='fase_b_ramo_1', index=False, header=False, startrow=0, startcol=0)
#         df_externa_calculada_fase_b_ramo_2.to_excel(writer, sheet_name='fase_b_ramo_2', index=False, header=False, startrow=0, startcol=nr_col_ext + 1)
#         df_externa_calculada_fase_c_ramo_1.to_excel(writer, sheet_name='fase_c_ramo_1', index=False, header=False, startrow=0, startcol=0)
#         df_externa_calculada_fase_c_ramo_2.to_excel(writer, sheet_name='fase_c_ramo_2', index=False, header=False, startrow=0, startcol=nr_col_ext + 1)
#
#
#     return [matriz_externa_calculada_fase_a_ramo_1, matriz_externa_calculada_fase_a_ramo_2, matriz_externa_calculada_fase_b_ramo_1, matriz_externa_calculada_fase_b_ramo_2, matriz_externa_calculada_fase_c_ramo_1, df_externa_calculada_fase_c_ramo_2]
#
#
# def calcula_capacitancias_externas_equivalentes_por_ramo():
#
#     mecfar1, mecfar2, mecfbr1, mecfbr2, mecfcr1, mecfcr2 = calcula_capacitancias_internas_equivalentes()
#
#     matriz_externa_calculada_fase_a_ramo_1_equiv_paralelo = mecfar1.sum(axis=0)
#     matriz_externa_calculada_fase_a_ramo_1_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_a_ramo_1_equiv_paralelo))
#     capac_equiv_a1 = matriz_externa_calculada_fase_a_ramo_1_equiv_serie
#     matriz_externa_calculada_fase_a_ramo_2_equiv_paralelo = mecfar1.sum(axis=0)
#     matriz_externa_calculada_fase_a_ramo_2_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_a_ramo_2_equiv_paralelo))
#     capac_equiv_a2 = matriz_externa_calculada_fase_a_ramo_2_equiv_serie
#
#     matriz_externa_calculada_fase_b_ramo_1_equiv_paralelo = mecfbr1.sum(axis=0)
#     matriz_externa_calculada_fase_b_ramo_1_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_b_ramo_1_equiv_paralelo))
#     capac_equiv_b1 = matriz_externa_calculada_fase_b_ramo_1_equiv_serie
#     matriz_externa_calculada_fase_b_ramo_2_equiv_paralelo = mecfbr1.sum(axis=0)
#     matriz_externa_calculada_fase_b_ramo_2_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_b_ramo_2_equiv_paralelo))
#     capac_equiv_b2 = matriz_externa_calculada_fase_b_ramo_2_equiv_serie
#
#     matriz_externa_calculada_fase_c_ramo_1_equiv_paralelo = mecfcr1.sum(axis=0)
#     matriz_externa_calculada_fase_c_ramo_1_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_c_ramo_1_equiv_paralelo))
#     capac_equiv_c1 = matriz_externa_calculada_fase_c_ramo_1_equiv_serie
#     matriz_externa_calculada_fase_c_ramo_2_equiv_paralelo = mecfcr1.sum(axis=0)
#     matriz_externa_calculada_fase_c_ramo_2_equiv_serie = 1.0 / np.sum(np.reciprocal(matriz_externa_calculada_fase_c_ramo_2_equiv_paralelo))
#     capac_equiv_c2 = matriz_externa_calculada_fase_c_ramo_2_equiv_serie
#
#     return [capac_equiv_a1, capac_equiv_a2, capac_equiv_b1, capac_equiv_b2, capac_equiv_c1, capac_equiv_c2]
#
#
# ce_a1, ce_a2, ce_b1, ce_b2, ce_c1, ce_c2 = calcula_capacitancias_externas_equivalentes_por_ramo()
# # %%
# # Acesse todas as planilhas
#
# file_name='capacitancias_internas_equivalentes.xlsx'
# fase = 'a'
# wb = load_workbook(file_name)
# sheet_final = wb.create_sheet(title=f'fase_{fase}')
# sheets = wb.sheetnames
#
# # Para cada planilha...
# for sheet in sheets:
#     # Se o nome da planilha começa com 'fase_ramo'...
#     if sheet.startswith(f'fase_{fase}'):
#         # Acesse essa planilha
#         current_sheet = wb[sheet]
#         # Para cada linha na planilha...
#         for row in range(1, current_sheet.max_row + 1):
#             # Para cada coluna na planilha...
#             for col in range(1, current_sheet.max_column + 1):
#                 # Se a célula na planilha final ainda não contém um valor, defina-o como 0
#                 if sheet_final.cell(row=row, column=col).value is None:
#                     sheet_final.cell(row=row, column=col).value = 0
#                 # Adicione o valor da célula atual ao valor correspondente na planilha final
#                 if current_sheet.cell(row=row, column=col).value is not None:
#                     sheet_final.cell(row=row, column=col).value = sheet_final.cell(row=row,
#                                                                                    column=col).value + current_sheet.cell(
#                         row=row, column=col).value
#                     print(sheet_final.cell)
#                     current_sheet.sheet_state = 'hidden'
#
# wb.save(file_name)