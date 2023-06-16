# import os
# import numpy as np
# import pandas as pd
# from openpyxl import load_workbook, Workbook
# from openpyxl.utils import get_column_letter
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import PatternFill, Font, Color, Alignment
# from openpyxl.styles.differential import DifferentialStyle
# from openpyxl.formatting.rule import Rule, ColorScaleRule
from funcoes_desbalanco_neutro import *


# %%
omega = 2*np.pi*60
tensao_nominal_fase_fase = 69e3
Vff = tensao_nominal_fase_fase*np.sqrt(3)*np.exp(+1j*np.pi/6)
a = 1*np.exp(-1j*2*np.pi/3)

nr_lin_int = 2
nr_col_int = 2

nr_lin_ext = 5
nr_col_ext = 8

potencia_nominal_trifásica = 100e6

tensao_fase_entro = tensao_nominal_fase_fase/np.sqrt(3)
reatancia = tensao_nominal_fase_fase**2 / potencia_nominal_trifásica
corrente_nominal_lata = tensao_fase_entro / reatancia

cap_interna = 1/(omega*reatancia)
des_int = 0.33

# %% gerar matriz no excel
wb, matriz = matriz_fases_ramos(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, cap_interna=cap_interna, des_int=des_int)
wb.save('matriz_total.xlsx')

# %% ler matriz do excel, depois de editada manualmente
data = load_excel_to_numpy('matriz_total.xlsx')
matriz_A = data['A']
matriz_B = data['B']
matriz_C = data['C']

matriz_A1 = matriz_A[:, :nr_col_ext*nr_col_int]
matriz_A2 = matriz_A[:, nr_col_ext*nr_col_int:]
super_matriz_A1, eq_paral_internos_A1, eq_serie_internos_A1 = matrizes_internas(matriz_FR=matriz_A1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_A2, eq_paral_internos_A2, eq_serie_internos_A2 = matrizes_internas(matriz_FR=matriz_A2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_A1 = np.sum(eq_serie_internos_A1, axis=1)
eq_serie_externos_A1 = 1 / np.sum(1 / eq_paral_externos_A1)
eq_paral_externos_A2 = np.sum(eq_serie_internos_A2, axis=1)
eq_serie_externos_A2 = 1 / np.sum(1 / eq_paral_externos_A2)
eq_paral_ramos_A = eq_serie_externos_A1 + eq_serie_externos_A2


matriz_B1 = matriz_B[:, :nr_col_ext*nr_col_int]
matriz_B2 = matriz_B[:, nr_col_ext*nr_col_int:]
super_matriz_B1, eq_paral_internos_B1, eq_serie_internos_B1 = matrizes_internas(matriz_FR=matriz_B1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_B2, eq_paral_internos_B2, eq_serie_internos_B2 = matrizes_internas(matriz_FR=matriz_B2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_B1 = np.sum(eq_serie_internos_B1, axis=1)
eq_serie_externos_B1 = 1 / np.sum(1 / eq_paral_externos_B1)
eq_paral_externos_B2 = np.sum(eq_serie_internos_B2, axis=1)
eq_serie_externos_B2 = 1 / np.sum(1 / eq_paral_externos_B2)
eq_paral_ramos_B = eq_serie_externos_B1 + eq_serie_externos_B2


matriz_C1 = matriz_C[:, :nr_col_ext*nr_col_int]
matriz_C2 = matriz_C[:, nr_col_ext*nr_col_int:]
super_matriz_C1, eq_paral_internos_C1, eq_serie_internos_C1 = matrizes_internas(matriz_FR=matriz_C1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_C2, eq_paral_internos_C2, eq_serie_internos_C2 = matrizes_internas(matriz_FR=matriz_C2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_C1 = np.sum(eq_serie_internos_C1, axis=1)
eq_serie_externos_C1 = 1 / np.sum(1 / eq_paral_externos_C1)
eq_paral_externos_C2 = np.sum(eq_serie_internos_C2, axis=1)
eq_serie_externos_C2 = 1 / np.sum(1 / eq_paral_externos_C2)
eq_paral_ramos_C = eq_serie_externos_C1 + eq_serie_externos_C2


Za = 1 / (1j*omega*eq_paral_ramos_A)
Zb = 1 / (1j*omega*eq_paral_ramos_B)
Zc = 1 / (1j*omega*eq_paral_ramos_C)

matriz_impedancia_sistema = np.array([[Za, 0, 0], [0, Zb, 0], [0, 0, Zc]])

matriz_impedancia_malha = np.array([[Za, -Zb, 0], [0, Zb, -Zc], [1, 1, 1]])
matriz_fontes_malha = np.array([[Vff], [Vff*a], [0]])
matriz_correntes_fase = np.linalg.inv(matriz_impedancia_malha) @ matriz_fontes_malha

matriz_tensoes_Vabco = matriz_impedancia_sistema @ matriz_correntes_fase
tensao_deslocamento_netro = np.sum(matriz_tensoes_Vabco) / 3

V_ao = matriz_tensoes_Vabco[0]
V_bo = matriz_tensoes_Vabco[1]
V_co = matriz_tensoes_Vabco[2]

I_ao = matriz_correntes_fase[0]
I_bo = matriz_correntes_fase[1]
I_co = matriz_correntes_fase[2]

# %% Começa o caminho inverso
# %% Fase A

df_Ca_eq_ext = pd.DataFrame([eq_serie_externos_A1, eq_serie_externos_A2])
df_Ca_pa_ext = pd.DataFrame([eq_paral_externos_A1, eq_paral_externos_A2])
df_Ca_eq_int = pd.DataFrame(np.append(eq_serie_internos_A1, eq_serie_internos_A2, axis=1))


with pd.ExcelWriter('capacitancias_nas_latas.xlsx', engine='openpyxl') as writer:
    df_Ca_eq_ext.T.to_excel(writer, sheet_name='df_Ca_eq_ext', index=False, header=False)
    df_Ca_pa_ext.T.to_excel(writer, sheet_name='df_Ca_pa_ext', index=False, header=False)
    df_Ca_eq_int.to_excel(writer, sheet_name='df_Ca_eq_int', index=False, header=False)

# %%
I_a1 = V_ao * (1j*omega*eq_serie_externos_A1)
V_a1_ser = I_a1 * 1 / (1j*omega*eq_paral_externos_A1)
V_a1_ser = V_a1_ser.reshape(-1, 1) # necessário para transformar em matriz com uma coluna
I_a1_par = V_a1_ser * (1j*omega*eq_serie_internos_A1)

I_a2 = V_ao * (1j*omega*eq_serie_externos_A2)
V_a2_ser = I_a2 * 1 / (1j*omega*eq_paral_externos_A2)
V_a2_ser = V_a2_ser.reshape(-1, 1) # necessário para transformar em matriz com uma coluna
I_a2_par = V_a2_ser * (1j*omega*eq_serie_internos_A2)

df_I_a1_par = pd.DataFrame(np.abs(I_a1_par))
with pd.ExcelWriter('correntes_nas_latas.xlsx', engine='openpyxl') as writer:
    df_I_a1_par.to_excel(writer, sheet_name='df_I_a1_par', index=False, header=False)

destaca_maiores_que_nominal(planilha='df_I_a1_par', aquivo='correntes_nas_latas.xlsx', valor_nominal=corrente_nominal_lata)
# %%

#
#
#
# # Valores iniciais do Y equivalente dos dois ramos
# Iao = matriz_correntes_fase[0]
# Vao = matriz_tensoes_Vabco[0]
#
# # Separando por ramos
# Ia1 = Vao / (1j*omega*eq_serie_externos_A1)
# Ia2 = Vao / (1j*omega*eq_serie_externos_A2)
#
#
#
#
#
#
#
#
#
#
# # Tensoes em cada capacitor série equivalente
# Va1_paral = Ia1 / (1j*eq_paral_externos_A1)
# Va2_paral = Ia2 / (1j*eq_paral_externos_A1)
#
# # Cada capacitor série é o equivalente de várias latas em paralelo
# Ia1_paral = Va1_paral * (1j*omega*eq_serie_internos_A1)
# Ia2_paral = Va2_paral * (1j*omega*eq_serie_internos_A1)
#
# Va1_serie_1 = Ia1_paral / (1j*omega*eq_serie_internos_A1)
# Va1_serie_2 = Ia2_paral / (1j*omega*eq_serie_internos_A1)
#
#
# tensoes_latas_A = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# tensoes_latas_A[:, :nr_col_ext] = Va1_paral
# tensoes_latas_A[:, nr_col_ext:] = Va2_paral
# df_A = pd.DataFrame(np.abs(tensoes_latas_A))
#
#
# #  %%
#
# # # %% Juntando as três fases e colocando no excel
# # # %% tensoes
# # tensoes_latas_A = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # tensoes_latas_A[:, :nr_col_ext] = Va1_paral
# # tensoes_latas_A[:, nr_col_ext:] = Va2_paral
# # df_A = pd.DataFrame(np.abs(tensoes_latas_A))
# #
# # tensoes_latas_B = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # tensoes_latas_B[:, :nr_col_ext] = Vb1_paral
# # tensoes_latas_B[:, nr_col_ext:] = Vb2_paral
# # df_B = pd.DataFrame(np.abs(tensoes_latas_B))
# # print(df_B)
# #
# # tensoes_latas_C = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # tensoes_latas_C[:, :nr_col_ext] = Vc1_paral
# # tensoes_latas_C[:, nr_col_ext:] = Vc2_paral
# # df_C = pd.DataFrame(np.abs(tensoes_latas_C))
# #
# #
# # with pd.ExcelWriter('tensoes_nas_latas.xlsx', engine='openpyxl') as writer:
# #     df_A.to_excel(writer, sheet_name='Va', index=False, header=False)
# #     df_B.to_excel(writer, sheet_name='Vb', index=False, header=False)
# #     df_C.to_excel(writer, sheet_name='Vc', index=False, header=False)
# #
# # destaca_maiores_que_nominal(planilha='Va', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)
# # destaca_maiores_que_nominal(planilha='Vb', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)
# # destaca_maiores_que_nominal(planilha='Vc', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)
# #
# # # %% correntes
# # correntes_latas_A = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # correntes_latas_A[:, :nr_col_ext] = Ia1_paral
# # correntes_latas_A[:, nr_col_ext:] = Ia2_paral
# # df_A = pd.DataFrame(np.abs(correntes_latas_A))
# #
# # correntes_latas_B = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # correntes_latas_B[:, :nr_col_ext] = Ib1_paral
# # correntes_latas_B[:, nr_col_ext:] = Ib2_paral
# # df_B = pd.DataFrame(np.abs(correntes_latas_B))
# #
# # correntes_latas_C = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
# # correntes_latas_C[:, :nr_col_ext] = Ic1_paral
# # correntes_latas_C[:, nr_col_ext:] = Ic2_paral
# # df_C = pd.DataFrame(np.abs(correntes_latas_C))
# #
# # with pd.ExcelWriter('correntes_nas_latas.xlsx', engine='openpyxl') as writer:
# #     df_A.to_excel(writer, sheet_name='Ia', index=False, header=False)
# #     df_B.to_excel(writer, sheet_name='Ib', index=False, header=False)
# #     df_C.to_excel(writer, sheet_name='Ic', index=False, header=False)
#
#
# destaca_maiores_que_nominal(planilha='Ia', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)
# destaca_maiores_que_nominal(planilha='Ib', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)
# destaca_maiores_que_nominal(planilha='Ic', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)
#
#
