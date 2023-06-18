# import os
# import numpy as np
# import pandas as pd
# from openpyxl import load_workbook, Workbook
# from openpyxl.utils import get_column_letter
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import PatternFill, Font, Color, Alignment
# from openpyxl.styles.differential import DifferentialStyle
# from openpyxl.formatting.rule import Rule, ColorScaleRule
from engineering_notation import EngNumber
from funcoes_desbalanco_neutro import *


# %%

a = 1*np.exp(-1j*2*np.pi/3)

nr_lin_int = 1
nr_col_int = 1

nr_lin_ext = 1
nr_col_ext = 1

potencia_nominal_trifásica = 10e6
omega = 2*np.pi*60
tensao_nominal_fase_fase = 69e3
Vff = tensao_nominal_fase_fase*np.sqrt(3)*np.exp(+1j*np.pi/6)
tensao_fase_neutro = tensao_nominal_fase_fase/np.sqrt(3)
corrente_fase_neutro = ( potencia_nominal_trifásica/3 ) / tensao_fase_neutro
reatancia = tensao_fase_neutro / corrente_fase_neutro
corrente_nominal_lata = corrente_fase_neutro / nr_col_ext

cap_interna = 1 / (omega*reatancia)
des_int = 0.0

print("corrente_nominal_lata = corrente_fase_neutro / nr_col_ext = ", corrente_nominal_lata)
print("cap_interna = 1 / (omega*reatancia) = ", EngNumber(cap_interna))

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
I_a1, V_a1_ser, I_a1_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_A1, eq_paral_externos_A1, eq_serie_internos_A1)
I_a2, V_a2_ser, I_a2_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_A2, eq_paral_externos_A2, eq_serie_internos_A2)
I_a_par = np.hstack((I_a1_par, I_a2_par))
df_I_a_par = pd.DataFrame(np.abs(I_a_par))

I_b1, V_b1_ser, I_b1_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_B1, eq_paral_externos_B1, eq_serie_internos_B1)
I_b2, V_b2_ser, I_b2_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_B2, eq_paral_externos_B2, eq_serie_internos_B2)
I_b_par = np.hstack((I_b1_par, I_b2_par))
df_I_b_par = pd.DataFrame(np.abs(I_b_par))

I_c1, V_c1_ser, I_c1_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_C1, eq_paral_externos_C1, eq_serie_internos_C1)
I_c2, V_c2_ser, I_c2_par = calcular_corrente_tensao(V_ao, omega, eq_serie_externos_C2, eq_paral_externos_C2, eq_serie_internos_C2)
I_c_par = np.hstack((I_c1_par, I_c2_par))
df_I_c_par = pd.DataFrame(np.abs(I_c_par))

with pd.ExcelWriter('correntes_nas_latas.xlsx', engine='openpyxl') as writer:
    df_I_a_par.to_excel(writer, sheet_name='df_I_a_par', index=False, header=False)
    df_I_b_par.to_excel(writer, sheet_name='df_I_b_par', index=False, header=False)
    df_I_c_par.to_excel(writer, sheet_name='df_I_c_par', index=False, header=False)

destaca_maiores_que_nominal(planilha='df_I_a_par', aquivo='correntes_nas_latas.xlsx', valor_nominal=corrente_nominal_lata)
destaca_maiores_que_nominal(planilha='df_I_b_par', aquivo='correntes_nas_latas.xlsx', valor_nominal=corrente_nominal_lata)
destaca_maiores_que_nominal(planilha='df_I_c_par', aquivo='correntes_nas_latas.xlsx', valor_nominal=corrente_nominal_lata)
# %%

