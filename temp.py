import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Color, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, ColorScaleRule
from funcoes_desbalanco_neutro import *


# %%
def destaca_maiores_que_nominal(planilha='Ia1', aquivo='correntes_nas_latas.xlsx', valor_nominal=10.1):
    wb = load_workbook(aquivo)
    ws = wb[planilha]
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    dxf = DifferentialStyle(fill=red_fill)
    rule = ColorScaleRule(start_type="min", start_color="FFFF99",
                          end_type="max", end_color="FF0000")
    ws.conditional_formatting.add("A1:Z26", rule)
    wb.save(aquivo)
# %%
def matriz_complexa_para_polar(matriz_complexa):
    magnitude = np.abs(matriz_complexa)
    angulo = np.angle(matriz_complexa)

    # Gera a matriz de strings em formato polar
    matriz_polar = np.empty(matriz_complexa.shape, dtype=object)

    for i in range(matriz_complexa.shape[0]):
        for j in range(matriz_complexa.shape[1]):
            matriz_polar[i, j] = f'{magnitude[i, j]:.2f}∠{np.degrees(angulo[i, j]):.2f}°'

    return matriz_polar

# %%
def load_excel_to_numpy(filename):
    wb = load_workbook(filename=filename, read_only=True)
    sheets = ['A', 'B', 'C']
    data = {}

    for sheet in sheets:
        ws = wb[sheet]
        data[sheet] = pd.DataFrame(ws.values)

    return {k: v.to_numpy() for k, v in data.items()}

# %%
def matriz_fases_ramos(nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int, cap_interna, des_int):
    matriz = np.ones((3, nr_lin_ext * nr_lin_int, 2 * nr_col_ext * nr_col_int))
    cores = ['FFFFCC', 'CCFFCC', 'CCCCFF', 'FFCCFF']

    wb = Workbook()
    ws = wb.active
    ws.title = "A"

    for fase in ['A', 'B', 'C']:
        if fase == 'A':
            ws = wb.active
            fase_cont = 0
        else:
            ws = wb.create_sheet(title=fase)
            fase_cont = fase_cont + 1

        for i_ext in range(nr_lin_ext):
            for j_ext in range(2*nr_col_ext):
                cor = cores[(i_ext + j_ext) % len(cores)]
                fill = PatternFill(start_color=cor, end_color=cor, fill_type='solid')
                for i_int in range(nr_lin_int):
                    for j_int in range(nr_col_int):
                        ii = i_ext*nr_lin_int + i_int
                        jj = j_ext*nr_col_int + j_int
                        valor = cap_interna * np.random.uniform(1 - des_int, 1 + des_int)
                        matriz[fase_cont, ii,jj] = valor
                        ws.cell(row=ii+1, column=jj+1, value=valor)
                        ws.cell(row=ii+1, column=jj+1).fill = fill
                        ws.cell(row=ii + 1, column=jj + 1).number_format = '##0.0'
                        ws.column_dimensions[get_column_letter(jj+1)].width = 5
                        ws.row_dimensions[ii + 1].height = 30
                        ws.cell(row=ii + 1, column=jj + 1).alignment  = Alignment(horizontal='center', vertical='center')
                        if j_ext >= nr_col_ext:
                            ws.cell(row=ii+1, column=jj+1).font = Font(name='Bahnschrift SemiBold Condensed', color="FF0000", underline="single")
                        else:
                            ws.cell(row=ii+1, column=jj+1).font = Font(name='Bahnschrift SemiBold Condensed', color="0000FF")

    return wb, matriz


# %%
def matrizes_internas(matriz_FR, nr_lin_ext, nr_col_ext, ramo=99):
    super_matriz_FR = np.ones((nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int))
    eq_paral_internos = np.ones((nr_lin_ext*nr_lin_int, nr_col_ext))
    eq_serie_internos = np.ones((nr_lin_ext, nr_col_ext))
    for i_ext in range(nr_lin_ext):
        for j_ext in range (nr_col_ext):
            row_sart = i_ext*nr_lin_int
            row_end = row_sart + nr_lin_int
            col_sart = j_ext * nr_col_int
            col_end = col_sart + nr_col_int
            # matriz interna temporária
            mat_temp = np.array(matriz_FR[row_sart:row_end, col_sart:col_end])
            super_matriz_FR[i_ext, j_ext] = mat_temp
            eq_paral_internos[i_ext*nr_lin_int:i_ext*nr_lin_int+nr_lin_int, j_ext] = np.sum(mat_temp, axis=1)
            eq_serie_internos[i_ext, j_ext] = 1 / np.sum(1 / eq_paral_internos[i_ext*nr_lin_int:i_ext*nr_lin_int+nr_lin_int, j_ext])


    return super_matriz_FR, eq_paral_internos, eq_serie_internos

# %%
omega = 2*np.pi*60
Vff = 100
a = 1*np.exp(-1j*2*np.pi/3)

nr_lin_int = 2
nr_col_int = 2

nr_lin_ext = 4
nr_col_ext = 4

cap_interna = 5
des_int = 0.5


wb, matriz = matriz_fases_ramos(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int, nr_col_int=nr_col_int, cap_interna=cap_interna, des_int=des_int)
wb.save('matriz_total.xlsx')

data = load_excel_to_numpy('matriz_total.xlsx')
matriz_A = data['A']
matriz_B = data['B']
matriz_C = data['C']

matriz_A1 = matriz_A[:, :nr_col_ext*nr_col_int]
matriz_A2 = matriz_A[:, nr_col_ext*nr_col_int:]
super_matriz_A1, eq_paral_internos_A1, eq_serie_internos_A1 = matrizes_internas(matriz_FR=matriz_A1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=1)
super_matriz_A2, eq_paral_internos_A2, eq_serie_internos_A2 = matrizes_internas(matriz_FR=matriz_A2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=2)
eq_paral_externos_A1 = np.sum(eq_serie_internos_A1, axis=1)
eq_serie_externos_A1 = 1 / np.sum(1 / eq_paral_externos_A1)
eq_paral_externos_A2 = np.sum(eq_serie_internos_A2, axis=1)
eq_serie_externos_A2 = 1 / np.sum(1 / eq_paral_externos_A2)
eq_paral_ramos_A = eq_serie_externos_A1 + eq_serie_externos_A2


matriz_B1 = matriz_B[:, :nr_col_ext*nr_col_int]
matriz_B2 = matriz_B[:, nr_col_ext*nr_col_int:]
super_matriz_B1, eq_paral_internos_B1, eq_serie_internos_B1 = matrizes_internas(matriz_FR=matriz_B1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=1)
super_matriz_B2, eq_paral_internos_B2, eq_serie_internos_B2 = matrizes_internas(matriz_FR=matriz_B2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=2)
eq_paral_externos_B1 = np.sum(eq_serie_internos_B1, axis=1)
eq_serie_externos_B1 = 1 / np.sum(1 / eq_paral_externos_B1)
eq_paral_externos_B2 = np.sum(eq_serie_internos_B2, axis=1)
eq_serie_externos_B2 = 1 / np.sum(1 / eq_paral_externos_B2)
eq_paral_ramos_B = eq_serie_externos_B1 + eq_serie_externos_B2


matriz_C1 = matriz_C[:, :nr_col_ext*nr_col_int]
matriz_C2 = matriz_C[:, nr_col_ext*nr_col_int:]
super_matriz_C1, eq_paral_internos_C1, eq_serie_internos_C1 = matrizes_internas(matriz_FR=matriz_C1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=1)
super_matriz_C2, eq_paral_internos_C2, eq_serie_internos_C2 = matrizes_internas(matriz_FR=matriz_C2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, ramo=2)
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

# %% Começa o caminho inverso
# %% Fase A

# Valores iniciais do Y equivalente dos dois ramos
Iao = matriz_correntes_fase[0]
Vao = matriz_tensoes_Vabco[0]

# Separando por ramos
Za1 = 1 / (1j*omega*eq_serie_externos_A1)
Za2 = 1 / (1j*omega*eq_serie_externos_A2)

Ia1 = Vao / Za1
Ia2 = Vao / Za2

# Separando por grupos em série equivalente
Za1_series = 1 / (1j*omega*eq_paral_externos_A1)
Za2_series = 1 / (1j*omega*eq_paral_externos_A1)

# Conterindo  se Va1_series = Va2_series = Vao
Va1_series = Ia1 * Za1_series
Va2_series = Ia2 * Za2_series

# Separando as latas paralelas de cada grupo série
Ya1_paralelo = (1j*eq_serie_internos_A1)
Ya2_paralelo = (1j*eq_serie_internos_A2)

# Conferindo se Ia1_paral + Ia2_paral = Iao
Ia1_paral = Va1_series * Ya1_paralelo
Ia2_paral = Va2_series * Ya2_paralelo

# Conferir se sum(Va1_paral) = sum(Va2_paral) = Vao
Va1_paral = Ia1_paral / Ya1_paralelo
Va2_paral = Ia2_paral / Ya2_paralelo

# %% Fase B

# Valores iniciais do Y equivalente dos dois ramos
Ibo = matriz_correntes_fase[1]
Vbo = matriz_tensoes_Vabco[1]

# Separando por ramos
Zb1 = 1 / (1j*omega*eq_serie_externos_B1)
Zb2 = 1 / (1j*omega*eq_serie_externos_B2)

Ib1 = Vbo / Zb1
Ib2 = Vbo / Zb2

# Separando por grupos em série equivalente
Zb1_series = 1 / (1j*omega*eq_paral_externos_B1)
Zb2_series = 1 / (1j*omega*eq_paral_externos_B1)

# Conterindo  se Vb1_series = Vb3_series = Vbo
Vb1_series = Ib1 * Zb1_series
Vb2_series = Ib2 * Zb2_series

# Separando as latas paralelas de cada grupo série
Yb1_paralelo = (1j*eq_serie_internos_B1)
Yb2_paralelo = (1j*eq_serie_internos_B2)

# Conferindo se Ib1_paral + Ib3_paral = Ibo
Ib1_paral = Vb1_series * Yb1_paralelo
Ib2_paral = Vb2_series * Yb2_paralelo

# Conferir se sum(Vb1_paral) = sum(Vb3_paral) = Vbo
Vb1_paral = Ib1_paral / Yb1_paralelo
Vb2_paral = Ib2_paral / Yb2_paralelo

# %% Fase C

# Valores iniciais do Y equivalente dos dois ramos
Ico = matriz_correntes_fase[2]
Vco = matriz_tensoes_Vabco[2]

# Separando por ramos
Zc1 = 1 / (1j*omega*eq_serie_externos_C1)
Zc2 = 1 / (1j*omega*eq_serie_externos_C2)

Ic1 = Vco / Zc1
Ic2 = Vco / Zc2

# Separando por grupos em série equivalente
Zc1_series = 1 / (1j*omega*eq_paral_externos_C1)
Zc2_series = 1 / (1j*omega*eq_paral_externos_C1)

# Conterindo  se Vc1_series = Vc2_series = Vco
Vc1_series = Ic1 * Zc1_series
Vc2_series = Ic2 * Zc2_series

# Separando as latas paralelas de cada grupo série
Yc1_paralelo = (1j*eq_serie_internos_C1)
Yc2_paralelo = (1j*eq_serie_internos_C2)

# Conferindo se Ic1_paral + Ic2_paral = Ico
Ic1_paral = Vc1_series * Yc1_paralelo
Ic2_paral = Vc2_series * Yc2_paralelo

# Conferir se sum(Vc1_paral) = sum(Vc2_paral) = Vco
Vc1_paral = Ic1_paral / Yc1_paralelo
Vc2_paral = Ic2_paral / Yc2_paralelo

# %% Juntando as três fases e colocando no excel
# %% tensoes
tensoes_latas_A = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
tensoes_latas_A[:, :nr_col_ext] = Va1_paral
tensoes_latas_A[:, nr_col_ext:] = Va2_paral
df_A = pd.DataFrame(np.abs(tensoes_latas_A))

tensoes_latas_B = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
tensoes_latas_B[:, :nr_col_ext] = Vb1_paral
tensoes_latas_B[:, nr_col_ext:] = Vb2_paral
df_B = pd.DataFrame(np.abs(tensoes_latas_B))
print(df_B)

tensoes_latas_C = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
tensoes_latas_C[:, :nr_col_ext] = Vc1_paral
tensoes_latas_C[:, nr_col_ext:] = Vc2_paral
df_C = pd.DataFrame(np.abs(tensoes_latas_C))


with pd.ExcelWriter('tensoes_nas_latas.xlsx', engine='openpyxl') as writer:
    df_A.to_excel(writer, sheet_name='Va', index=False, header=False)
    df_B.to_excel(writer, sheet_name='Vb', index=False, header=False)
    df_C.to_excel(writer, sheet_name='Vc', index=False, header=False)

destaca_maiores_que_nominal(planilha='Va', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)
destaca_maiores_que_nominal(planilha='Vb', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)
destaca_maiores_que_nominal(planilha='Vc', aquivo='tensoes_nas_latas.xlsx', valor_nominal=14.43)

# %% correntes
correntes_latas_A = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
correntes_latas_A[:, :nr_col_ext] = Ia1_paral
correntes_latas_A[:, nr_col_ext:] = Ia2_paral
df_A = pd.DataFrame(np.abs(correntes_latas_A))

correntes_latas_B = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
correntes_latas_B[:, :nr_col_ext] = Ib1_paral
correntes_latas_B[:, nr_col_ext:] = Ib2_paral
df_B = pd.DataFrame(np.abs(correntes_latas_B))

correntes_latas_C = np.ones((nr_lin_ext, 2*nr_col_ext), dtype=complex)
correntes_latas_C[:, :nr_col_ext] = Ic1_paral
correntes_latas_C[:, nr_col_ext:] = Ic2_paral
df_C = pd.DataFrame(np.abs(correntes_latas_C))

with pd.ExcelWriter('correntes_nas_latas.xlsx', engine='openpyxl') as writer:
    df_A.to_excel(writer, sheet_name='Ia', index=False, header=False)
    df_B.to_excel(writer, sheet_name='Ib', index=False, header=False)
    df_C.to_excel(writer, sheet_name='Ic', index=False, header=False)


destaca_maiores_que_nominal(planilha='Ia', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)
destaca_maiores_que_nominal(planilha='Ib', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)
destaca_maiores_que_nominal(planilha='Ic', aquivo='correntes_nas_latas.xlsx', valor_nominal=72.16)


