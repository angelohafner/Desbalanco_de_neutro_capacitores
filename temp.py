import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Color, Alignment
from funcoes_desbalanco_neutro import *

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
des_int = 0.0


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

Iao = matriz_correntes_fase[0]
Vao = matriz_tensoes_Vabco[0]

Za1 = eq_serie_externos_A1
Ia1 = Vao / Za1

Z_par_ext = 1/ (1j*omega * eq_paral_externos_A1)

Va1_par_ext = Ia1 * Z_par_ext


print(eq_paral_internos_A1.shape)
print(Va1_par_ext.shape)














