### Fase A
matriz_A1 = matriz_A[:, :nr_col_ext*nr_col_int]
matriz_A2 = matriz_A[:, nr_col_ext*nr_col_int:]
super_matriz_A1, eq_paral_internos_A1, eq_serie_internos_A1 = matrizes_internas(matriz_FR=matriz_A1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_A2, eq_paral_internos_A2, eq_serie_internos_A2 = matrizes_internas(matriz_FR=matriz_A2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_A1 = (np.sum(eq_serie_internos_A1, axis=1)).reshape(-1, 1)
eq_serie_externos_A1 = 1 / np.sum(1 / eq_paral_externos_A1)
eq_paral_externos_A2 = (np.sum(eq_serie_internos_A2, axis=1)).reshape(-1, 1)
eq_serie_externos_A2 = 1 / np.sum(1 / eq_paral_externos_A2)
eq_paral_ramos_A = eq_serie_externos_A1 + eq_serie_externos_A2


### Fase B
matriz_B1 = matriz_B[:, :nr_col_ext*nr_col_int]
matriz_B2 = matriz_B[:, nr_col_ext*nr_col_int:]
super_matriz_B1, eq_paral_internos_B1, eq_serie_internos_B1 = matrizes_internas(matriz_FR=matriz_B1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_B2, eq_paral_internos_B2, eq_serie_internos_B2 = matrizes_internas(matriz_FR=matriz_B2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_B1 = (np.sum(eq_serie_internos_B1, axis=1)).reshape(-1, 1)
eq_serie_externos_B1 = 1 / np.sum(1 / eq_paral_externos_B1)
eq_paral_externos_B2 = (np.sum(eq_serie_internos_B2, axis=1)).reshape(-1, 1)
eq_serie_externos_B2 = 1 / np.sum(1 / eq_paral_externos_B2)
eq_paral_ramos_B = eq_serie_externos_B1 + eq_serie_externos_B2

### Fase C
matriz_C1 = matriz_C[:, :nr_col_ext*nr_col_int]
matriz_C2 = matriz_C[:, nr_col_ext*nr_col_int:]
super_matriz_C1, eq_paral_internos_C1, eq_serie_internos_C1 = matrizes_internas(matriz_FR=matriz_C1, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=1)
super_matriz_C2, eq_paral_internos_C2, eq_serie_internos_C2 = matrizes_internas(matriz_FR=matriz_C2, nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,  nr_col_int=nr_lin_int, ramo=2)
eq_paral_externos_C1 = (np.sum(eq_serie_internos_C1, axis=1)).reshape(-1, 1)
eq_serie_externos_C1 = 1 / np.sum(1 / eq_paral_externos_C1)
eq_paral_externos_C2 = (np.sum(eq_serie_internos_C2, axis=1)).reshape(-1, 1)
eq_serie_externos_C2 = 1 / np.sum(1 / eq_paral_externos_C2)
eq_paral_ramos_C = eq_serie_externos_C1 + eq_serie_externos_C2


#############################################################################################################################################

matriz_impedancia_malha = np.array([[Za+Zb, -Zb], [-Zb, Zb+Zc]])
matriz_fontes_malha = np.array([[Vff], [Vff*(a**2)]])
matriz_correntes_alphabeta = np.linalg.inv(matriz_impedancia_malha) @ matriz_fontes_malha
I_alpha = matriz_correntes_alphabeta[0, 0]
I_beta = matriz_correntes_alphabeta[1, 0]
matriz_correntes_fase = np.array([[I_alpha], [I_beta-I_alpha], [-I_beta]])
matriz_tensoes_Vabco = matriz_impedancia_sistema @ matriz_correntes_fase
tensao_deslocamento_netro = np.sum(matriz_tensoes_Vabco) / 3