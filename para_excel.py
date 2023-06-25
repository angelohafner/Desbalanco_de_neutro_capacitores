#%% Fase A
df_Ca_eq_paral_ramos_A = pd.DataFrame([eq_paral_ramos_A])
df_Ca_eq_ext = pd.DataFrame([[eq_serie_externos_A1], [eq_serie_externos_A2]])
df_Ca_pa_ext = pd.DataFrame([[eq_paral_externos_A1], [eq_paral_externos_A2]])
df_Ca_eq_int = pd.DataFrame(np.append(eq_serie_internos_A1, eq_serie_internos_A2, axis=1))

with pd.ExcelWriter('capacitancias_nas_latas.xlsx', engine='openpyxl') as writer:
    df_Ca_eq_ext.to_excel(writer, sheet_name='df_Ca_eq_ext', index=False, header=False)
    df_Ca_pa_ext.to_excel(writer, sheet_name='df_Ca_pa_ext', index=False, header=False)
    df_Ca_eq_int.to_excel(writer, sheet_name='df_Ca_eq_int', index=False, header=False)