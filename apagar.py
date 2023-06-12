wb.save(filename)
wb = load_workbook(filename)
ws = wb[f'fase_{fase}']
# Definir o padrão de preenchimento para preto
black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

# Percorrer todas as células na planilha
for row in ws.iter_rows():
    for cell in row:
        # Se a célula contém o valor 0, altere a cor de fundo para preto
        if cell.value == 0:
            cell.fill = black_fill

# Salvar a planilha
wb.save(filename)