import pdfplumber
import pandas as pd
import sys
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
import openpyxl
import os
'''
A função pdf_to_excel lê o ficheiro pdf e converte-o para um ficheiro excel
As páginas de pdf que são convertidas para excel são aquelas que contém tabelas com medidas
'''
def pdf_to_excel(nome_pdf,excel_name):
    with pdfplumber.open(nome_pdf) as pdf:

        with pd.ExcelWriter(excel_name, engine='xlsxwriter') as writer:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines"
                })

                # Procurar a partir da tabela que contém "POM"
                collected_tables = []

                for table in tables:
                    df = pd.DataFrame(table).astype(str)
                    collected_tables.append(df)

                # Só escreve no Excel se encontrou o "POM"
                if collected_tables:
                    # Juntar os DataFrames
                    final_df = pd.DataFrame()
                    for df in collected_tables:
                        final_df = pd.concat([final_df, df, pd.DataFrame([[""] * len(df.columns)])], ignore_index=True)

                    final_df.to_excel(writer, sheet_name=f'Page_{i+1}', index=False, header=False)

        return 
    
# ...existing code...
def trim_excel_before_marker(excel_path,excel_saida):

    """
    Abre o ficheiro excel, procura na primeira coluna da primeira sheet a primeira linha que contém `marker`
    (comparação case-insensitive) e elimina todas as linhas anteriores. Sobrescreve o
    ficheiro por defeito, ou grava em out_path se fornecido.
    """
    sheets = pd.read_excel(excel_path, sheet_name=None, header=None, dtype=str)
    #devolve algo do género {'Sheet1': df1, 'Sheet2': df2, …}

    #escolher a primeira sheet do excel
    first_sheet_name = list(sheets.keys())[0]
    #garante que todos os valores são string e sem NaN (substitui NaN pela string "")
    df = sheets[first_sheet_name].fillna('').astype(str)

    # Usa apenas a primeira coluna para procurar o marcador
    #remover espaço em branco no inicio e fim - strip - converter tudo a minusculas - lower
    first_col = df.iloc[:, 0].str.strip().str.lower()
    #criar mascara para encontrar onde está a linha que diz "Malhas e Tecidos"
    marker="Malhas e Tecidos"
    mask = first_col.eq(marker.lower())

    if not mask.any():
        #para de correr o código caso não econtre "Malhas e Tecidos"
        raise ValueError(f"'{marker}' não encontrado na primeira folha do excel.")
    
    """
    
    preparar excel
    
    """
    linhas_excel=[]
    header_colunas = ["","","consumo", "preço original", "valor margem","distrib. margem corte + (acessorios*perc)","desconto","preço final"]
    #linhas_excel.append(header_colunas)


    """
    Obter valor da margem a aplicar
    """
    #procurar na segunda sheet do excel
    second_sheet_name = list(sheets.keys())[1]
    #garante que todos os valores são string e sem NaN (substitui NaN pela string "")
    df2 = sheets[second_sheet_name].fillna('').astype(str)

    #selecionar a primeira coluna
    first_col2 = df2.iloc[:, 0].str.strip().str.lower()
    
    # margem_aplicar
    marker_margem_aplicar = ["margem"]
    for marker in marker_margem_aplicar:
        mask_margem_aplicar = first_col2.eq(marker.lower())
        if not mask_margem_aplicar.any():
            raise ValueError(f"'{marker}' não encontrado na segunda folha do excel.")

        margem_aplicar_idxs = mask_margem_aplicar[mask_margem_aplicar].index.tolist()
        margem_aplicar_idx = margem_aplicar_idxs[-1]  # Pega no último índice

    perc = pd.to_numeric(df2.iloc[margem_aplicar_idx, 2], errors='coerce')

    percent_value = perc/100

    """
    Obter o consumo e preço por malha e definir main fabric, 2nd fabric, etc
    """
    #devolve o indice da primeira ocorrência de "Malhas e Tecidos"
    first_idx = mask.idxmax()
    #devolve um df com as linhas a partir de "Malhas e Tecidos" e redefine a numeração dos indices
    trimmed = df.iloc[first_idx:].reset_index(drop=True)

    #filtrar pelas linhas cuja coluna "C" tem "SORTIMENTO"
    col_c = trimmed.iloc[:, 2].str.strip().str.lower()
    mask_sortimento = col_c.str.contains("sortimento", na=False)
    filtered_df = trimmed[mask_sortimento].reset_index(drop=True)

    # Na coluna A (índice 0), contar cells não vazias
    col_a = filtered_df.iloc[:, 0].str.strip()  
    malha_indices = col_a[col_a != ""].index.tolist()


    #indice da ultima linha do excel 
    ultima_linha = len(filtered_df) - 1

    #guardar os consumos e preços por malha antes e apos aplicar a margem
    consumos_malha=[]
    malha_precos_original=[]
    malhas_precos_margem=[]
    malhas_precos_final=[]
    nome_malhas=[]

    for i in range(0,len(malha_indices)):
        if i < len(malha_indices)-1:
            consumo=pd.to_numeric(filtered_df.iloc[malha_indices[i], 4], errors='coerce')
            consumos_malha.append(consumo)
            nome_malhas.append(filtered_df.iloc[malha_indices[i], 0])
            soma1=pd.to_numeric(filtered_df.iloc[malha_indices[i]:malha_indices[i+1], 6], errors='coerce').sum()
            malha_precos_original.append(soma1)
            malhas_precos_margem.append(soma1*percent_value)
            soma1 = soma1*(1+percent_value)
            malhas_precos_final.append(soma1)
        else:
            consumo=pd.to_numeric(filtered_df.iloc[malha_indices[i], 4], errors='coerce')
            consumos_malha.append(consumo)
            nome_malhas.append(filtered_df.iloc[malha_indices[i], 0])
            soma1=pd.to_numeric(filtered_df.iloc[malha_indices[i]:ultima_linha+1, 6], errors='coerce').sum()
            malha_precos_original.append(soma1)
            malhas_precos_margem.append(soma1*percent_value)
            soma1 = soma1*(1+percent_value)
            malhas_precos_final.append(soma1)

    # Criar lista de tuplos com (indice, consumo, preco)
    dados_malhas = list(zip(malha_indices,nome_malhas, consumos_malha, malha_precos_original, malhas_precos_margem,malhas_precos_final))

    # Ordenar malha_indices por ordem crescente dos valores na coluna E (índice 4)
    col_e = pd.to_numeric(filtered_df.iloc[:, 4], errors='coerce')
    # Ordenar pelo valor na coluna E correspondente a cada indice (x[0] no codigo corresponde ao indice da malha que está em cada tuplo)
    dados_malhas_sorted = sorted(dados_malhas, key=lambda x: col_e.iloc[x[0]], reverse=True)


    # Contar malhas diferentes
    num_malhas = len(dados_malhas_sorted)
    
    fabrics =["Main fabric","2nd fabric","3rd fabric","4th fabric","5th fabric","6th fabric"]
    # Criar lista de nomes
    fabric_names = []
    if num_malhas >= 1:
        fabric_names.append(fabrics[0])
    for i in range(1, num_malhas):
        fabric_names.append(fabrics[i])

    for i in range(0, num_malhas):
        linha_malha = []
        linha_malha.append(dados_malhas_sorted[i][1])  # nome da malha
        linha_malha.append(fabric_names[i])  # main fabric, 2nd fabric, etc
        linha_malha.append(dados_malhas_sorted[i][2])  # consumo
        linha_malha.append(dados_malhas_sorted[i][3])  # preço antes de aplicar a margem
        linha_malha.append(dados_malhas_sorted[i][4])  # valor da margem
        linha_malha.append("")  # distribuição margem corte + (acessórios*perc)
        linha_malha.append("")  # desconto
        linha_malha.append(dados_malhas_sorted[i][5])  # preço após aplicar a margem
        linhas_excel.append(linha_malha)

    
    """
    FIM DE - Obter o consumo e preço por malha e definir main fabric, 2nd fabric, etc

    listas a usar: fabric_names e dados_malhas_sorted (tuplo com indice, consumo, preco)
    """


    """
    CALCULAR RESTANTES COMPONENTES DO PREÇO: ARTWORKS, WASHING, CMT(CORTE,CONFEÇÃO,EMBALAMENTO), OTHER COSTS, ACESSÓRIOS
    """

    # CMT (MANIFACTURING COST)
    markers_cmt = ["corte", "confecção", "embalamento"]
    indices_cmt = []
    for marker in markers_cmt:
        mask_cmt = first_col2.eq(marker.lower())
        if not mask_cmt.any():
            raise ValueError(f"'{marker}' não encontrado na segunda folha do excel.")
        cmt_idx = mask_cmt.idxmax()
        indices_cmt.append(cmt_idx)

    cmt_cost = 0
    for idx in indices_cmt:
        cost = pd.to_numeric(df2.iloc[idx, 3], errors='coerce')
        if not pd.isna(cost):
            cmt_cost += cost
    cmt_margem = cmt_cost * percent_value
    cmt_cost_final = cmt_cost * (1+percent_value)

    linhas_excel.append(["","CMT", "", cmt_cost, cmt_margem, "", "", cmt_cost_final])

    # acessorios
    marker_acessorios = "Acessorios"
    mask_acessorios = first_col2.eq(marker_acessorios.lower())
    if not mask_acessorios.any():
        acessorios_cost=0
        perc_acessorios = 0
    else:
        acessorios_idx = mask_acessorios.idxmax()
        acessorios_cost = pd.to_numeric(df2.iloc[acessorios_idx, 3], errors='coerce')
        perc_acessorios = acessorios_cost * percent_value
    
    linhas_excel.append(["","Accessories", "", acessorios_cost, "", "", "", acessorios_cost])

    #margem corte
    marker_margem_corte = "Margem Corte"
    mask_margem_corte = first_col2.eq(marker_margem_corte.lower())
    if not mask_margem_corte.any(): 
        margem_corte_cost=0
    else:
        margem_corte_idx = mask_margem_corte.idxmax()
        margem_corte_cost = pd.to_numeric(df2.iloc[margem_corte_idx, 3], errors='coerce')
        margem_corte_cost = margem_corte_cost * (1+percent_value)

    if mask_acessorios.any() and mask_margem_corte.any():
        add_cost_div= margem_corte_cost + perc_acessorios
    elif mask_acessorios.any() and not mask_margem_corte.any():
        add_cost_div = perc_acessorios
    elif not mask_acessorios.any() and mask_margem_corte.any():
        add_cost_div = margem_corte_cost


    div_perc = 0

    # artworks

    marker_descontos = "Desconto"
    mask_descontos = first_col2.eq(marker_descontos.lower())
    if not mask_descontos.any():
        descontos_cost=0
    else:
        descontos_idx = mask_descontos.idxmax()
        descontos_cost = pd.to_numeric(df2.iloc[descontos_idx, 3], errors='coerce')
    
    marker_artworks = "Bord./Est. (Animações)"
    mask_artworks = first_col2.eq(marker_artworks.lower())
    if not mask_artworks.any():
        artworks_cost=0
        artworks_margem = 0
    else:
        artworks_idx = mask_artworks.idxmax()
        artworks_cost = pd.to_numeric(df2.iloc[artworks_idx, 3], errors='coerce')
        artworks_margem = artworks_cost * percent_value
        div_perc += 1


    # washing
    marker_washing = "Acabamentos a Peça"
    mask_washing = first_col2.eq(marker_washing.lower())
    if not mask_washing.any():  
        washing_cost=0
        washing_margem = 0
    else:
        washing_idx = mask_washing.idxmax()
        washing_cost = pd.to_numeric(df2.iloc[washing_idx, 3], errors='coerce')
        washing_margem = washing_cost * percent_value
        div_perc += 1
    
    if mask_acessorios.any():
        if mask_artworks.any() and not mask_washing.any():
            washing_cost_final = 0
            artworks_cost_final = (artworks_cost * (1+percent_value)) + (add_cost_div/div_perc) + descontos_cost
        elif not mask_artworks.any() and mask_washing.any():
            washing_cost_final = (washing_cost * (1+percent_value)) + (add_cost_div/div_perc) 
            artworks_cost_final = 0
        elif mask_artworks.any() and mask_washing.any():
            artworks_cost_final = (artworks_cost * (1+percent_value)) + (add_cost_div/div_perc) + descontos_cost
            washing_cost_final = (washing_cost * (1+percent_value)) + (add_cost_div/div_perc)  


    linhas_excel.append(["","Artworks", "", artworks_cost, artworks_margem, add_cost_div/div_perc, descontos_cost, artworks_cost_final])
    linhas_excel.append(["","Washing", "", washing_cost, washing_margem, add_cost_div/div_perc, "", washing_cost_final])

    # other costs
    marker_other_costs = ['Gastos Gerais', 'Transporte']
    indices_other_costs = []
    for marker in marker_other_costs:
        mask_other_costs = first_col2.eq(marker.lower())
        if not mask_other_costs.any():
            raise ValueError(f"'{marker}' não encontrado na segunda folha do excel.")
        other_costs_idx = mask_other_costs.idxmax()
        indices_other_costs.append(other_costs_idx)
    
    other_costs = 0
    for idx in indices_other_costs:
        cost = pd.to_numeric(df2.iloc[idx, 3], errors='coerce')
        if not pd.isna(cost):
            other_costs += cost
    
    other_cost_margem = other_costs * percent_value
    other_costs_final = other_costs * (1+percent_value)

    linhas_excel.append(["","Other", "", other_costs, other_cost_margem, "", "", other_costs_final]) 

    """
    FIM - CALCULAR RESTANTES COMPONENTES DO PREÇO: ARTWORKS, WASHING, CMT(CORTE,CONFEÇÃO,EMBALAMENTO), OTHER COSTS, ACESSÓRIOS
    
    variaveis a usar: cmt_cost, acessorios_cost, artworks_cost, washing_cost
    """
    
    """
    CRIAR TABELA COM TODAS AS INFORMAÇÕES
    """

    df_final = pd.DataFrame(linhas_excel,columns = header_colunas)

    
    total_cost_per_garment = df_final['preço final'].sum()
    nova_linha = pd.DataFrame([["", "Total per garment", "", "", "", "", "", total_cost_per_garment]], columns=header_colunas)
    df_final = pd.concat([df_final, nova_linha], ignore_index=True)
          # Escrever no Excel
    with pd.ExcelWriter(excel_saida, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='Sheet1', index=False,header=True)

    
    # Formatar (opcional)
    wb = load_workbook(excel_saida)
    #com workbook os indices começam em 1
    ws = wb.active
    
    for sheet in wb.worksheets:
        # Alterar a altura de todas as linhas
        for row in sheet.iter_rows():
            row[0].parent.row_dimensions[row[0].row].height = 18.5
        
        # Alterar a altura da primeira linha
        sheet.row_dimensions[1].height = 55 # Linha 1 com altura 25

        # Definir a largura da primeira coluna (A)
        sheet.column_dimensions['A'].width = 17
        # Definir a largura da primeira coluna (B)
        sheet.column_dimensions['B'].width = 26



        start_col_idx = openpyxl.utils.column_index_from_string('C')

        # Iterar por todas as colunas a partir de 'C' e ajustar a largura
        for col_idx in range(start_col_idx, sheet.max_column + 1):
            column_letter = openpyxl.utils.get_column_letter(col_idx)
            sheet.column_dimensions[column_letter].width = 20

        for row in sheet.iter_rows():
            for cell in row:
                cell.font = Font(size=14)  # Tamanho da letra 14
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                #if cell.column >= start_col_idx:  # Apenas células a partir da coluna 'D'
                #    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Definir a espessura da borda
        border_style = Border(
            left=Side(style='thin', color='000000'),  # Bordas à esquerda
            right=Side(style='thin', color='000000'),  # Bordas à direita
            top=Side(style='thin', color='000000'),  # Bordas em cima
            bottom=Side(style='thin', color='000000')  # Bordas embaixo
        )

        # Iterar sobre todas as linhas e colunas da planilha
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = border_style  # Definir a borda da célula


    wb.save(excel_saida)

    """
    FIM- CRIAR TABELA COM TODAS AS INFORMAÇÕES
    """
    


    return 

