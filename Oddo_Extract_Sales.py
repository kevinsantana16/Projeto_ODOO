from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from Oddo_Extract_NF import Extract_NF
import pandas as pd
import openpyxl


def Extract_sales(browser, linhas_total, user_id, data_texto, tabela):

# Fazer filtro
    print("Iniciando extração de Sales Order...")
    k = 1
    contador = 1
    while True:
        sleep(5)
        browser.refresh()
        sleep(3)
        
        sales = []
        status_extra = []
        for k in range(1,linhas_total + 1):
            # Ver se o status está finalizado
            if tabela.iloc[k-1]['Status - Carga Massiva'] == "Cancelado" or tabela.iloc[k-1]['Status - Carga Massiva'] == "Rejeitado":
                sales.append('Not found')
                status_text = "Rejeitado"
                status_extra.append(status_text)
            else:
                status = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, f'/html/body/div[1]/div/div[2]/div/div[1]/table/tbody/tr[{k}]/td[6]')))
                status_text = status.text.strip()
                status_extra.append(status_text)
                # Pegar planta e Sales Order
                Planta = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, f'/html/body/div[1]/div/div[2]/div/div[1]/table/tbody/tr[{k}]/td[4]')))
                Planta_text = Planta.text.strip()
                Sales_order = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, f'/html/body/div[1]/div/div[2]/div/div[1]/table/tbody/tr[{k}]/td[8]')))
                Sales_order_text = Sales_order.text.strip()
                data_criacao = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, f'/html/body/div[1]/div/div[2]/div/div[1]/table/tbody/tr[{k}]/td[12]')))
                data_criacao_text = data_criacao.text.strip()
                criado_data = data_criacao_text.split()[0]
                sleep(2)
            # Armazenar e transformar em tabela

                if status_text == "Finalizado" and k == 1:
                    
                    sales.append(Sales_order_text)
                    sleep(1)
                    contador += 1
                elif status_text == "Finalizado" and k > 1:
                    contador += 1
                    
                    sales.append(Sales_order_text)
                else:
                    contador += 1
                    
                    sales.append(Sales_order_text)
        if all(status_text in ("Finalizado","Rejeitado","Cancelado") for status_text in status_extra) == True:
            break
        else:
            print("Nova tentativa")
            sleep(50)
            pass


    print("Sales extraídas!")
    sleep(25)
    
    
    
    
    Extract_NF(user_id, tabela)

def definir_col(directory, excel_files, sheet, sheet2, cost_center):
        
        df = pd.read_excel(f'{directory}\\{excel_files}', engine = "openpyxl", sheet_name="Presença de carga")
        df = df.ffill()
        df = df.drop_duplicates() #garantir que não tenha linhas iguais para faturar 2 vezes


        linhas = len(df.index)

        value_UOM = 'KG'
        value_operacao = 'ZM32'
        value_destinacao = "Revenda"

        indice = 0
        indice_2 = 0
        
        for k in range(0, linhas,1):
            
            planta = df.iloc[k]["Planta"] #Planta

            GMID = df.iloc[k]["GMID"] #GMID

            filtro = (cost_center['Planta'] == planta)
            filtered_df = cost_center[filtro]["Centro de custo"]
            filtered_df_company = cost_center[filtro]["Company Code"]

            filtro_customer = cost_center['Planta'] == planta
            filtered_customer = cost_center[filtro_customer]["Customer"]
            
            try:
                Po = df.iloc[k]["P.O"] #P.O
            except:
                Po = df.iloc[k]["PO:"] #PO
           
            try:
                valor_NF = float(df.iloc[k]["VALOR NF"])
            except:
                valor_NF = float(df.iloc[k]["Valor da NF"])

            try:
                customer = filtered_customer.iloc[0] if not filtered_customer.empty else ""
            except:
                customer = ""

            try:
                cc = filtered_df.iloc[0]
                company_code = int(filtered_df_company.iloc[0])
            except:
                cc = ""
                company_code = ""
            
            try:
                peso_bruto = df.iloc[k]["VOL. TOTAL"]
            except:
                peso_bruto = df.iloc[k]["Peso Bruto"]

          

            try:
                peso_liquido_unit = df.iloc[k]["PESO LIQUIDO"]
            except:
               peso_liquido_unit = df.iloc[k]["Peso Líquido (Quantidade Real) "]

            
            peso_liquido_total = peso_bruto
            preco_unit = f"{valor_NF / peso_liquido_unit:.6f}"

            
                    
            

            col_po = "P.O" if "P.O" in df.columns else "PO:" 
            col_nf = "NOTA FISCAL" if "NOTA FISCAL" in df.columns else "NF de origem:"      
            nf_dic = (df.groupby(col_po)[col_nf]
                        .apply(list)       # agrega as NFs por PO
                        .to_dict())

            # Uso:
            NF = list(dict.fromkeys(nf_dic.get(Po, [])))
            Observacoes = f"PO {Po} NF {', '.join(map(str, NF))}"

       

            
            if cc != "":
                indice += 1
            else:
                indice_2 += 1
                cel_operacao_2 = f'E{indice_2 + 1}'
                cel_UOMtext_2 = f'N{indice_2 + 1}'
                cel_Obs_2 = f'K{indice_2 + 1}'
                cel_indice_2 = f'A{indice_2 + 1}'
                cel_customer2 = f'C{indice_2 + 1}'
                cel_linhaitem_2 = f'L{indice_2 + 1}'
                cel_planta_2 = f'D{indice_2 + 1}'
                cel_destinacao_2 = f'F{indice_2 + 1}'
                cel_gmid_2 = f'M{indice_2 + 1}'
                cel_liquido_2 = f'Q{indice_2 + 1}'
                cel_bruto_2 = f'R{indice_2 + 1}'
                cel_quantidade_2 = f'O{indice_2 + 1}'
                cel_unitario_2 = f'P{indice_2 + 1}'
                cel_costcenter_2 = f'G{indice_2 + 1}'
                cel_companycode_2 = f'B{indice_2 + 1}'

                sheet2[cel_indice_2] = indice_2
                sheet2[cel_linhaitem_2] = indice_2
                sheet2[cel_operacao_2] = value_operacao
                sheet2[cel_UOMtext_2] = value_UOM
                sheet2[cel_destinacao_2] = value_destinacao
                sheet2[cel_customer2] = customer

                sheet2[cel_liquido_2] = peso_liquido_total
                sheet2[cel_bruto_2] = peso_bruto
                sheet2[cel_quantidade_2] = peso_bruto
                sheet2[cel_unitario_2] = preco_unit
                sheet2[cel_gmid_2] = GMID
                sheet2[cel_costcenter_2] = cc
                sheet2[cel_companycode_2] = company_code

                sheet2[cel_planta_2] = planta
                sheet2[cel_Obs_2] = Observacoes
                
                continue

            cel_operacao = f'E{indice + 1}'
            cel_UOMtext = f'N{indice + 1}'
            cel_Obs = f'K{indice + 1}'
            cel_indice = f'A{indice + 1}'
            cel_customer = f'C{indice + 1}'
            cel_linhaitem = f'L{indice + 1}'
            cel_planta = f'D{indice + 1}'
            cel_destinacao = f'F{indice + 1}'
            cel_gmid = f'M{indice + 1}'
            cel_liquido = f'Q{indice + 1}'
            cel_bruto = f'R{indice + 1}'
            cel_quantidade = f'O{indice + 1}'
            cel_unitario = f'P{indice + 1}'
            cel_costcenter = f'G{indice + 1}'
            cel_companycode = f'B{indice + 1}'


            sheet[cel_indice] = indice
            sheet[cel_linhaitem] = indice
            sheet[cel_operacao] = value_operacao
            sheet[cel_UOMtext] = value_UOM
            sheet[cel_destinacao] = value_destinacao
            sheet[cel_customer] = customer

            sheet[cel_liquido] = peso_liquido_total
            sheet[cel_bruto] = peso_bruto
            sheet[cel_quantidade] = peso_bruto
            sheet[cel_unitario] = preco_unit
            sheet[cel_gmid] = GMID
            sheet[cel_costcenter] = cc
            sheet[cel_companycode] = company_code

            sheet[cel_planta] = planta
            sheet[cel_Obs] = Observacoes

            
def cedula_branco(linhas, sheet3, i, excel_files, workbook3,directory):
    
    for k in range(0, linhas, 1):
        celulas = [sheet3[f'B{k + 2}'], sheet3[f'H{k + 2}'], sheet3[f'I{k + 2}'],sheet3[f'J{k + 2}'],
                 sheet3[f'L{k + 2}'],sheet3[f'M{k + 2}'],sheet3[f'N{k + 2}'], sheet3[f'S{k + 2}'],sheet3[f'T{k + 2}'],sheet3[f'V{k + 2}']]
        # Verificar se está vazia
        for celula in celulas:
            if celula.value is None or str(celula.value).strip() == '':
                celula.value = 'Valor não encontrado'
                workbook3.save(f'{directory}\\{excel_files[i]}')
