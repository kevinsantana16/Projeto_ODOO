# 🧾 Projeto Odoo: Automação de Invoices e Integração SAP

## 📋 Descrição
Sistema de automação corporativa desenvolvido em **Python** para a criação de invoices, processamento de relatórios de vendas e sincronização de Notas Fiscais entre os ecossistemas **SAP** e **Odoo ERP**.

## 🛠️ Stack Tecnológico
* **Linguagem:** Python 3.x
* **Sistemas:** SAP, Odoo ERP
* **Arquitetura:** Scripts modulares para rotinas de ETL (Extração, Transformação e Carga)

## 📁 Estrutura de Módulos

| Arquivo | Responsabilidade |

| `Sap_NF_Extract_Oddo.py` | Captura Notas Fiscais geradas no SAP e formata os dados para o padrão de leitura do Odoo. |
| `Oddo_Extract_Sales.py` | Extrai relatórios e registros de vendas diretas a partir da base de dados do Odoo. |
| `Oddo_Extract_NF.py` | Monitora e extrai dados de Notas Fiscais já processadas ou com status pendente no Odoo. |
| `Oddo_Armazenagem.py` | Consolida a criação das invoices e efetua o armazenamento final no sistema destino. |

## ⚙️ Pré-requisitos
* Python 3.x instalado.
* Acessos de rede e permissões ativas para os ambientes SAP e Odoo.

