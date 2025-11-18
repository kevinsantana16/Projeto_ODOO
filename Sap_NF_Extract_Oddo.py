import pandas as pd
import win32com.client
import subprocess

from time import sleep
import time

import pandas as pd

class SapGui(object):
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        self.SapGuiAuto = None
        self.connection = None
        self.session = None
        try:
            # Tenta obter o objeto SAPGUI se já estiver aberto
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            print("SAP já está aberto.")

            self.application = self.SapGuiAuto.GetScriptingEngine
            if self.application.Connections.Count > 0:
                self.connection = self.application.Children(0)
                if self.connection.Sessions.Count > 0:
                    self.session = self.connection.Children(0)
                    print("Sessão SAP ativa encontrada.")
                else:
                    print("Nenhuma sessão SAP ativa encontrada.")
            else:
                print("Nenhuma conexão SAP ativa encontrada.")
                application = self.SapGuiAuto.GetScriptingEngine
                self.connection = application.OpenConnection("NEA Prod - ERP Central Component (ECC) (SSO) (001)", True)
                time.sleep(10)
                self.session = self.connection.Children(0)
                self.session.findById("wnd[0]").maximize()

        except:
            # Se não estiver aberto, inicia o SAP Logon
            print("SAP não está aberto. Iniciando SAP Logon...")
            subprocess.Popen(self.path)
            time.sleep(5)
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")

            application = self.SapGuiAuto.GetScriptingEngine
            self.connection = application.OpenConnection("NEA Prod - ERP Central Component (ECC) (SSO) (001)", True)
            time.sleep(10)
            self.session = self.connection.Children(0)
            self.session.findById("wnd[0]").maximize()

        if not self.SapGuiAuto:
            raise Exception("Não foi possível obter o objeto SAPGUI.")

      
    def Execute_SAP(self, tabela, user_id):
        application = self.SapGuiAuto.GetScriptingEngine
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "VA03"
        self.session.findById("wnd[0]").sendVKey(0)
        rows_tbl = len(tabela.index)
        nf_text = []
        for m in range(0, rows_tbl):
    
            Sales = str(tabela.loc[m, 'Sales Order'])
            if Sales != "Not found":
         
                self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = f"{Sales}"
                self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 9
                self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
                self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem("          3", "&Hierarchy")
                self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem(
                    "          3", "&Hierarchy")
                self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                self.session.findById("wnd[0]/usr/txtBKPF-XBLNR").setFocus()
                NF = self.session.findById("wnd[0]/usr/txtBKPF-XBLNR").text
                nf_text.append(NF)
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem("          2", "&Hierarchy")
                self.session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem(
                    "          2", "&Hierarchy")
                self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
                self.session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").currentCellRow = 1
                self.session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "1"
                self.session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").doubleClickCurrentCell()
                self.session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
                self.session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_VIEW_ATTA")
                try:
                    self.session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
                    self.session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
                    self.session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton(
                        "%ATTA_EXPORT")
                    self.session.findById(
                        "wnd[2]/usr/ctxtDY_PATH").text = "" # Colocar o caminho desejado para salvar o arquivo
                    self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = f"{NF}.PDF"
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 1
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    sleep(4)

                    self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                    sleep(4)
                    try:
       
                        self.session.findById("wnd[1]").close()
                        sleep(4)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        sleep(2)
                        self.session.findById("wnd[1]").close()
                        sleep(2)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                    except:
  
                        sleep(3)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        sleep(2)
                        self.session.findById("wnd[1]").close()
                        sleep(2)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        pass
                except:
                    try:
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        sleep(2)
                        self.session.findById("wnd[1]").close()
                        sleep(2)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
          
                    except:
                       
                        sleep(3)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        sleep(2)
                        self.session.findById("wnd[1]").close()
                        sleep(2)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
             
                        pass
            else:
                nf_text.append("Not found")

        return nf_text

