#--------------------------------------------------------------Definicióndelibrerias---------------------------------------------------#

import os, os.path
import win32com.client
from os import walk
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import sys
import os
from getpass import getuser
import zipfile
import shutil
import openpyxl
from datetime import date
from datetime import datetime
import win32com.client as win32
from os import remove
import threading
import pythoncom
from datetime import datetime
import requests
import win32com.client
import sys
import subprocess
import time
import pythoncom
from openpyxl import load_workbook


def ingresarsap(u,c):
    try:

        pythoncom.CoInitialize()
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(11)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection("QAS", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = u
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = c
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.CreateSession()
        time.sleep(1)
        session.CreateSession()
        time.sleep(1)
        session.CreateSession()
        time.sleep(1)
        session.CreateSession()
        time.sleep(1)

        
    except:
        print(sys.exc_info()[0] + "hola")

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None



def activarmacro(numero,user):
    path='C:/Users/' + user + '/Desktop/txt_osde/Bot_TXT_v1_' + str(numero) + '.xlsm'
    if os.path.exists(path):
        pythoncom.CoInitialize()
        Excel_macro = win32com.client.DispatchEx("Excel.Application") # DispatchEx is required in the newest versions of Python.
        Excel_path = os.path.expanduser(path)
        workbook = Excel_macro.Workbooks.Open(Filename = Excel_path, ReadOnly =1)
        Excel_macro.Application.Run('Bot_TXT_v1_' + str(numero)+'.xlsm' + "!" + "z_integrador.integrador") # update Module1 with your module, Macro1 with your macro
        workbook.Save()
        Excel_macro.Application.Quit()  
        del Excel_macro
        for i in range(1,50):
            print(numero)
    pass

def meteteensap(ped,numerodesesion):
        pythoncom.CoInitialize()
        time.sleep(2)
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.Children(0)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(numerodesesion)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        
        time.sleep(2)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NZSD_TOMA"
        time.sleep(1)
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_ERDAT-LOW").text = "01.01.2020"
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").text = ped
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").setFocus()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").caretPosition = 7
        session.findById("wnd[0]").sendVKey (0)
        time.sleep(1)
        try:
            session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellColumn = "BSTKD"
        except:
            respuesta="error_en_sap"
            return respuesta
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton ("FN_MODDEL")
        try:
            session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/ctxtZSD_TOMA_CABEC-VTWEG").text = "10"
        except:
            respuesta=session.findById("wnd[0]/sbar").text
            return respuesta
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/ctxtZSD_TOMA_CABEC-VTWEG").setFocus()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/ctxtZSD_TOMA_CABEC-VTWEG").caretPosition = 2
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT").select()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-FECHA_DISP").setFocus()
        fechadeentrega=session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-FECHA_DISP").text
        print(fechadeentrega)
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/txtGS_ENTREGA-FECHA_DISP").caretPosition = 5
        session.findById("wnd[0]").sendVKey (2)
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        try:
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        respuesta=session.findById("wnd[0]/sbar").text
        print(respuesta)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
        session.findById("wnd[0]").sendVKey (0)
        try:
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = str(respuesta[23:30])
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
            session.findById("wnd[0]").sendVKey (0)
   
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").text = str(fechadeentrega)
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").setFocus()
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").caretPosition = 10
        except:
            respuesta="error_en_sap"
            return respuesta   
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        return respuesta


def sen(numerodesesion):
        pythoncom.CoInitialize()

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.Children(0)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(numerodesesion)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        
#         session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
#         session.findById("wnd[0]").sendVKey (0)
#         session.findById("wnd[0]/usr/ctxtGD-TAB").text = "vbak"
#         session.findById("wnd[0]").sendVKey (0)
#         session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = "500"
#         session.findById("wnd[0]/tbar[1]/btn[18]").press()
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,1]").selected = True
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,1]").setFocus
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 149
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 167
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 185
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 174
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 173
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 172
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 171
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 170
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 169
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,3]").selected = True
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,3]").setFocus()
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 28
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 0
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
#         session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
#         session.findById("wnd[1]/tbar[0]/btn[24]").press()
#         session.findById("wnd[1]/tbar[0]/btn[0]").press()
#         session.findById("wnd[1]/tbar[0]/btn[8]").press()
#         session.findById("wnd[0]/tbar[1]/btn[8]").press()
#         tabla=session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getcellvalue(7,"VBELN")
#         print(tabla)
    


# sen(0)
        
        
        
def facturadesap(id_cliente,numerodesesion,factura_sap):
        pythoncom.CoInitialize()

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.Children(0)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(numerodesesion)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NVF31"
        time.sleep(0.1)
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.1)
        session.findById("wnd[0]/usr/ctxtRG_KSCHL-LOW").text = "ZLEG"
        time.sleep(0.1)
        try:
            session.findById("wnd[0]/usr/ctxtPM_VERMO").text = "2"
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").text = factura_sap
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").setFocus()
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").caretPosition = 8
            time.sleep(0.1)
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.1)
            session.findById("wnd[0]/usr/chk[1,3]").selected = True
            time.sleep(0.1)
        except:
            session.findById("wnd[0]/usr/ctxtPM_VERMO").text = "1"
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").text = factura_sap
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").setFocus()
            time.sleep(0.1)
            session.findById("wnd[0]/usr/ctxtRG_VBELN-LOW").caretPosition = 8
            time.sleep(0.1)
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.1)
            session.findById("wnd[0]/usr/chk[1,3]").selected = True
            time.sleep(0.1)
        
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        time.sleep(0.1)
        session.findById("wnd[0]/tbar[1]/btn[18]").press()
        time.sleep(0.1)
        texto=session.findById("wnd[1]/usr/tblSAPLV70ATCPROT/txtPROT-MSGTX[0,1]").text
        time.sleep(0.1)
        posfinal=len(texto)
        time.sleep(0.1)
        calculotexto=posfinal-52
        time.sleep(0.1)
        posinicial=posfinal-calculotexto
        spool=texto[posinicial:posfinal]
        time.sleep(0.1)
        session.findById("wnd[1]/usr/tblSAPLV70ATCPROT/txtPROT-MSGTX[0,1]").caretPosition = 54
        time.sleep(0.1)
        session.findById("wnd[1]").sendVKey(2)
        time.sleep(0.1)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NSP01"
        
        session.findById("wnd[0]").sendVKey(0)
        
        session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQIDEN-LOW").text = spool
        
        session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQIDEN-LOW").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        #session.findById("wnd[1]/usr/btnSEL2").press()
        session.findById("wnd[0]/usr/chk[1,3]").selected = True
        session.findById("wnd[0]/usr/chk[1,3]").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[13]").press()
        session.findById("wnd[0]/usr/txtTSP01_SP0R-RQTITLE").text = factura_sap
        session.findById("wnd[0]/usr/txtTSP01_SP0R-RQTITLE").setFocus()
        session.findById("wnd[0]/usr/txtTSP01_SP0R-RQTITLE").caretPosition = 13
        session.findById("wnd[0]/tbar[1]/btn[13]").press()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NVF31"
        session.findById("wnd[0]").sendVKey(0)

