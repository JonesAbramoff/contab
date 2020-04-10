VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmForprint"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "funcoes uteis para interface com o gerador de relatorios Forprint"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function Forprint_ConvInt(iIntNum As Integer) As String
Attribute Forprint_ConvInt.VB_Description = "converte long p/formato Forprint: o numero como string"
    Forprint_ConvInt = CStr(iIntNum)
End Function

Public Function Forprint_ConvLong(lLongNum As Long) As String
Attribute Forprint_ConvLong.VB_Description = "converte long p/formato Forprint: o numero como string"
    Forprint_ConvLong = CStr(lLongNum)
End Function

Public Function Forprint_ConvDouble(dDoubleNum As Double) As String
Attribute Forprint_ConvDouble.VB_Description = "converte double p/formato Forprint: expresso com ponto decimal, sem separador de milhar"
    Forprint_ConvDouble = Replace(Format(dDoubleNum, "##############.####;-##############.####;0"), ",", ".")
End Function

Public Function Forprint_ConvData(dtData As Date) As String
Attribute Forprint_ConvData.VB_Description = "converte uma data p/o formato forprint 'dd/mm/aaaa'"
    Forprint_ConvData = Format(dtData, "'dd/mm/yyyy'")
End Function

Public Function Forprint_ConvTexto(sTexto As String) As String
Attribute Forprint_ConvTexto.VB_Description = "converte um texto p/o formato Forprint: entre aspas"
    Forprint_ConvTexto = """" & sTexto & """"
End Function

Public Function Forprint_ConvDataVB(ByVal dtData As Date) As Date
    Forprint_ConvDataVB = DateAdd("d", -693594, dtData)
End Function

