Attribute VB_Name = "PrincINPALIntranet"
Option Explicit

Public gasTitulos(1 To 10) As String
Public gasComandos(1 To 10)   As String


Sub Main()
Dim sComando As String

    sComando = Command$
    
    gasComandos(1) = "c:\sge\programa\sgeprinc2.exe"
    gasComandos(2) = "http://svr7/default.asp"
    gasComandos(3) = "C:\Arquivos de programas\BrOffice.org 3\program\soffice.exe"
    
    INPALIntranet.Show
    
End Sub

