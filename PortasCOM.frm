VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form PortasCOM 
   Caption         =   "COM"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm ComIC 
      Left            =   330
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "PortasCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function BotaoImprimirIC_Click(ByVal iContaCorrente As Integer, ByVal iBanco As Integer, ByVal dtData As Date, ByVal sFavorecido As String, ByVal lNumCheque As Long, ByVal dValor As Double)
'Faz a ligação telefonica usando o modem

Dim lErro As Long
Dim sImpCheque As String
Dim sModelo As String
Dim objBanco As New ClassBanco
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim sCidade As String
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoImprimirIC_Click

    If ComIC.PortOpen Then gError 182910
    
    sImpCheque = String(128, 0)
    sModelo = String(128, 0)

    Call GetPrivateProfileString("Geral", "ImpChequeCOM", "2", sImpCheque, 128, "ADM100.INI")
    Call GetPrivateProfileString("Geral", "ImpChequeModelo", "NSC 2.18", sModelo, 128, "ADM100.INI")

    sImpCheque = Replace(sImpCheque, Chr(0), "")
    sModelo = Replace(sModelo, Chr(0), "")
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 182924
    
    sCidade = objFilialEmpresa.objEndereco.sCidade
    
    If dtData = DATA_NULA Then gError 182915
    If dValor = 0 Then gError 182916
    If Len(Trim(sFavorecido)) = 0 Then gError 182923
    
    If iBanco = 0 Then
    
        If iContaCorrente = 0 Then gError 182917
    
        'Le a Conta Corrente a partir de iCodigo passado como parâmetro
        lErro = CF("ContaCorrenteInt_Le", iContaCorrente, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 182918
    
        'Caso a Conta Corrente não tiver sido encontrada dispara erro
        If lErro <> SUCESSO Then gError 182919
    
        'Caso a Conta Corrente não for bancária dispara erro
        If objContaCorrenteInt.iCodBanco = 0 Then gError 182920
        
        'Atribui o valor retornado de objContaCorrenteInt.iCodBanco a objBanco.iCodBanco
        objBanco.iCodBanco = objContaCorrenteInt.iCodBanco
              
        'Le o Banco a partir de objBanco.iCodBanco
        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then gError 182921
            
        'Caso o banco não tiver sido encontrado dispara erro
        If lErro = 16091 Then gError 182922
        
        iBanco = objBanco.iCodBanco
    
    End If
    
    ComIC.CommPort = CInt(sImpCheque)
    ComIC.Settings = "9600,N,8,2"
    ComIC.PortOpen = True
       
    Select Case sModelo
    
        Case "NSC 2.18"
        
            lErro = ImpressoraDeCheque_NSC218(iBanco, sCidade, dtData, sFavorecido, lNumCheque, dValor)
            If lErro <> SUCESSO Then gError 182911
        
        Case Else
            gError 182912
    
    End Select
    
    ComIC.PortOpen = False
    
    Exit Function

Erro_BotaoImprimirIC_Click:

    'Fecha a Porta
    If ComIC.PortOpen Then ComIC.PortOpen = False

    Select Case gErr
    
        Case 8002
             Call Rotina_Erro(vbOKOnly, "ERRO_COM_INVALIDA", gErr, sImpCheque)

        Case 182910
             Call Rotina_Erro(vbOKOnly, "ERRO_IMPRESSORA_NAO_RESPONDE", gErr)
             
        Case 182911
        
        Case 182912
             Call Rotina_Erro(vbOKOnly, "ERRO_IMPCHEQUE_MODELO_NAO_TRATADO", gErr)
             
        Case 182915
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_INFORMADA", gErr)

        Case 182916
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
             
        Case 182917
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
        
        Case 182918, 182921
       
        Case 182919
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, iContaCorrente)
        
        Case 182920
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", gErr)

        Case 182922
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)

        Case 182923
            Call Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_NAO_PREENCHIDO", gErr)

        Case 182924

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182913)

    End Select

    Exit Function
    
End Function

Private Function ImpressoraDeCheque_NSC218(ByVal iBanco As Integer, ByVal sCidade As String, ByVal dtData As Date, ByVal sFavorecido As String, ByVal lNumCheque As Long, ByVal dValor As Double)
'Faz a ligação telefonica usando o modem

Dim sValor As String
Dim sData As String
Dim sNumCheque As String
Dim sBanco As String

On Error GoTo Erro_ImpressoraDeCheque_NSC218

    sValor = FormataCpoValor(dValor, 14)
    sData = Format(dtData, "ddmmyy")
    sNumCheque = FormataCpoNum(lNumCheque, 7)
    sBanco = FormataCpoNum(iBanco, 3)
       
    ComIC.Output = "ESC B " & sBanco
    ComIC.Output = "ESC C " & sCidade & "$"
    ComIC.Output = "ESC D " & sData
    ComIC.Output = "ESC F " & sFavorecido & "$"
    
    If lNumCheque <> 0 Then
        ComIC.Output = "ESC N " & sNumCheque
    End If
    
    ComIC.Output = "ESC V " & sValor
    
    Exit Function

Erro_ImpressoraDeCheque_NSC218:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182914)

    End Select

    Exit Function
    
End Function

