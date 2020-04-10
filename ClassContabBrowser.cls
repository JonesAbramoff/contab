VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContabBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gsSiglaModuloChamador As String

Function ContaCclLista_BotaoEdita(objContaCcl As ClassContaCcl, lErro As Long) As Long

On Error GoTo Erro_ContaCclLista_BotaoEdita

    If lErro = SUCESSO Then
    
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            Call Chama_Tela("ContaCcl", objContaCcl)
        
        Else
        
            Call Chama_Tela("ContaCcl2", objContaCcl)
            
        End If
    
    End If
    
    ContaCclLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_ContaCclLista_BotaoEdita:

    ContaCclLista_BotaoEdita = Err
    
    Select Case Err

        Case 14274
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147799)

    End Select

    Exit Function

End Function

'Function LotePendenteLista_BotaoEdita(objLote As ClassLote, lErro As Long) As Long
'
'On Error GoTo Erro_LotePendenteLista_BotaoEdita
'
'    gsSiglaModuloChamador = StringZ(gsSiglaModuloChamador)
'
'    If gsSiglaModuloChamador <> "" Then
'        Call Chama_Tela("LoteTela", objLote, gsSiglaModuloChamador)
'    Else
'        Call Chama_Tela("LoteTela", objLote)
'    End If
'
'    Exit Function
'
'Erro_LotePendenteLista_BotaoEdita:
'
'    Select Case Err
'
'        Case 9197
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147800)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function LotePendenteLista_Trata_Parametros(Optional objTela As Object, Optional colSelecao As Collection, Optional objLote As ClassLote, Optional objEvento1 As AdmEvento) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_LotePendenteLista_Trata_Parametros
'
'    objTela.gsSiglaModuloChamador = StringZ(objLote.sOrigem)
'
'    LotePendenteLista_Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_LotePendenteLista_Trata_Parametros:
'
'    LotePendenteLista_Trata_Parametros = Err
'
'    Select Case Err
'
'        Case Else
'
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147801)
'
'    End Select
'
'    Exit Function
'
'
'
'End Function
