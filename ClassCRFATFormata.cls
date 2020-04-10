VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Sub Converte_Status(ByVal iStatus As Integer, sStatus As String)
'converte o tipo de status do valor inteiro para o seu correspondente textual

    If iStatus = STATUS_LANCADO Then
        sStatus = STRING_STATUS_LANCADO
    ElseIf iStatus = STATUS_BAIXADO Then
        sStatus = STRING_STATUS_BAIXADO
    ElseIf iStatus = STATUS_CANCELADO Then
        sStatus = STRING_STATUS_CANCELADO
    End If

End Sub

Public Function Obtem_ViaTransp_Descricao(sViaTransp As String, ByVal iViaTransp As Integer) As Long
'Retorna a descrição de uma via de transporte a partir do codigo da via passado
'sViaTransp RETORNA (Output) a descrição da via de transporte
'iViaTransp RECEBE (Input) o código da via de transporte

On Error GoTo Erro_Obtem_ViaTransp_Descricao

    Select Case iViaTransp
    
        Case TRANSPORTE_AEREO
            sViaTransp = TRANSPORTE_AEREO_DESC
        Case TRANSPORTE_MARITIMO
            sViaTransp = TRANSPORTE_MARITIMO_DESC
        Case TRANSPORTE_RODOVIARIO
            sViaTransp = TRANSPORTE_RODOVIARIO_DESC
        Case TRANSPORTE_FERROVIARIO
            sViaTransp = TRANSPORTE_FERROVIARIO_DESC
        Case TRANSPORTE_FLUVIAL
            sViaTransp = TRANSPORTE_FLUVIAL_DESC
        Case TRANSPORTE_AEREOFLUVIAL
            sViaTransp = TRANSPORTE_AEREOFLUVIAL_DESC
        Case TRANSPORTE_RODOVIARIOFERROVIARIO
            sViaTransp = TRANSPORTE_RODOVIARIOFERROVIARIO_DESC
        Case Else
            sViaTransp = TRANSPORTE_OUTROS_DESC
    
    End Select
    
    Obtem_ViaTransp_Descricao = SUCESSO
    
    Exit Function

Erro_Obtem_ViaTransp_Descricao:
    
    Obtem_ViaTransp_Descricao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148829)
            
    End Select
    
    Exit Function
    
End Function

