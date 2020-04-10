VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelectJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Cliente_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto e cliente passados

Dim lErro As Long

On Error GoTo Erro_Cliente_Le_UltimoItemPedido

    gError 500019
    
    Cliente_Le_UltimoItemPedido = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_UltimoItemPedido:

    Cliente_Le_UltimoItemPedido = gErr
    
    Select Case gErr
    
        Case 500019 'ItemPV Não encontrado
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152752)
    
    End Select
    
    Exit Function
    
End Function


Function ClienteFilial_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long, ByVal iFilialCli As Integer) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto, cliente e FilialCliente passados

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String, sUnidadeMed As String
Dim dPrecoBase As Double
Dim dPrecoUnitario As Double

On Error GoTo Erro_ClienteFilial_Le_UltimoItemPedido

    gError 94412 '500019
    
    ClienteFilial_Le_UltimoItemPedido = SUCESSO
    
    Exit Function
    
Erro_ClienteFilial_Le_UltimoItemPedido:

    ClienteFilial_Le_UltimoItemPedido = gErr
    
    Select Case gErr
    
        Case 94412 'ItemPV Não encontrado
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152753)
    
    End Select
    
    Exit Function
    
End Function




