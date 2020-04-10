VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedVendCCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela
Public Property Get ObservacaoPV(ByVal objCT As Object) As Object
     Set ObservacaoPV = objCT.objUserControl.Controls("ObservacaoPV")
End Property
'--- fim dos properties get dos controles da tela

Public Function PVC_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objPVCro As ClassPedidoDeVendaCro

On Error GoTo Erro_PVC_Traz_Pedido_Tela

    If Not (objPedidoVenda.objInfoUsu Is Nothing) Then
    
        Set objPVCro = objPedidoVenda.objInfoUsu
       
        ObservacaoPV(objCT).Text = objPVCro.sObservacao
       
    End If
       
    PVC_Traz_Pedido_Tela = SUCESSO
     
    Exit Function
    
Erro_PVC_Traz_Pedido_Tela:

    PVC_Traz_Pedido_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function
    
End Function
