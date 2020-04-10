VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaCro"
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

Public Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objPVCro As New ClassPedidoDeVendaCro

On Error GoTo Erro_PV_Move_Pedido_Memoria

    objPVCro.sObservacao = ObservacaoPV(objCT).Text
    
    Set objPedidoVenda.objInfoUsu = objPVCro
   
    PV_Move_Pedido_Memoria = SUCESSO
     
    Exit Function
    
Erro_PV_Move_Pedido_Memoria:

    PV_Move_Pedido_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function PV_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objPVCro As ClassPedidoDeVendaCro

On Error GoTo Erro_PV_Traz_Pedido_Tela
    
    If Not (objPedidoVenda.objInfoUsu Is Nothing) Then
    
        Set objPVCro = objPedidoVenda.objInfoUsu
       
        ObservacaoPV(objCT).Text = objPVCro.sObservacao
       
    End If
       
    PV_Traz_Pedido_Tela = SUCESSO
     
    Exit Function
    
Erro_PV_Traz_Pedido_Tela:

    PV_Traz_Pedido_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function
    
End Function

Public Sub ObservacaoPV_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ObservacaoPV_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ObservacaoPV_Validate

    Exit Sub

Erro_ObservacaoPV_Validate:

    Cancel = True

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Function PV_Traz_Dados_Orcamento_Tela(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Transfere os dados do objOrcamentoVenda para objPedidoDeVenda

Dim objItemOV As ClassItemOV
Dim objItemOVCro As ClassItemOVCro
Dim iCondPagto As Integer
Dim iIndice As Integer
Dim bModificou As Boolean

On Error GoTo Erro_PV_Traz_Dados_Orcamento_Tela
   
    If objCT.Name <> "NFiscal" Then
   
        For Each objItemOV In objOrcamentoVenda.colItens
            
            Set objItemOVCro = objItemOV.objInfoUsu
            
            If objItemOVCro.iEscolhido = MARCADO Then
                
                iCondPagto = objItemOVCro.iCondPagto
                
                Exit For
                
            End If
    
        Next
        
        'Altera a condição de Pagamento
        If iCondPagto <> 0 Then
        
            objCT.CondicaoPagamento.Text = CStr(iCondPagto)
            Call objCT.CondicaoPagamento_Validate(bSGECancelDummy)
            
        End If
    
    End If
    
    bModificou = False
        
    For iIndice = objOrcamentoVenda.colItens.Count To 1 Step -1
    
        Set objItemOV = objOrcamentoVenda.colItens.Item(iIndice)
        Set objItemOVCro = objItemOV.objInfoUsu
        
        If objItemOVCro.iEscolhido = DESMARCADO Then
        
            bModificou = True
        
            objCT.GridItens.Row = iIndice
            
            Call objCT.GridItens_KeyDown(vbKeyDelete, 0, False)
        
        End If
    
    Next
    
    'Se está diferente do orçamento de venda
    If bModificou And objCT.Name <> "NFiscal" Then
        objCT.CobrancaAutomatica.Value = vbChecked
        
        Call objCT.CobrancaAutomatica_Click
    End If
    
    Exit Function

Erro_PV_Traz_Dados_Orcamento_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Function
