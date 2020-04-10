VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaAla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela
Public Property Get DataEntregaReal(ByVal objCT As Object) As Object
     Set DataEntregaReal = objCT.objUserControl.Controls("DataEntregaReal")
End Property

Public Property Get UpDownEntregaReal(ByVal objCT As Object) As Object
     Set UpDownEntregaReal = objCT.objUserControl.Controls("UpDownEntregaReal")
End Property

'--- fim dos properties get dos controles da tela

Public Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objPVAla As New ClassPedidoDeVendaAla

On Error GoTo Erro_PV_Move_Pedido_Memoria

    objPVAla.dtDataEntregaReal = StrParaDate(DataEntregaReal(objCT).Text)
    
    Set objPedidoVenda.objInfoUsu = objPVAla
   
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
Dim objPVAla As ClassPedidoDeVendaAla

On Error GoTo Erro_PV_Traz_Pedido_Tela
    
    If Not (objPedidoVenda.objInfoUsu Is Nothing) Then
    
        Set objPVAla = objPedidoVenda.objInfoUsu
        
        Call DateParaMasked(DataEntregaReal(objCT), objPVAla.dtDataEntregaReal)
       
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

Public Sub DataEntregaReal_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEntregaReal_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataEntregaReal(objCT), objCT.iAlterado)
End Sub

Public Sub DataEntregaReal_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long, iIndice As Integer, dtData As Date

On Error GoTo Erro_DataEntregaReal_Validate

    'Verifica se a Data foi digitada
    If Len(Trim(DataEntregaReal(objCT).ClipText)) <> 0 Then
    
        'Critica a data digitada
        lErro = Data_Critica(DataEntregaReal(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub

Erro_DataEntregaReal_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158095)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaReal_DownClick(ByVal objCT As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaReal_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaReal(objCT), DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEntregaReal_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158096)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaReal_UpClick(ByVal objCT As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaReal_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaReal(objCT), AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEntregaReal_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158097)

    End Select

    Exit Sub

End Sub
