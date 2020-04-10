VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalRemPedidoArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'precodesc
Public Property Get PrecoDesc(ByVal objCT As Object) As Object
     Set PrecoDesc = objCT.objUserControl.Controls("PrecoDesc")
End Property

Public Function NFRP_Form_Load2(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFRP_Form_Load2

'    Call CF2(objct.objCTNFiscalPedido, "NFCF_Form_Load_PrecoDesc", "", 0)
    
    NFRP_Form_Load2 = SUCESSO
    
    Exit Function
    
Erro_NFRP_Form_Load2:

    NFRP_Form_Load2 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184276)

    End Select
    
    Exit Function

End Function

Public Function NFRP_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NFRP_Inicializa_Grid_Itens = CF2(objCT.objCTNFiscalPedido, "NFP_Inicializa_Grid_Itens", objGridInt)
End Function

Public Sub BotaoImprimirPI_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoImprimirPI_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(objCT.objCTNFiscalPedido.Pedido.ClipText)) > 0 Then objPedidoVenda.lCodigo = CLng(objCT.objCTNFiscalPedido.Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(objCT.objCTNFiscalPedido.FilialPedido.Text)
       
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 132437

    If lErro = 26509 Then gError 132438
    
    lErro = objRelatorio.ExecutarDireto("Pedido Interno", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVend4", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 132439
 
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimirPI_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 132437, 132439
        
        Case 132438
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177043)

    End Select

    Exit Sub

End Sub
