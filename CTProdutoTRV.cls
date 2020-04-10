VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdutoTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Property Get PercComissao(ByVal objCT As Object) As Object
     Set PercComissao = objCT.objUserControl.Controls("PercComissao")
End Property

Public Property Get Vendedor(ByVal objCT As Object) As Object
     Set Vendedor = objCT.objUserControl.Controls("Vendedor")
End Property

Public Function Prod_Form_UnLoad(ByVal objCT As Object) As Long

    Set objEventoVendedor = Nothing
    Set gobjCT = Nothing

    Prod_Form_UnLoad = SUCESSO
    
End Function

Public Function Prod_Gravar_Registro(ByVal objCT As Object, objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoTRV As New ClassProdutoTRV
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Prod_Gravar_Registro

    If Len(Trim(Vendedor(objCT).Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor(objCT).Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 197229

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 197230

        objProdutoTRV.iCodVendedor = objVendedor.iCodigo

    End If
    
    If Len(Trim(PercComissao(objCT).Text)) > 0 Then objProdutoTRV.dPercComissao = CDbl(PercComissao(objCT).Text) / 100

    Set objProduto.objInfoUsu = objProdutoTRV

    Prod_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Prod_Gravar_Registro:
    
    Prod_Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 197229
        
        Case 197230
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197231)

    End Select

    Exit Function

End Function

Public Function Prod_Traz_Produto_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoTRV As ClassProdutoTRV
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Prod_Traz_Produto_Tela

    lErro = CF("Produto_Le_Customizado", objProduto)
    If lErro <> SUCESSO Then gError 197232
    
    If Not (objProduto.objInfoUsu Is Nothing) Then
    
        Set objProdutoTRV = objProduto.objInfoUsu
        
        If objProdutoTRV.iCodVendedor <> 0 Then
        
            Vendedor(objCT).Text = objProdutoTRV.iCodVendedor
        
            lErro = TP_Vendedor_Le(Vendedor(objCT), objVendedor)
            If lErro <> SUCESSO Then gError 197233
        
        End If
        
        PercComissao(objCT).Text = CStr(100 * objProdutoTRV.dPercComissao)
        
    End If
    
    Prod_Traz_Produto_Tela = SUCESSO
     
    Exit Function
    
Erro_Prod_Traz_Produto_Tela:

    Prod_Traz_Produto_Tela = gErr
     
    Select Case gErr
    
        Case 197232, 197233
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197234)
     
    End Select
     
    Exit Function

End Function

Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is Vendedor(objCT) Then
            Call LabelVendedor_Click(objCT)
        End If
    
    End If

End Sub

Public Sub Vendedor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Vendedor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor(objCT).Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor(objCT), objVendedor)
        If lErro <> SUCESSO Then gError 197226
        
        If objVendedor.iAtivo = DESMARCADO Then gError 197227

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 197226
        
        Case 197227
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197228)
    
    End Select

End Sub

Public Sub LabelVendedor_Click(ByVal objCT As Object)

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor(objCT).Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor(objCT).Text
    
    Set gobjCT = objCT

    Set objEventoVendedor = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor(gobjCT).Text = objVendedor.sNomeReduzido

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub PercComissao_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComissao_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComissao(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComissao(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    PercComissao(objCT).Text = Format(PercComissao(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComissao_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub
