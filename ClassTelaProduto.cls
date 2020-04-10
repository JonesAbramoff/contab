VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Produto_Perde_Foco(ByVal Produto As Object, ByVal Desc As Object) As Long
'recebe MaskEdBox do Produto e o label da descrição

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Perde_Foco

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", Produto.Text, sProdFormatado, iProdPreenchido)
    If lErro Then Error 27093

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 27094

        If lErro = 28030 Then

            Produto.PromptInclude = False
            Produto.Text = ""
            Produto.PromptInclude = True
            Produto.SetFocus
            Desc.Caption = ""

            Error 27095

        End If

        Desc.Caption = objProduto.sDescricao

    Else

        Desc.Caption = ""

    End If

    Produto_Perde_Foco = SUCESSO

    Exit Function

Erro_Produto_Perde_Foco:

    Produto_Perde_Foco = Err

    Select Case Err

        Case 27093, 27094, 27095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153602)

    End Select

    Exit Function

End Function

Public Function Traz_Produto_MaskEd(ByVal sProduto As String, ByVal MaskEdProduto As Object, ByVal Label As Object) As Long

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoEnxuto As String

On Error GoTo Erro_Traz_Produto_MaskEd

    sProdutoEnxuto = String(STRING_PRODUTO, 0)

    If sProduto <> "" Then
        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro Then Error 27096
    Else
        sProdutoEnxuto = "0"
    End If

    If sProdutoEnxuto <> "0" Then

        'le o Produto
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 27097

        'nao encontrou o Produto
        If lErro = 28030 Then Error 27098

        MaskEdProduto.PromptInclude = False
        MaskEdProduto.Text = sProdutoEnxuto
        MaskEdProduto.PromptInclude = True
        Label.Caption = objProduto.sDescricao
    Else
        MaskEdProduto.PromptInclude = False
        MaskEdProduto.Text = ""
        MaskEdProduto.PromptInclude = True
        Label.Caption = ""
    End If

    Traz_Produto_MaskEd = SUCESSO

    Exit Function

Erro_Traz_Produto_MaskEd:

    Traz_Produto_MaskEd = Err

    Select Case Err

        Case 27098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, sProduto)

        Case 27097 'erro já tratado

        Case 27096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153603)

    End Select

    Exit Function

End Function
Function Inicializa_Mascara_Produto_MaskEd(ByVal MaskEdProduto As Object) As Long
'inicializa a mascara de Produto

Dim sMascaraProduto As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Produto_MaskEd

    'Inicializa a máscara de Produto
    sMascaraProduto = String(STRING_PRODUTO, 0)

    'le a mascara das Produtos
    lErro = MascaraProduto(sMascaraProduto)
    If lErro Then Error 27092

    MaskEdProduto.Mask = ">" & sMascaraProduto
    MaskEdProduto.AllowPrompt = True
    
    If MaskEdProduto.BorderStyle = 0 Then
        If gobjCTB.iESTTamanhoMinimoProdutoGrid > MaskEdProduto.Width Then
            MaskEdProduto.Width = gobjCTB.iESTTamanhoMinimoProdutoGrid
        End If
    End If

    Inicializa_Mascara_Produto_MaskEd = SUCESSO

    Exit Function

Erro_Inicializa_Mascara_Produto_MaskEd:

    Inicializa_Mascara_Produto_MaskEd = Err

    Select Case Err

        Case 27092

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153604)

    End Select

    Exit Function

End Function


