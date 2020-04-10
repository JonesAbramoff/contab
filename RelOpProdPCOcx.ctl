VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpProdPCOcx 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ScaleHeight     =   2610
   ScaleWidth      =   6210
   Begin VB.Frame Frame4 
      Caption         =   "Pedidos de Compra"
      Height          =   720
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   3795
      Begin MSMask.MaskEdBox CodigoPCDe 
         Height          =   300
         Left            =   540
         TabIndex        =   1
         Top             =   285
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoPCAte 
         Height          =   300
         Left            =   2460
         TabIndex        =   2
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigoPCAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   330
         Width           =   360
      End
      Begin VB.Label LabelCodigoPCDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3915
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdPCOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpProdPCOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpProdPCOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpProdPCOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4215
      Picture         =   "RelOpProdPCOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   930
      Width           =   1590
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdPCOcx.ctx":0A96
      Left            =   840
      List            =   "RelOpProdPCOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2550
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produtos"
      Height          =   720
      Left            =   120
      TabIndex        =   9
      Top             =   735
      Width           =   3795
      Begin MSMask.MaskEdBox CodigoProdDe 
         Height          =   300
         Left            =   570
         TabIndex        =   14
         Top             =   255
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoProdAte 
         Height          =   300
         Left            =   2430
         TabIndex        =   15
         Top             =   240
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigoProdDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   315
         Width           =   435
      End
      Begin VB.Label LabelCodigoProdAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1980
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "RelOpProdPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoPCDe As AdmEvento
Attribute objEventoPCDe.VB_VarHelpID = -1
Private WithEvents objEventoPCAte As AdmEvento
Attribute objEventoPCAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 74538
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 74539

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 74538
        
        Case 74539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171708)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub CodigoPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCAte, iAlterado)
    
End Sub

Private Sub CodigoPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCDe, iAlterado)
    
End Sub


Private Sub LabelCodigoPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodigoPCAte_Click

    If Len(Trim(CodigoPCAte.Text)) > 0 Then
        
        objPedCompra.lCodigo = StrParaLong(CodigoPCAte.Text)
    End If

    'Chama Tela PedCompraTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoPCAte)

   Exit Sub

Erro_LabelCodigoPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171709)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodigoPCDe_Click

    If Len(Trim(CodigoPCDe.Text)) > 0 Then
        
        objPedCompra.lCodigo = StrParaLong(CodigoPCDe.Text)
    End If

    'Chama Tela PedCompraTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoPCDe)

   Exit Sub

Erro_LabelCodigoPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171710)

    End Select

    Exit Sub

End Sub


Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 74540
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 74540
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171711)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome


On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
        
    Set objEventoPCDe = New AdmEvento
    Set objEventoPCAte = New AdmEvento
    
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdDe)
    If lErro <> SUCESSO Then gError 74541

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdAte)
    If lErro <> SUCESSO Then gError 74542

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 74541, 74542, 74812
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171712)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
        
    Set objEventoPCDe = Nothing
    Set objEventoPCAte = Nothing
    
End Sub



Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74543
    
    CodigoProdAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr
    
        Case 74543
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171713)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74544
    
    CodigoProdDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr
    
        Case 74544
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171714)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 74545

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74546

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 74547
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 74548
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 74545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 74546 To 74548
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171715)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 74549

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 74550

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 74549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 74550

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171716)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74551

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 74551

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171717)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sPC_I As String
Dim sPC_F As String
Dim sCheckTipoProd As String
Dim sProdutoTipo As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sPC_I, sPC_F)
    If lErro <> SUCESSO Then gError 74552

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 74553
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 74554
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 74555
         
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 74556
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 74557
         
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sPC_I, sPC_F)
    If lErro <> SUCESSO Then gError 74560

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 74552 To 74560
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171718)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sPC_I As String, sPC_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_F As Integer
Dim iProdPreenchido_I As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", CodigoProdDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 74984

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", CodigoProdAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 74985

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 74561

    End If
    
    'critica PedCompra Inicial e Final
    If CodigoPCDe.Text <> "" Then
        sPC_I = CStr(CodigoPCDe.Text)
    Else
        sPC_I = ""
    End If
    
    If CodigoPCAte.Text <> "" Then
        sPC_F = CStr(CodigoPCAte.Text)
    Else
        sPC_F = ""
    End If
            
    If sPC_I <> "" And sPC_F <> "" Then
        
        If StrParaLong(sPC_I) > StrParaLong(sPC_F) Then gError 74562
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 74561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            CodigoProdDe.SetFocus
            
        Case 74562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodigoPCDe.SetFocus
        
        Case 74563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 74984
            CodigoProdDe.SetFocus
            
        Case 74985
            CodigoProdAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171719)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sPC_I As String, sPC_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sProd_I <> "" Then sExpressao = "CodProduto >= " & Forprint_ConvTexto((sProd_I))
   
    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodProduto <= " & Forprint_ConvTexto((sProd_F))

    End If
   
    If sPC_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PedComCod >= " & Forprint_ConvLong(StrParaLong(sPC_I))
    
    End If
    
    If sPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PedComCod <= " & Forprint_ConvLong(StrParaLong(sPC_F))

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171720)

    End Select

    Exit Function

End Function


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoProduto As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74564
   
    'pega  Codigo Produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 74565
                   
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sParam
    CodigoProdDe.PromptInclude = True
    
    'pega  Codigo Produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 74566
                   
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sParam
    CodigoProdAte.PromptInclude = True
    
    'pega PedCompra Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 74567
                   
    CodigoPCDe.Text = sParam
    
    'pega PedCompra Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 74568
                   
    CodigoPCAte.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74564 To 74570
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171721)

    End Select

    Exit Function

End Function

Private Sub CodigoProdDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdDe_Validate

    If Len(Trim(CodigoProdDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74571
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74572
        
        If lErro = 28030 Then gError 74573
        
    End If
    
    Exit Sub
    
Erro_CodigoProdDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74571, 74572
        
        Case 74573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171722)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodigoProdAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdAte_Validate

    If Len(Trim(CodigoProdAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74574
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74575
        
        If lErro = 28030 Then gError 74576
        
    End If
    
    Exit Sub
    
Erro_CodigoProdAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74574, 74575
        
        Case 74576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171723)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub LabelCodigoProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdDe_Click
    
    If Len(Trim(CodigoProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74577
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoDe)

   Exit Sub

Erro_LabelCodigoProdDe_Click:

    Select Case gErr

        Case 74577
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171724)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdAte_Click
    
    If Len(Trim(CodigoProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74578
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProdutoAte)

   Exit Sub

Erro_LabelCodigoProdAte_Click:

    Select Case gErr

        Case 74578
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171725)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Produtos x Pedidos de Compra"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProdutosPC"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub objEventoPCAte_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodigoPCAte.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub
Private Sub objEventoPCDe_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodigoPCDe.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoProdDe Then
            Call LabelCodigoProdDe_Click
            
        ElseIf Me.ActiveControl Is CodigoProdAte Then
            Call LabelCodigoProdAte_Click
           
        ElseIf Me.ActiveControl Is CodigoPCDe Then
            Call LabelCodigoPCDe_Click
        
        ElseIf Me.ActiveControl Is CodigoPCAte Then
            Call LabelCodigoPCAte_Click
        
        End If
    
    End If

End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub









Private Sub LabelCodigoPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCAte, Source, X, Y)
End Sub

Private Sub LabelCodigoPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCDe, Source, X, Y)
End Sub

Private Sub LabelCodigoPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdAte, Source, X, Y)
End Sub

Private Sub LabelCodigoProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdDe, Source, X, Y)
End Sub

Private Sub LabelCodigoProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdDe, Button, Shift, X, Y)
End Sub

