VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpSaldoEstoque_LOcx 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   KeyPreview      =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   9045
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
      Left            =   4635
      Picture         =   "RelOpSaldoEstoque_LOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpSaldoEstoque_LOcx.ctx":0102
      Left            =   1485
      List            =   "RelOpSaldoEstoque_LOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   2916
   End
   Begin VB.ComboBox ComboTotaliza 
      Height          =   315
      ItemData        =   "RelOpSaldoEstoque_LOcx.ctx":0106
      Left            =   1485
      List            =   "RelOpSaldoEstoque_LOcx.ctx":0110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   855
      Width           =   2520
   End
   Begin VB.CheckBox ProdutosZerados 
      Caption         =   "Exibe Produtos Zerados no Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   6
      Top             =   3795
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   840
      Left            =   120
      TabIndex        =   19
      Top             =   2820
      Width           =   5655
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   690
         TabIndex        =   4
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxarifadoFinal 
         Height          =   315
         Left            =   3345
         TabIndex        =   5
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label labelAlmoxarifadoFinal 
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
         Height          =   195
         Left            =   2925
         TabIndex        =   21
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Left            =   315
         TabIndex        =   20
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1332
      Left            =   120
      TabIndex        =   14
      Top             =   1260
      Width           =   5655
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   3
         Top             =   885
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   2
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   17
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label LabelProdutoDe 
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
         Height          =   255
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   420
         Width           =   615
      End
      Begin VB.Label LabelProdutoAte 
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
         Height          =   255
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   930
         Width           =   555
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   2790
      ItemData        =   "RelOpSaldoEstoque_LOcx.ctx":012B
      Left            =   5940
      List            =   "RelOpSaldoEstoque_LOcx.ctx":012D
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1215
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpSaldoEstoque_LOcx.ctx":012F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpSaldoEstoque_LOcx.ctx":02AD
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpSaldoEstoque_LOcx.ctx":07DF
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpSaldoEstoque_LOcx.ctx":0969
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      Height          =   255
      Left            =   780
      TabIndex        =   24
      Top             =   420
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Totaliza por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   23
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5970
      TabIndex        =   22
      Top             =   960
      Width           =   1185
   End
End
Attribute VB_Name = "RelOpSaldoEstoque_LOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim giProdInicial As Integer
Dim giAlmoxInicial As Integer
Dim iAlmoxarifadoInicialAlterado As Integer
Dim iAlmoxarifadoFinalAlterado As Integer

Private Sub AlmoxarifadoInicial_GotFocus()
'Mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoInicial_GotFocus

    giAlmoxInicial = 1

    Exit Sub

Erro_AlmoxarifadoInicial_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173281)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoFinal_GotFocus()
'mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoFinal_GotFocus

    giAlmoxInicial = 0

    Exit Sub

Erro_AlmoxarifadoFinal_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173282)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_GotFocus()
'Mostra a arvore de produtos

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_GotFocus

    giProdInicial = 1

    Exit Sub

Erro_ProdutoInicial_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173283)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_GotFocus()
'Mostra a arvore de produtos

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_GotFocus

    giProdInicial = 0

    Exit Sub

Erro_ProdutoFinal_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173284)

    End Select

    Exit Sub

End Sub
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 65136

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then Error 65137

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then Error 65138

    Call DefinePadrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   
   lErro_Chama_Tela = Err

    Select Case Err

        Case 65136, 65137, 65138

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173285)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iFilialInic, iFilialFim As Integer
Dim iTotaliza As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 65140

    'pega Produto Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 65141

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 65142

    'pega parâmetro Produto Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 65143

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 65144
    
    'pega parâmetro Almoxarifado Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro Then Error 65145
    If sParam > 0 Then
        AlmoxarifadoInicial.Text = sParam
    Else
        AlmoxarifadoInicial.Text = ""
    End If
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro Then Error 65146
    If sParam > 0 Then
        AlmoxarifadoFinal.Text = sParam
    Else
        AlmoxarifadoFinal.Text = ""
    End If
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)

    'pega parâmetro de totalização
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTOTALIZA", sParam)
    If lErro Then Error 65147
   
    'seleciona ítem no ComboTotaliza
    iTotaliza = CInt(sParam)
    ComboTotaliza.ListIndex = iTotaliza
    
    'pega parâmetro de ProdutosZerados
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NPRODZERADO", sParam)
    If lErro Then Error 65148
   
    'seleciona ítem na CheckBox ProdutosZerados
    ProdutosZerados.Value = CInt(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 65140 To 65148
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173286)

    End Select

    Exit Function

End Function

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa",giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 65149

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 65149

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173287)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub


Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82430

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82431

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82432

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82430, 82432

        Case 82431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173288)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82478

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82479

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82480

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82478, 82480

        Case 82479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173289)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata",ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82515

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82515

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173290)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata",ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82514

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82514

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173291)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 65150
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 65151

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 65151
        
        Case 65150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173292)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""

    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sAlmox_I As String, sAlmox_F As String, sFilial_I As String, sFilial_F As String) As Long
'Formata os parâmetros de produto
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 65153

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 65154

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 65155

    End If
    
    'critica Almoxarifado Inicial e Final
    If AlmoxarifadoInicial.Text <> "" Then
        sAlmox_I = CStr(Codigo_Extrai(AlmoxarifadoInicial.Text))
        
    Else
        sAlmox_I = ""
        
    End If
    
        
    If AlmoxarifadoFinal.Text <> "" Then
        sAlmox_F = CStr(Codigo_Extrai(AlmoxarifadoFinal.Text))
    
    Else
        sAlmox_F = ""
        
    End If
    
    
    If sAlmox_I <> "" And sAlmox_F <> "" Then
          
        If sAlmox_I <> "" And sAlmox_F <> "" Then
        
            If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 65152
        
        End If
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
    
        Case 65152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)

        Case 65153
            ProdutoInicial.SetFocus

        Case 65154
            ProdutoFinal.SetFocus

        Case 65155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173293)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    ProdutosZerados.Value = 1
    Call DefinePadrao
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim sAlmox_I As String
Dim sAlmox_F As String
Dim sFilial_I As String
Dim sFilial_F As String
Dim iAlmoxInicial As Integer
Dim iAlmoxFinal As Integer
Dim sTotaliza As String
Dim sProd_zerado As String
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sAlmox_I, sAlmox_F, sFilial_I, sFilial_F)
    If lErro <> SUCESSO Then Error 65161
      
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 65162

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 65163

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 65164
    
    sTotaliza = CStr(ComboTotaliza.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NTOTALIZA", sTotaliza)
    If lErro <> AD_BOOL_TRUE Then Error 65165
    
    iAlmoxInicial = Codigo_Extrai(AlmoxarifadoInicial.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", CStr(iAlmoxInicial))
    If lErro <> AD_BOOL_TRUE Then Error 65166
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65167
    
    iAlmoxFinal = Codigo_Extrai(AlmoxarifadoFinal.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", CStr(iAlmoxFinal))
    If lErro <> AD_BOOL_TRUE Then Error 65168
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65169
    
    sProd_zerado = CStr(ProdutosZerados.Value)
    
    lErro = objRelOpcoes.IncluirParametro("NPRODZERADO", sProd_zerado)
    If lErro <> AD_BOOL_TRUE Then Error 65170
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
        'Ler o mês e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
        lErro = CF("EstoqueMes_Le_Aberto",objEstoqueMes)
        If lErro <> SUCESSO And lErro <> 40673 Then Error 65171

        If lErro = 40673 Then Error 65160
 
        lErro = objRelOpcoes.IncluirParametro("NANO", objEstoqueMes.iAno)
        If lErro <> AD_BOOL_TRUE Then Error 65172
 
        lErro = objRelOpcoes.IncluirParametro("NMES", objEstoqueMes.iMes)
        If lErro <> AD_BOOL_TRUE Then Error 65173
    
        lErro = CF("EstoqueMes_Le_Apurado",objEstoqueMes)
        If lErro <> SUCESSO And lErro <> 46225 Then Error 65174
        
        If lErro = 46225 Then
            objEstoqueMes.iAno = 0
            objEstoqueMes.iMes = 0
        End If

        lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
        If lErro <> AD_BOOL_TRUE Then Error 65175
 
        lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
        If lErro <> AD_BOOL_TRUE Then Error 65176
    
    End If
    
    If ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "SldEsAll"
    If ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "saldestl"
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, iAlmoxInicial, iAlmoxFinal, sFilial_I, sFilial_F, sTotaliza, sProd_zerado)
    If lErro <> SUCESSO Then Error 65177

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err
                   
        Case 65160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAOEXISTE_MES_ABERTO", Err)
    
        Case 65161 To 65177
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173294)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 65178

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 65179

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 65178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 65179

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173295)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 65180

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 65180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173296)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 65181

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 65182

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 65183

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 65181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 65182, 65183

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173297)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 65184
    
    If lErro <> SUCESSO Then Error 65185

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 65184

        Case 65185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173298)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 65186
    
    If lErro <> SUCESSO Then Error 65187

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 65186

        Case 65187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173299)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim lErro As Long
Dim sListBoxItem As String
Dim objCodigoDescricao As New AdmCodigoNome
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objAlmoxSelecionado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_DblClick

    'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
 
    If giAlmoxInicial = 1 Then
    
        AlmoxarifadoInicial.Text = sListBoxItem
        
    Else
        AlmoxarifadoFinal.Text = sListBoxItem

    End If

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173300)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, iAlmoxInicial As Integer, iAlmoxFinal As Integer, sFilial_I As String, sFilial_F As String, sTotaliza As String, sProdZerado As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If

    If iAlmoxInicial <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado >= " & Forprint_ConvInt(iAlmoxInicial)

    End If

    If iAlmoxFinal <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado <= " & Forprint_ConvInt(iAlmoxFinal)

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173301)

    End Select

    Exit Function

End Function

Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If iAlmoxarifadoInicialAlterado = 1 Then

        If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then

            'Tenta ler o Almoxarifado (NomeReduzido ou Código)
            lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
            If lErro <> SUCESSO Then Error 65188

        End If
        
        iAlmoxarifadoInicialAlterado = 0

    End If

    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 65188
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173302)

    End Select

End Sub

Private Sub AlmoxarifadoInicial_Change()

    iAlmoxarifadoInicialAlterado = 1

End Sub

Private Sub AlmoxarifadoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoFinal_Validate

    If iAlmoxarifadoFinalAlterado = 1 Then

        If Len(Trim(AlmoxarifadoFinal.Text)) > 0 Then

            'Tenta ler o Almoxarifado (NomeReduzido ou Código)
            lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoFinal, objAlmoxarifado)
            If lErro <> SUCESSO Then Error 65189

        End If
        
        iAlmoxarifadoFinalAlterado = 0

    End If

    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 65189
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173303)

    End Select

End Sub

Private Sub AlmoxarifadoFinal_Change()

    iAlmoxarifadoFinalAlterado = 1

End Sub

Sub DefinePadrao()
'Preenche a tela com as opções padrão de FilialEmpresa e totalização

Dim iIndice As Integer

    giProdInicial = 1
    giAlmoxInicial = 1

    ComboTotaliza.ListIndex = 1
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE_L
    Set Form_Load_Ocx = Me
    Caption = "Saldo em Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpSaldoEstoque_L"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
                
    End If

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
   ' Parent.UnloadDoFilho
    
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



Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelAlmoxarifadoFinal, Source, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelAlmoxarifadoFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


