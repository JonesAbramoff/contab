VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpSaldoEstoqueOcx 
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   8805
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpSaldoEstoqueOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpSaldoEstoqueOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpSaldoEstoqueOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpSaldoEstoqueOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   3375
      ItemData        =   "RelOpSaldoEstoqueOcx.ctx":0994
      Left            =   5940
      List            =   "RelOpSaldoEstoqueOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1095
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1332
      Left            =   120
      TabIndex        =   17
      Top             =   1140
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   930
         Width           =   555
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   405
         Width           =   615
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   21
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   22
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   840
      Left            =   120
      TabIndex        =   18
      Top             =   3540
      Width           =   5655
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   690
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
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
         TabIndex        =   23
         Top             =   375
         Width           =   315
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
         TabIndex        =   24
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filiais"
      Height          =   840
      Left            =   120
      TabIndex        =   16
      Top             =   2580
      Width           =   5655
      Begin VB.ComboBox FilialEmpresaFinal 
         Height          =   315
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   345
         Width           =   2040
      End
      Begin VB.ComboBox FilialEmpresaInicial 
         Height          =   315
         Left            =   705
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   345
         Width           =   2040
      End
      Begin VB.Label Label6 
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
         Left            =   2940
         TabIndex        =   25
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label7 
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
         Left            =   330
         TabIndex        =   26
         Top             =   390
         Width           =   315
      End
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
      TabIndex        =   8
      Top             =   4515
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin VB.ComboBox ComboTotaliza 
      Height          =   315
      ItemData        =   "RelOpSaldoEstoqueOcx.ctx":0998
      Left            =   1485
      List            =   "RelOpSaldoEstoqueOcx.ctx":09A5
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   735
      Width           =   2520
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpSaldoEstoqueOcx.ctx":09C8
      Left            =   1485
      List            =   "RelOpSaldoEstoqueOcx.ctx":09CA
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2916
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
      Left            =   4635
      Picture         =   "RelOpSaldoEstoqueOcx.ctx":09CC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1575
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
      TabIndex        =   27
      Top             =   840
      Width           =   1185
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
      TabIndex        =   28
      Top             =   750
      Width           =   1080
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
      TabIndex        =   29
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpSaldoEstoqueOcx"
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173255)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173256)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173257)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173258)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Preenche as combos de filial Empresa guardando no itemData o codigo
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then Error 34037

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 34038

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then Error 34039

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then Error 34040

    Call DefinePadrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   
   lErro_Chama_Tela = Err

    Select Case Err

        Case 34037, 34038, 34039, 34040

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173259)

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
    If lErro Then Error 34042

    'pega Produto Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 34043

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 34044

    'pega parâmetro Produto Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 34045

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 34046
    
    'pega parâmetro Almoxarifado Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro Then Error 34047
    If sParam > 0 Then
        AlmoxarifadoInicial.Text = sParam
    Else
        AlmoxarifadoInicial.Text = ""
    End If
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro Then Error 34048
    If sParam > 0 Then
        AlmoxarifadoFinal.Text = sParam
    Else
        AlmoxarifadoFinal.Text = ""
    End If
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)

    'pega parâmetro FilialEmpresa Inicial
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NFILIALINIC", sParam)
    If lErro Then Error 34049
    
    FilialEmpresaInicial.Text = sParam
    Call FilialEmpresaInicial_Validate(bSGECancelDummy)
        
    'pega parâmetro FilialEmpresa Final
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
    If lErro Then Error 34050
   
    FilialEmpresaFinal.Text = sParam
    Call FilialEmpresaFinal_Validate(bSGECancelDummy)
    
    'pega parâmetro de totalização
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTOTALIZA", sParam)
    If lErro Then Error 34051
   
    'seleciona ítem no ComboTotaliza
    iTotaliza = CInt(sParam)
    ComboTotaliza.ListIndex = iTotaliza
    
    'pega parâmetro de ProdutosZerados
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NPRODZERADO", sParam)
    If lErro Then Error 34054
   
    'seleciona ítem na CheckBox ProdutosZerados
    ProdutosZerados.Value = CInt(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 34042
        
        'erro ObterParametro
        Case 34043, 34045, 34047, 34048, 34049, 34050, 34051, 34054
         
        Case 34044, 34046

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173260)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long
'Carrega as Combos FilialEmpresaInicial e FilialEmpresaFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le","FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 34053
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            FilialEmpresaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaInicial.ItemData(FilialEmpresaInicial.NewIndex) = objCodigoNome.iCodigo
    
            FilialEmpresaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaFinal.ItemData(FilialEmpresaFinal.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = Err

    Select Case Err

        'Erro já tratado
        Case 34053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173261)

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
    If lErro <> SUCESSO Then Error 34058

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

        Case 34058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173262)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82433

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82434

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82435

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82433, 82435

        Case 82434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173263)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82481

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82482

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82483

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82481, 82483

        Case 82482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173264)

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
        If lErro <> SUCESSO Then gError 82517

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82517

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173265)

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
        If lErro <> SUCESSO Then gError 82516

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82516

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173266)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29885
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34035

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34035
        
        Case 29885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173267)

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
    If lErro <> SUCESSO Then Error 34059

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 34060

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 34061

    End If
    
    'critica FilialEmpresa Inicial e Final
    If FilialEmpresaInicial.ListIndex <> -1 Then
        sFilial_I = CStr(FilialEmpresaInicial.ItemData(FilialEmpresaInicial.ListIndex))
    Else
        sFilial_I = ""
    End If
    
    If FilialEmpresaFinal.ListIndex <> -1 Then
        sFilial_F = CStr(FilialEmpresaFinal.ItemData(FilialEmpresaFinal.ListIndex))
    Else
        sFilial_F = ""
    End If
            
    If sFilial_I <> "" And sFilial_F <> "" Then
        
        If CInt(sFilial_I) > CInt(sFilial_F) Then Error 34056
        
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
        
            If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 34057
        
        End If
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
    
        Case 34056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", Err)
   
        Case 34057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)

        Case 34059
            ProdutoInicial.SetFocus

        Case 34060
            ProdutoFinal.SetFocus

        Case 34061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173268)

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
    If lErro <> SUCESSO Then Error 34067
      
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 34068

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 34069

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 34070
    
    sTotaliza = CStr(ComboTotaliza.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NTOTALIZA", sTotaliza)
    If lErro <> AD_BOOL_TRUE Then Error 34071
  
    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then Error 34072

    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then Error 34073
        
    lErro = objRelOpcoes.IncluirParametro("TFILIALINIC", FilialEmpresaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54660
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALFIM", FilialEmpresaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54661
    
    iAlmoxInicial = Codigo_Extrai(AlmoxarifadoInicial.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", CStr(iAlmoxInicial))
    If lErro <> AD_BOOL_TRUE Then Error 34074
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54662
    
    iAlmoxFinal = Codigo_Extrai(AlmoxarifadoFinal.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", CStr(iAlmoxFinal))
    If lErro <> AD_BOOL_TRUE Then Error 34075
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54663
    
    sProd_zerado = CStr(ProdutosZerados.Value)
    
    lErro = objRelOpcoes.IncluirParametro("NPRODZERADO", sProd_zerado)
    If lErro <> AD_BOOL_TRUE Then Error 34055
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    
        'Ler o mês e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
        lErro = CF("EstoqueMes_Le_Aberto",objEstoqueMes)
        If lErro <> SUCESSO And lErro <> 40673 Then Error 45134

        If lErro = 40673 Then Error 45135
 
        lErro = objRelOpcoes.IncluirParametro("NANO", objEstoqueMes.iAno)
        If lErro <> AD_BOOL_TRUE Then Error 45136
 
        lErro = objRelOpcoes.IncluirParametro("NMES", objEstoqueMes.iMes)
        If lErro <> AD_BOOL_TRUE Then Error 45137
    
        lErro = CF("EstoqueMes_Le_Apurado",objEstoqueMes)
        If lErro <> SUCESSO And lErro <> 46225 Then Error 45138
        
        If lErro = 46225 Then
            objEstoqueMes.iAno = 0
            objEstoqueMes.iMes = 0
        End If

        lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
        If lErro <> AD_BOOL_TRUE Then Error 45139
 
        lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
        If lErro <> AD_BOOL_TRUE Then Error 45140
    
    End If
    
    If ComboTotaliza.ListIndex = 0 Then gobjRelatorio.sNomeTsk = "SldEstAl"
    If ComboTotaliza.ListIndex = 1 Then gobjRelatorio.sNomeTsk = "SaldEst"
    If ComboTotaliza.ListIndex = 2 Then
        gobjRelatorio.sNomeTsk = "SaEstFil"
        If giFilialEmpresa <> EMPRESA_TODA Then Error 54506
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, iAlmoxInicial, iAlmoxFinal, sFilial_I, sFilial_F, sTotaliza, sProd_zerado)
    If lErro <> SUCESSO Then Error 34076

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err
        
        Case 54506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_E_EMPRESATODA", Err)
            
        Case 45135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAOEXISTE_MES_ABERTO", Err)
    
        Case 45134, 45136, 45137, 45138, 45139, 45140
        
        Case 34067, 34068, 34069, 34070, 34071, 34072, 34073, 34074, 34075, 34090
        
        Case 34076, 54660, 54661, 54662, 54663

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173269)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 34077

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 34078

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 34077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 34078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173270)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 34079

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 34079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173271)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 34080

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 34081

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 34082

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 34081

        Case 34082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173272)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34083
    
    If lErro <> SUCESSO Then Error 43281

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34083

         Case 43281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173273)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34084
    
    If lErro <> SUCESSO Then Error 43282

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34084

         Case 43282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173274)

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
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173275)

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
    
     If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(CInt(sFilial_I))

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))

    End If
    
''    If sTotaliza <> "" Then
''
''       If sExpressao <> "" Then sExpressao = sExpressao & " E "
''       sExpressao = sExpressao & "Totaliza = " & Forprint_ConvInt(CInt(sTotaliza))
''
''    End If
''
''    If sExpressao <> "" Then sExpressao = sExpressao & " E "
''    sExpressao = sExpressao & "NPRODZERADO = " & Forprint_ConvInt(CInt(sProdZerado))

    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If
    
          

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173276)

    End Select

    Exit Function

End Function


Private Sub FilialEmpresaInicial_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialEmpresaInicial_Validate

    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaInicial.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaInicial.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaInicial, iCodigo)
    If lErro <> SUCESSO Then Error 34086
    
    Exit Sub

Erro_FilialEmpresaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173277)

    End Select

    Exit Sub

End Sub


Private Sub FilialEmpresaFinal_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialEmpresaFinal_Validate

    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaFinal.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaFinal.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaFinal, iCodigo)
    If lErro <> SUCESSO Then Error 34087
    
    Exit Sub

Erro_FilialEmpresaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173278)

    End Select

    Exit Sub

End Sub


Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If iAlmoxarifadoInicialAlterado = 1 Then

        If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then

            'Tenta ler o Almoxarifado (NomeReduzido ou Código)
            lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
            If lErro <> SUCESSO Then Error 34088

        End If
        
        iAlmoxarifadoInicialAlterado = 0

    End If

    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34088
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173279)

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
            If lErro <> SUCESSO Then Error 34089

        End If
        
        iAlmoxarifadoFinalAlterado = 0

    End If

    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34089
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173280)

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

    If giFilialEmpresa = EMPRESA_TODA Then

        'seleciona totalizacao por Empresa
        ComboTotaliza.ListIndex = 2 'vide a property List da combo
        ComboTotaliza.Enabled = False
        
    Else
    
       'seleciona totalizacao por filial
        ComboTotaliza.ListIndex = 1 'vide a property List da combo
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    End If

''    If giFilialEmpresa <> EMPRESA_TODA Then
''
''       'seleciona giFilialEmpresa p/ FilialEmpresaInicial
''        For iIndice = 0 To FilialEmpresaInicial.ListCount - 1
''
''            If FilialEmpresaInicial.ItemData(iIndice) = giFilialEmpresa Then
''
''                FilialEmpresaInicial.ListIndex = iIndice
''                FilialEmpresaFinal.ListIndex = iIndice
''
''                Exit For
''
''            End If
''
''       Next
''
''    Else
''
''        FilialEmpresaInicial.ListIndex = -1
''        FilialEmpresaFinal.ListIndex = -1
''
''    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Saldo em Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpSaldoEstoque"
    
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
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

