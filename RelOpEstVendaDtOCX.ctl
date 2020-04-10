VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpEstVendaDtOCX 
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9030
   Begin VB.CheckBox ExibirCusto 
      Caption         =   "Exibir Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3270
      TabIndex        =   3
      Top             =   915
      Width           =   1590
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   3960
      ItemData        =   "RelOpEstVendaDtOCX.ctx":0000
      Left            =   6270
      List            =   "RelOpEstVendaDtOCX.ctx":0002
      TabIndex        =   16
      Top             =   1350
      Width           =   2535
   End
   Begin VB.ComboBox TabelaPrecos 
      Height          =   315
      Left            =   1770
      TabIndex        =   4
      Top             =   1335
      Width           =   3060
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Produto"
      Height          =   870
      Left            =   195
      TabIndex        =   25
      Top             =   2775
      Width           =   5850
      Begin MSMask.MaskEdBox TipoFinal 
         Height          =   315
         Left            =   3495
         TabIndex        =   8
         Top             =   345
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoInicial 
         Height          =   315
         Left            =   660
         TabIndex        =   7
         Top             =   330
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoInicial 
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
         TabIndex        =   30
         Top             =   375
         Width           =   315
      End
      Begin VB.Label LabelTipoFinal 
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
         Height          =   225
         Left            =   3030
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   390
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6270
      ScaleHeight     =   495
      ScaleWidth      =   2115
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   315
      Width           =   2175
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpEstVendaDtOCX.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpEstVendaDtOCX.ctx":018E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpEstVendaDtOCX.ctx":06C0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpEstVendaDtOCX.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
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
      Left            =   4185
      Picture         =   "RelOpEstVendaDtOCX.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   180
      Width           =   1515
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpEstVendaDtOCX.ctx":0A9A
      Left            =   900
      List            =   "RelOpEstVendaDtOCX.ctx":0A9C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   323
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1515
      Left            =   225
      TabIndex        =   18
      Top             =   3795
      Width           =   5775
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   765
         TabIndex        =   10
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
         Left            =   765
         TabIndex        =   9
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   900
         Width           =   360
      End
      Begin VB.Label LabelProdutoDe 
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   375
         Width           =   315
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2310
         TabIndex        =   20
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   19
         Top             =   375
         Width           =   3135
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2775
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInv 
      Height          =   300
      Left            =   1770
      TabIndex        =   1
      Top             =   855
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Almoxarifado 
      Height          =   315
      Left            =   1755
      TabIndex        =   5
      Top             =   1830
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1740
      TabIndex        =   6
      Top             =   2325
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
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
      Height          =   240
      Left            =   6285
      TabIndex        =   31
      Top             =   1110
      Width           =   1185
   End
   Begin VB.Label LabelFornecedor 
      Caption         =   "Fornecedor:"
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
      Height          =   255
      Left            =   645
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   2355
      Width           =   1050
   End
   Begin VB.Label lblAlmoxarifado 
      Caption         =   "Almoxarifado:"
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
      Height          =   255
      Left            =   525
      TabIndex        =   28
      Top             =   1860
      Width           =   1185
   End
   Begin VB.Label LabelTab 
      AutoSize        =   -1  'True
      Caption         =   "Tabela de Preços:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   1380
      Width           =   1575
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
      Left            =   210
      TabIndex        =   24
      Top             =   353
      Width           =   555
   End
   Begin VB.Label Data 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1230
      TabIndex        =   17
      Top             =   915
      Width           =   480
   End
End
Attribute VB_Name = "RelOpEstVendaDtOCX"
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
Private WithEvents objEventoTipoInicial As AdmEvento
Attribute objEventoTipoInicial.VB_VarHelpID = -1
Private WithEvents objEventoTipoFinal As AdmEvento
Attribute objEventoTipoFinal.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoFornecedor = New AdmEvento
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    Set objEventoTipoInicial = New AdmEvento
    Set objEventoTipoFinal = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 85202 '85154

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 85207 '85155
    
    'mostra na tela a data de dia atual
    DataInv.PromptInclude = False
    DataInv.Text = Format(Date, "dd/mm/yy")
    DataInv.PromptInclude = True
    
    lErro = Carrega_TabelaPrecos()
    If lErro <> SUCESSO Then gError 85207
    
    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then gError 85207
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 85202, 85207
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168576)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 85211 '85156
   
    'pega parâmetro e exibe
    lErro = objRelOpcoes.ObterParametro("TTABPRECO", sParam)
    If lErro Then gError 85212
    
    TabelaPrecos.Text = sParam
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 85212 '85157

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85216 '85158

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 85218 '85159

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85219 '85160
    
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODINI", sParam)
    If lErro <> SUCESSO Then gError 85220 '85200
    
    TipoInicial.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 85241 '85201
    
    TipoFinal.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then gError 85242 '85220

    Call DateParaMasked(DataInv, CDate(sParam))
    
    'pega parâmetro e exibe
    lErro = objRelOpcoes.ObterParametro("TFORNECEDOR", sParam)
    If lErro Then gError 85242
    
    Fornecedor.Text = sParam
    
    'pega parâmetro Almoxarifado e exibe
    lErro = objRelOpcoes.ObterParametro("NALMOX", sParam)
    If lErro Then gError 85242
    
    If sParam = "0" Then sParam = ""
    Almoxarifado.Text = sParam
    Call Almoxarifado_Validate(bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("NEXIBIRCUSTO", sParam)
    If lErro Then gError 189309
    
    If StrParaInt(sParam) = MARCADO Then
        ExibirCusto.Value = vbChecked
    Else
        ExibirCusto.Value = vbUnchecked
    End If
    
    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 85211, 85212, 85216, 85218, 85219, 85220, 85241, 85242, 189309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168577)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 85243 '85161
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 85244 '85162
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 85244
        
        Case 85243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168578)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'verifica se a data foi preenchida,se não erro
    If Len(DataInv.ClipText) = 0 Then gError 85245 '85245
    
    'Verifica se o campo está preenchid0
    If Len(Trim(TabelaPrecos.Text)) = 0 Then gError 184208

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 85246  '85163

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 85247 '85164

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 85249 '85165

    End If
    
    'tipo inicial não pode ser maior que o tipo final
    If Trim(TipoInicial.Text) <> "" And Trim(TipoFinal.Text) <> "" Then
    
         If TipoInicial.Text > TipoFinal.Text Then gError 85250 '85197
         
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 85246
            ProdutoInicial.SetFocus

        Case 85247
            ProdutoFinal.SetFocus

        Case 85249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 85250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", gErr)
                        
        Case 85245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 184208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168579)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 85251 '85166
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    TipoInicial.Text = ""
    TipoFinal.Text = ""
    
    TabelaPrecos.Text = ""
    
    ExibirCusto.Value = vbChecked

    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 85251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168580)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoFornecedor = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub


Private Sub DataInv_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInv_Validate

    'Se a data informada for maior que a data atual erro.
    If StrParaDate(DataInv.Text) > gdtDataAtual Then gError 85252
    
    Exit Sub

Erro_DataInv_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 85252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_MAIOR_DATA_HOJE", gErr, DataInv.Text, Date)

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
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 85253 '85167#

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 85253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168581)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto


On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85254 '85168

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85255 '85169

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 85256 '85170

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 85254, 85256

        Case 85255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168582)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto


On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 85256 '85171

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 85257 '85172

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 85258 '85173

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 85256, 85258

        Case 85257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168583)

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
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 85259 '85174

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 85259

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168584)

    End Select

    Exit Sub
    
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String, iTabPreco As Integer
Dim objRelEstVendaDt As New ClassRelEstVendaDt
Dim sAlmox As String
Dim iAlmox As Integer, lFornecedor As Long
Dim iExibirCusto As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85260 '85175
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 85261 '85176
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 85262 '85177

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 85263 '85178
             
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODINI", TipoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85264 '85198
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODFIM", TipoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85265 '85199
    
    lErro = objRelOpcoes.IncluirParametro("DDATA", DataInv.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85266 '85259
    
    iTabPreco = Codigo_Extrai(TabelaPrecos.Text)
    
    lErro = objRelOpcoes.IncluirParametro("NTABPRECO", CStr(iTabPreco))
    If lErro <> AD_BOOL_TRUE Then gError 85266
    
    lErro = objRelOpcoes.IncluirParametro("TTABPRECO", TabelaPrecos.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85266
    
    lFornecedor = LCodigo_Extrai(Fornecedor.Text)
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDOR", CStr(lFornecedor))
    If lErro <> AD_BOOL_TRUE Then gError 85266
    
    lErro = objRelOpcoes.IncluirParametro("TFORNECEDOR", Fornecedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85266
        
    iAlmox = Codigo_Extrai(Almoxarifado.Text)

    lErro = objRelOpcoes.IncluirParametro("NALMOX", CStr(iAlmox))
    If lErro <> AD_BOOL_TRUE Then gError 85266
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXARIFADO", Almoxarifado.Text)
    If lErro <> AD_BOOL_TRUE Then gError 85266
    
    If ExibirCusto.Value = vbChecked Then
        iExibirCusto = MARCADO
    Else
        iExibirCusto = DESMARCADO
    End If
        
    lErro = objRelOpcoes.IncluirParametro("NEXIBIRCUSTO", CStr(iExibirCusto))
    If lErro <> AD_BOOL_TRUE Then gError 189310
        
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 85267 '85179

    If bExecutar Then
    
        With objRelEstVendaDt
            .iFilialEmpresa = giFilialEmpresa
            .dtData1 = MaskedParaDate(DataInv)
            .sProdInicial = sProd_I
            .sProdFinal = sProd_F
            .lFornecedor = lFornecedor
            .iTabelaPreco = iTabPreco
            .iAlmoxarifado = iAlmox
            .iTipoProdutoInicial = Codigo_Extrai(TipoInicial.Text)
            .iTipoProdutoFinal = Codigo_Extrai(TipoFinal.Text)
        End With
        
        GL_objMDIForm.MousePointer = vbHourglass
        lErro = CF("RelEstVendaDt_Prepara", objRelEstVendaDt)
        GL_objMDIForm.MousePointer = vbDefault
        If lErro <> SUCESSO Then gError 184206
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(objRelEstVendaDt.lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 184207
    
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 85260 To 85267, 184206, 184207, 189310
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168585)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 85268 '85180

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 85269 '85181

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 85270 '85182
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 85268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 85269, 85270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168586)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 85271 '85182
    
    If ExibirCusto.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "ESTVENDC"
    Else
        gobjRelatorio.sNomeTsk = "ESTVENDT"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 85271

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168587)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 85272 '85183

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 85273 '85184

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 84308 '85185
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 84180 '85186
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 85272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 85273, 84308, 84180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168588)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 84433 '85187
    
    If lErro <> SUCESSO Then gError 84434 '85188

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 84433

        Case 84434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168589)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 84435 '85189
    
    If lErro <> SUCESSO Then gError 84436 '85190
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 84435
        
        Case 84436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168590)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168591)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ANALISE_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Estoque pelo Preço de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEstVendaDtOCX"
    
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
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call LabelFornecedor_Click
        End If
                
    End If

End Sub

Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
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

Private Sub TipoInicial_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoInicial_Validate
    
    If Len(Trim(TipoInicial.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoInicial.Text))
        If lErro <> SUCESSO Then gError 84439 '85191
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoInicial.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 84437 '85192
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 84438 '85193
        
        TipoInicial.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 84439, 84437

        Case 84438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168592)

    End Select

    Exit Sub

End Sub

Private Sub TipoFinal_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoFinal_Validate
    
    If Len(Trim(TipoFinal.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoFinal.Text))
        If lErro <> SUCESSO Then gError 84440 '85194
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoFinal.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 84441 '85195
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then gError 84442 '85196
        
        TipoFinal.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 84441, 84440

        Case 84442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168593)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoInicial_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelTipoInicial_Click

    If Len(Trim(TipoInicial.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = Codigo_Extrai(TipoInicial.Text)

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoInicial)

    Exit Sub

Erro_LabelTipoInicial_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168594)

    End Select

    Exit Sub

End Sub
Private Sub objEventoTipoInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoInicial_evSelecao

    Set objTipoProduto = obj1

    TipoInicial.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoInicial_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168595)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoFinal_evSelecao

    Set objTipoProduto = obj1

    TipoFinal.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoFinal_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168596)

    End Select

    Exit Sub

End Sub


Private Sub LabelTipoFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objTipoProduto As ClassTipoDeProduto

On Error GoTo Erro_LabelTipoFinal_Click

    If Len(Trim(TipoFinal.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = Codigo_Extrai(TipoFinal.Text)

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoFinal)

    Exit Sub

Erro_LabelTipoFinal_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168597)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

   
    lErro = Data_Up_Down_Click(DataInv, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84443 '85272

    
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 84443
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168598)

    End Select

    Exit Sub

End Sub
Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInv, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84444 '85273

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 84444
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168599)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TabelaPrecos() As Long
'Carrega a Combo TabelaPrecos

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TabelaPrecos

    'lê códigos e descrições da tabela TabelasDePrecos e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 37131
    
    'preenche a combo
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            TabelaPrecos.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TabelaPrecos.ItemData(TabelaPrecos.NewIndex) = objCodigoNome.iCodigo
             
        End If
    
    Next

    Carrega_TabelaPrecos = SUCESSO

    Exit Function

Erro_Carrega_TabelaPrecos:

    Carrega_TabelaPrecos = Err

    Select Case Err

        'Erro já tratado
        Case 37131

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171394)

    End Select

    Exit Function

End Function

Private Sub TabelaPrecos_Validate(Cancel As Boolean)
'Busca a descricao com código digitado na Combo

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPrecos_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TabelaPrecos.ListIndex <> -1 Then Exit Sub

    If Len(Trim(TabelaPrecos.Text)) = 0 Then Exit Sub

    lErro = Combo_Seleciona(TabelaPrecos, iCodigo)
    If lErro <> SUCESSO Then Error 37132

    Exit Sub

Erro_TabelaPrecos_Validate:

    Cancel = True


    Select Case Err

        Case 37132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", Err, iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171395)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCria As Integer

On Error GoTo Erro_Fornecedor_Validate

        If Len(Trim(Fornecedor.Text)) > 0 Then

            iCria = 0 'Não deseja criar Fornecedor caso não exista
            lErro = TP_Fornecedor_Le2(Fornecedor, objFornecedor, iCria)
            If lErro <> SUCESSO Then gError 87586

        End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 87586
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169375)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(Fornecedor.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    'Preenche campo Cliente
    Fornecedor.Text = CStr(objFornecedor.lCodigo)
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub
    
End Sub

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 37195

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

        Case 37195

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169556)

    End Select

    Exit Function

End Function

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim sListBoxItem As String
Dim lErro As Long

On Error GoTo Erro_Almoxarifados_DblClick

   'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
    
    Almoxarifado.Text = sListBoxItem

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169557)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(Almoxarifado.Text)) > 0 Then
        
        'Le o almoxarifado pelo código ou pelo nome reduzido e joga o nome reduzido em Almoxarifado.Text
        lErro = TP_Almoxarifado_Le_ComCodigo(Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 37196
                
        Almoxarifado = objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        
        'se o almoxarifado não pertencer a filial em questão ==> erro
        If objAlmoxarifado.iFilialEmpresa <> giFilialEmpresa Then gError 93740
        
    End If
    
    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True

    Select Case gErr

        Case 37196

        Case 93740
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIAL_DIFERENTE", gErr, objAlmoxarifado.iCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 169559)

    End Select

    Exit Sub

End Sub

Private Sub LblAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(lblAlmoxarifado, Source, X, Y)
End Sub

Private Sub LblAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblAlmoxarifado, Button, Shift, X, Y)
End Sub


