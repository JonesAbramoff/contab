VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPrevVendaOcx 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   7545
   Begin VB.Frame Frame5 
      Caption         =   "Região de Venda"
      Height          =   1395
      Left            =   2805
      TabIndex        =   34
      Top             =   3855
      Width           =   2505
      Begin MSMask.MaskEdBox RegiaoInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   35
         Top             =   345
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   36
         Top             =   825
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelRegiaoFinal 
         Alignment       =   1  'Right Justify
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   885
         Width           =   360
      End
      Begin VB.Label LabelRegiaoInicial 
         Alignment       =   1  'Right Justify
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo"
      Height          =   1395
      Left            =   120
      TabIndex        =   29
      Top             =   3855
      Width           =   2505
      Begin MSMask.MaskEdBox TipoInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   30
         Top             =   345
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   31
         Top             =   825
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoInicial 
         Alignment       =   1  'Right Justify
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   405
         Width           =   315
      End
      Begin VB.Label LabelTipoFinal 
         Alignment       =   1  'Right Justify
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   885
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data do Período de Venda"
      Height          =   750
      Left            =   120
      TabIndex        =   22
      Top             =   1515
      Width           =   5055
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   330
         Left            =   1725
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicialVenda 
         Height          =   315
         Left            =   750
         TabIndex        =   24
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   330
         Left            =   4095
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinalVenda 
         Height          =   330
         Left            =   3120
         TabIndex        =   26
         Top             =   255
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFimVenda 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2715
         TabIndex        =   28
         Top             =   300
         Width           =   570
      End
      Begin VB.Label dIniVenda 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   345
         TabIndex        =   27
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrevVendaOcx.ctx":0000
      Left            =   900
      List            =   "RelOpPrevVendaOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   270
      Width           =   2910
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
      Left            =   5550
      Picture         =   "RelOpPrevVendaOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   855
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data da Previsão"
      Height          =   750
      Left            =   120
      TabIndex        =   12
      Top             =   675
      Width           =   5055
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1725
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicialPrev 
         Height          =   315
         Left            =   750
         TabIndex        =   14
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4095
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinalPrev 
         Height          =   330
         Left            =   3120
         TabIndex        =   16
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIniPrev 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   390
      End
      Begin VB.Label dFimPrev 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2715
         TabIndex        =   17
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1395
      Left            =   120
      TabIndex        =   5
      Top             =   2295
      Width           =   5055
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   510
         TabIndex        =   6
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   495
         TabIndex        =   7
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2055
         TabIndex        =   11
         Top             =   840
         Width           =   2730
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2055
         TabIndex        =   10
         Top             =   360
         Width           =   2730
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   390
         Width           =   360
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   840
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5280
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrevVendaOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrevVendaOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrevVendaOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrevVendaOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   21
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPrevVendaOcx"
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

Private WithEvents objEventoTipoInicial As AdmEvento
Attribute objEventoTipoInicial.VB_VarHelpID = -1
Private WithEvents objEventoTipoFinal As AdmEvento
Attribute objEventoTipoFinal.VB_VarHelpID = -1

Private Sub TipoInicial_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoInicial_Validate
    
    If Len(Trim(TipoInicial.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoInicial.Text))
        If lErro <> SUCESSO Then Error 59555
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoInicial.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then Error 59556
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then Error 59557
        
        TipoInicial.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 59555, 59556

        Case 59557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171493)

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
        If lErro <> SUCESSO Then Error 59558
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoFinal.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then Error 59559
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then Error 59560
        
        TipoFinal.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 59558, 59559

        Case 59560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171494)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_RegiaoInicial_Validate
    
    If Len(Trim(RegiaoInicial.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(RegiaoInicial.Text))
        If lErro <> SUCESSO Then Error 59555
    
        objRegiaoVenda.iCodigo = CInt(Codigo_Extrai(RegiaoInicial.Text))
    
        'Lê a Região de Venda
        lErro = CF("RegiaoVenda_Le",objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 59556
        
        'Se não encontrar --> Erro
        If lErro = 16137 Then Error 59557
        
        RegiaoInicial.Text = objRegiaoVenda.iCodigo & SEPARADOR & objRegiaoVenda.sDescricao
    
    End If
    
    Exit Sub

Erro_RegiaoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 59555, 59556

        Case 59557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, objRegiaoVenda.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171495)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_RegiaoFinal_Validate
    
    If Len(Trim(RegiaoFinal.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(RegiaoFinal.Text))
        If lErro <> SUCESSO Then Error 59558
    
        objRegiaoVenda.iCodigo = CInt(Codigo_Extrai(RegiaoFinal.Text))
    
        'Lê a Região de Venda
        lErro = CF("RegiaoVenda_Le",objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 59559
        
        'Se não encontrar --> Erro
        If lErro = 16137 Then Error 59560
        
        RegiaoFinal.Text = objRegiaoVenda.iCodigo & SEPARADOR & objRegiaoVenda.sDescricao
    
    End If
    
    Exit Sub

Erro_RegiaoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 59558, 59559

        Case 59560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, objRegiaoVenda.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171496)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iOpcao As Integer

On Error GoTo Erro_Form_Load
               
    Set objEventoTipoInicial = New AdmEvento
    Set objEventoTipoFinal = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoInicial)
    If lErro <> SUCESSO Then Error 47271

    lErro = CF("Inicializa_Mascara_Produto_MaskEd",ProdutoFinal)
    If lErro <> SUCESSO Then Error 47272

    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 48561
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 47271, 47272, 48560, 48561

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171497)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar()
    If lErro <> SUCESSO Then Error 47274
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then Error 47275

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 47276

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then Error 47277

    lErro = CF("Traz_Produto_MaskEd",sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 47278
        
    'pega data de previsão inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINICPREV", sParam)
    If lErro <> SUCESSO Then Error 47279

    Call DateParaMasked(DataInicialPrev, CDate(sParam))
    
    'pega data de previsão final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIMPREV", sParam)
    If lErro <> SUCESSO Then Error 47280

    Call DateParaMasked(DataFinalPrev, CDate(sParam))
    
    'pega data de venda inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINICVENDA", sParam)
    If lErro <> SUCESSO Then Error 47279

    Call DateParaMasked(DataInicialVenda, CDate(sParam))
    
    'pega data de venda final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIMVENDA", sParam)
    If lErro <> SUCESSO Then Error 47280

    Call DateParaMasked(DataFinalVenda, CDate(sParam))
                      
    'pega parâmetro tipo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODINI", sParam)
    If lErro Then Error 54600
    
    TipoInicial.Text = sParam
    
    'pega parâmetro tipo final e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODFIM", sParam)
    If lErro Then Error 54601
    
    TipoFinal.Text = sParam
                      
    'pega parâmetro Região de Venda inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOREGIAOINI", sParam)
    If lErro Then Error 54600
    
    RegiaoInicial.Text = sParam
    
    'pega parâmetro Região de Venda final e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOREGIAOFIM", sParam)
    If lErro Then Error 54601
    
    RegiaoFinal.Text = sParam
                      
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47274, 47275, 47276, 47277, 47278, 47279, 47280, 47281, 48555, 48556

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171498)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47270

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47270
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171499)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoTipoInicial = Nothing
    Set objEventoTipoFinal = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47283
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 48562
        
    TipoInicial.Text = ""
    TipoFinal.Text = ""
    RegiaoInicial.Text = ""
    RegiaoFinal.Text = ""
          
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47283, 48562
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171500)

    End Select

    Exit Sub

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
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 47286
    
    lErro = objRelOpcoes.Limpar()
    If lErro <> AD_BOOL_TRUE Then Error 47287
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 47288

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 47289
                
    If Trim(DataInicialPrev.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINICPREV", DataInicialPrev.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINICPREV", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47290
    
    If Trim(DataFinalPrev.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIMPREV", DataFinalPrev.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIMPREV", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47291
          
    If Trim(DataInicialVenda.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINICVENDA", DataInicialVenda.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINICVENDA", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47290
    
    If Trim(DataFinalVenda.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIMVENDA", DataFinalVenda.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIMVENDA", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47291
          
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODINI", TipoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54633
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODFIM", TipoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54634
                    
    lErro = objRelOpcoes.IncluirParametro("TTIPOREGIAOINI", RegiaoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54633
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOREGIAOFIM", RegiaoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54634
                    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then Error 47293
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 47286, 47287, 47288, 47289, 47290, 47291, 47292, 47293, 48557, 48558, 54844, 54845

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171501)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click
'????? Verificar se ao digitar na combo o nome de uma opcao, pode exclui-la.
    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47294

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47295

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47296
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 48563
            
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47295, 47296, 48563

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171502)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47297

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47297

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171503)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47298

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47299

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47300

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47301
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47299, 47300, 47301

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171504)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If
    
    If Trim(DataInicialPrev.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataPrevisao >= " & Forprint_ConvData(CDate(DataInicialPrev.Text))

    End If
    
    If Trim(DataFinalPrev.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataPrevisao <= " & Forprint_ConvData(CDate(DataFinalPrev.Text))

    End If
           
    If Trim(DataInicialVenda.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVendaIni >= " & Forprint_ConvData(CDate(DataInicialVenda.Text))

    End If
    
    If Trim(DataFinalVenda.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVendaFim <= " & Forprint_ConvData(CDate(DataFinalVenda.Text))

    End If
           
    If TipoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto  >= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoInicial.Text)))

    End If
        
    If TipoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto <= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoFinal.Text)))

    End If
        
    If RegiaoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda  >= " & Forprint_ConvInt(CInt(Codigo_Extrai(RegiaoInicial.Text)))

    End If
        
    If RegiaoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(CInt(Codigo_Extrai(RegiaoFinal.Text)))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171505)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata",ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 47303

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata",ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 47304

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 47305

    End If
        
    'data inicial de previsão não pode ser maior que a data final de previsão
    If Trim(DataInicialPrev.ClipText) <> "" And Trim(DataFinalPrev.ClipText) <> "" Then
    
         If CDate(DataInicialPrev.Text) > CDate(DataFinalPrev.Text) Then Error 47306
    
    End If
            
    'data inicial de venda não pode ser maior que a data final de venda
    If Trim(DataInicialVenda.ClipText) <> "" And Trim(DataFinalVenda.ClipText) <> "" Then
    
         If CDate(DataInicialVenda.Text) > CDate(DataFinalVenda.Text) Then Error 47306
    
    End If
    
    'tipo inicial não pode ser maior que o tipo final
    If Trim(TipoInicial.Text) <> "" And Trim(TipoFinal.Text) <> "" Then
    
         If TipoInicial.Text > TipoFinal.Text Then Error 54610
         
    End If

    'Região inicial não pode ser maior que a Região final
    If Trim(RegiaoInicial.Text) <> "" And Trim(RegiaoFinal.Text) <> "" Then
    
         If RegiaoInicial.Text > RegiaoFinal.Text Then Error 54610
         
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                     
        Case 47303
            ProdutoInicial.SetFocus

        Case 47304
            ProdutoFinal.SetFocus

        Case 47305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus
            
        Case 47306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicialPrev.SetFocus
            
        Case 47306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicialVenda.SetFocus
            
        Case 54610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", Err)
            TipoInicial.SetFocus
        
        Case 54610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_INICIAL_MAIOR", Err)
            RegiaoInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171506)

    End Select

    Exit Function

End Function

Private Sub DataFinalPrev_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinalPrev)

End Sub

Private Sub DataFinalPrev_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinalPrev_Validate

    If Len(DataFinalPrev.ClipText) > 0 Then

        sDataFim = DataFinalPrev.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 47307

    End If

    Exit Sub

Erro_DataFinalPrev_Validate:

    Cancel = True


    Select Case Err

        Case 47307

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171507)

    End Select

    Exit Sub

End Sub

Private Sub DataInicialPrev_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicialPrev)

End Sub

Private Sub DataInicialPrev_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicialPrev_Validate

    If Len(DataInicialPrev.ClipText) > 0 Then

        sDataInic = DataInicialPrev.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 47308

    End If

    Exit Sub

Erro_DataInicialPrev_Validate:

    Cancel = True


    Select Case Err

        Case 47308

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171508)

    End Select

    Exit Sub

End Sub

Private Sub DataFinalVenda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinalVenda)

End Sub

Private Sub DataFinalVenda_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinalVenda_Validate

    If Len(DataFinalVenda.ClipText) > 0 Then

        sDataFim = DataFinalVenda.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 47307

    End If

    Exit Sub

Erro_DataFinalVenda_Validate:

    Cancel = True


    Select Case Err

        Case 47307

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171509)

    End Select

    Exit Sub

End Sub

Private Sub DataInicialVenda_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicialVenda)

End Sub

Private Sub DataInicialVenda_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicialVenda_Validate

    If Len(DataInicialVenda.ClipText) > 0 Then

        sDataInic = DataInicialVenda.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 47308

    End If

    Exit Sub

Erro_DataInicialVenda_Validate:

    Cancel = True


    Select Case Err

        Case 47308

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171510)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicialPrev, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47309

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 47309
            DataInicialPrev.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171511)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicialPrev, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47310

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 47310
            DataInicialPrev.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171512)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinalPrev, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47311

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 47311
            DataFinalPrev.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171513)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinalPrev, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47312

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 47312
            DataFinalPrev.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171514)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_DownClick

    lErro = Data_Up_Down_Click(DataInicialVenda, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47309

    Exit Sub

Erro_UpDown3_DownClick:

    Select Case Err

        Case 47309
            DataInicialVenda.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171515)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_UpClick

    lErro = Data_Up_Down_Click(DataInicialVenda, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47310

    Exit Sub

Erro_UpDown3_UpClick:

    Select Case Err

        Case 47310
            DataInicialVenda.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171516)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_DownClick

    lErro = Data_Up_Down_Click(DataFinalVenda, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47311

    Exit Sub

Erro_UpDown4_DownClick:

    Select Case Err

        Case 47311
            DataFinalVenda.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171517)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_UpClick

    lErro = Data_Up_Down_Click(DataFinalVenda, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47312

    Exit Sub

Erro_UpDown4_UpClick:

    Select Case Err

        Case 47312
            DataFinalVenda.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171518)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco",ProdutoFinal, DescProdFim)
     If lErro <> SUCESSO And lErro <> 27095 Then Error 47316
    
    If lErro = 27095 Then Error 47319

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47316
            
        Case 47319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171519)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco",ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 47317
    
    If lErro = 27095 Then Error 47318
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47317

        Case 47318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171520)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Define_Padrao
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171521)

    End Select

    Exit Function

End Function

Private Sub LabelTipoInicial_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelTipoInicial_Click

    If Len(Trim(TipoInicial.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = TipoInicial.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoInicial)

    Exit Sub

Erro_LabelTipoInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171522)

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
        objTipoProduto.iTipo = TipoFinal.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoFinal)

    Exit Sub

Erro_LabelTipoFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171523)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171524)

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

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171525)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PREVISAO_VENDAS
    Set Form_Load_Ocx = Me
    Caption = "Relação de Previsão de Vendas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrevVenda"
    
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

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82634

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82635
    
    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82636

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82634, 82636

        Case 82635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171526)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82637

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82638

    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82639

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82637, 82639

        Case 82638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171527)

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
        If lErro <> SUCESSO Then gError 82640

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82640

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171528)

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
        If lErro <> SUCESSO Then gError 82641

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82641

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171529)

    End Select

    Exit Sub

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

Private Sub LabelRegiaoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoFinal, Source, X, Y)
End Sub

Private Sub LabelRegiaoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoInicial, Source, X, Y)
End Sub

Private Sub LabelRegiaoInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoInicial, Source, X, Y)
End Sub

Private Sub LabelTipoInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoFinal, Source, X, Y)
End Sub

Private Sub LabelTipoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoFinal, Button, Shift, X, Y)
End Sub

Private Sub dFimVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFimVenda, Source, X, Y)
End Sub

Private Sub dFimVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFimVenda, Button, Shift, X, Y)
End Sub

Private Sub dIniVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIniVenda, Source, X, Y)
End Sub

Private Sub dIniVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIniVenda, Button, Shift, X, Y)
End Sub

Private Sub dIniPrev_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIniPrev, Source, X, Y)
End Sub

Private Sub dIniPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIniPrev, Button, Shift, X, Y)
End Sub

Private Sub dFimPrev_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFimPrev, Source, X, Y)
End Sub

Private Sub dFimPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFimPrev, Button, Shift, X, Y)
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

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

