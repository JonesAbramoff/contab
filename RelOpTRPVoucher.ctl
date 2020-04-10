VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTRPVouchers 
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   8070
   Begin VB.Frame Frame6 
      Caption         =   "Representante"
      Height          =   810
      Left            =   105
      TabIndex        =   46
      Top             =   3345
      Width           =   5625
      Begin VB.CheckBox ReprExibeVendas 
         Caption         =   "Exibe as próprias vendas"
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
         Left            =   2730
         TabIndex        =   12
         Top             =   405
         Width           =   2790
      End
      Begin MSMask.MaskEdBox ReprInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   11
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelReprDe 
         AutoSize        =   -1  'True
         Caption         =   "Repr:"
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
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   47
         Top             =   405
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data de Vigência dos Voucher"
      Height          =   810
      Left            =   120
      TabIndex        =   43
      Top             =   1650
      Width           =   5625
      Begin MSComCtl2.UpDown UpDownVigDe 
         Height          =   300
         Left            =   1590
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVigInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   5
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownVigAte 
         Height          =   315
         Left            =   4170
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVigFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   7
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   2835
         TabIndex        =   45
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo"
      Height          =   1065
      Left            =   105
      TabIndex        =   41
      Top             =   6360
      Width           =   5625
      Begin VB.OptionButton Resumido 
         Caption         =   "Resumido"
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
         Left            =   3075
         TabIndex        =   19
         Top             =   690
         Width           =   1890
      End
      Begin VB.OptionButton Detalhado 
         Caption         =   "Detalhado"
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
         Left            =   1125
         TabIndex        =   18
         Top             =   675
         Width           =   1890
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   4545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Index           =   5
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   330
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clientes"
      Height          =   810
      Left            =   120
      TabIndex        =   38
      Top             =   2490
      Width           =   5625
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   9
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   405
         Width           =   315
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5820
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTRPVoucher.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTRPVoucher.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTRPVoucher.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTRPVoucher.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data de Emissão dos Voucher"
      Height          =   810
      Left            =   120
      TabIndex        =   34
      Top             =   780
      Width           =   5625
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1590
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   4170
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   2835
         TabIndex        =   35
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1320
      Left            =   105
      TabIndex        =   29
      Top             =   4200
      Width           =   5625
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   630
         TabIndex        =   13
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
         Left            =   630
         TabIndex        =   14
         Top             =   840
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   855
         Width           =   435
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   390
         Width           =   360
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2145
         TabIndex        =   31
         Top             =   360
         Width           =   3405
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2145
         TabIndex        =   30
         Top             =   840
         Width           =   3405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Executivo de Contas"
      Height          =   795
      Left            =   105
      TabIndex        =   26
      Top             =   5520
      Width           =   5625
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   3285
         TabIndex        =   16
         Top             =   345
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorAte 
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
         Left            =   2880
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   405
         Width           =   360
      End
      Begin VB.Label LabelVendedorDe 
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTRPVoucher.ctx":0994
      Left            =   1275
      List            =   "RelOpTRPVoucher.ctx":0996
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
      Left            =   6000
      Picture         =   "RelOpTRPVoucher.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   870
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   37
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTRPVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const TODOS_OS_NIVEIS = 0
Const UM_NIVEL = 1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim giClienteInicial As Integer
Dim giReprInicial As Integer
Dim giProdInicial As Integer
Dim giVendedorInicial As Integer

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoRepr As AdmEvento
Attribute objEventoRepr.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoRepr = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 37932

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 37933
    
    Tipo.AddItem "Por Produto"
    Tipo.ItemData(Tipo.NewIndex) = 1
    
    Tipo.AddItem "Por Cliente"
    Tipo.ItemData(Tipo.NewIndex) = 2
    
    Tipo.AddItem "Por Executivo de Contas"
    Tipo.ItemData(Tipo.NewIndex) = 3
    
    Tipo.AddItem "Por Filial"
    Tipo.ItemData(Tipo.NewIndex) = 4
    
    Tipo.AddItem "Por Destino"
    Tipo.ItemData(Tipo.NewIndex) = 5
    
    Tipo.AddItem "Por Filial\Executivo de Contas"
    Tipo.ItemData(Tipo.NewIndex) = 6
    
    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 37932, 37933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171092)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    giVendedorInicial = 1
    
    giProdInicial = 1
    
    giClienteInicial = 1
    
    giReprInicial = 1
    
    Detalhado.Value = True
    
    Tipo.ListIndex = 0
    
    ReprExibeVendas.Value = vbUnchecked
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171093)

    End Select

    Exit Function

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteInicial.Text = CStr(objCliente.lCodigo)
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = CStr(objCliente.lCodigo)
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRepr_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giReprInicial = 1 Then
        ReprInicial.Text = CStr(objCliente.lCodigo)
        Call ReprInicial_Validate(bSGECancelDummy)
'    Else
'        ReprFinal.Text = CStr(objCliente.lCodigo)
'        Call ReprFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 37936
   
    'pega Vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TVENDDE", sParam)
    If lErro <> SUCESSO Then gError 37937
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  Vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TVENDATE", sParam)
    If lErro Then gError 37938
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODDE", sParam)
    If lErro <> SUCESSO Then gError 37939

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 37940

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODATE", sParam)
    If lErro <> SUCESSO Then gError 37941

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 37942
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 37943

    Call DateParaMasked(DataInicial, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 37944

    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVIGDE", sParam)
    If lErro <> SUCESSO Then gError 37943

    Call DateParaMasked(DataVigInicial, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DVIGATE", sParam)
    If lErro <> SUCESSO Then gError 37944

    Call DateParaMasked(DataVigFinal, StrParaDate(sParam))
    
     'pega parâmetro Tipo
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError 37945
   
    Tipo.ListIndex = StrParaInt(sParam)
    
    'pega parâmetro detalhamento
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NDET", sParam)
    If lErro <> SUCESSO Then gError 37946
   
    If StrParaInt(sParam) = MARCADO Then
        Detalhado.Value = True
    Else
        Resumido.Value = True
    End If
    
    'pega parâmetro detalhamento
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NEXIBE", sParam)
    If lErro <> SUCESSO Then gError 37946
   
    If StrParaInt(sParam) = MARCADO Then
        ReprExibeVendas.Value = vbChecked
    Else
        ReprExibeVendas.Value = vbUnchecked
    End If
    
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCLIDE", sParam)
    If lErro <> SUCESSO Then gError 20848
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("TCLIATE", sParam)
    If lErro <> SUCESSO Then gError 20849
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TREPR", sParam)
    If lErro <> SUCESSO Then gError 20848
    
    ReprInicial.Text = sParam
    Call ReprInicial_Validate(bSGECancelDummy)
    
'    'pega  Cliente final e exibe
'    lErro = objRelOpcoes.ObterParametro("TREPRATE", sParam)
'    If lErro <> SUCESSO Then gError 20849
'
'    ReprFinal.Text = sParam
'    Call ReprFinal_Validate(bSGECancelDummy)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 20848, 20849, 37936 To 37946

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171094)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 37930

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 37930
        
        Case 29884
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171095)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoVendedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCliente = Nothing
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82580

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82581
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82582

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82580, 82582

        Case 82581
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171096)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82583

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82584

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82585

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82583, 82585

        Case 82584
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171097)

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
        If lErro <> SUCESSO Then gError 82593

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82593

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171098)

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
        If lErro <> SUCESSO Then gError 82592

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82592

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171099)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 43188
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 43220
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 43188, 43220
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171100)

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
Dim sVend_I As String
Dim sVend_F As String
Dim sProd_I As String
Dim sProd_F As String
Dim iIndice As Integer
Dim sNivel As String
Dim sTipoNivel As String
Dim sCliente_I As String
Dim sCliente_F As String
Dim sRepr As String
Dim iDet As Integer
Dim iExibe As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sVend_I, sVend_F, sCliente_I, sCliente_F, sRepr)
    If lErro <> SUCESSO Then gError 37950
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37951
    
    lErro = objRelOpcoes.IncluirParametro("TPRODDE", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 37952

    lErro = objRelOpcoes.IncluirParametro("TPRODATE", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 37953
         
    lErro = objRelOpcoes.IncluirParametro("TVENDDE", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37954

    lErro = objRelOpcoes.IncluirParametro("TVENDATE", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 37955
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37956

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37957
    
    If DataVigInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVIGDE", DataVigInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVIGDE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37956

    If DataVigFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVIGATE", DataVigFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVIGATE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 37957
    
    lErro = objRelOpcoes.IncluirParametro("TCLIDE", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54849

    lErro = objRelOpcoes.IncluirParametro("TCLIATE", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54850
    
    lErro = objRelOpcoes.IncluirParametro("TREPR", ReprInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54849

'    lErro = objRelOpcoes.IncluirParametro("TREPRATE", ReprFinal.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 54850
    
    If Detalhado.Value Then
        iDet = MARCADO
    Else
        iDet = DESMARCADO
    End If
        
    lErro = objRelOpcoes.IncluirParametro("NDET", CStr(iDet))
    If lErro <> AD_BOOL_TRUE Then gError 37959
    
    If ReprExibeVendas.Value Then
        iExibe = MARCADO
    Else
        iExibe = DESMARCADO
    End If
        
    lErro = objRelOpcoes.IncluirParametro("NEXIBE", CStr(iExibe))
    If lErro <> AD_BOOL_TRUE Then gError 37959
        
    lErro = objRelOpcoes.IncluirParametro("NTIPO", CStr(Tipo.ListIndex))
    If lErro <> AD_BOOL_TRUE Then gError 37959
        
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sVend_I, sVend_F, sCliente_I, sCliente_F, sRepr)
    If lErro <> SUCESSO Then gError 37960
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37950 To 37960

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171101)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 37961

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 37962

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 43189
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then gError 43221
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 37961
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 37962, 43189, 43221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171102)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim sTsk As String

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 37963
    
    sTsk = "TRPVou"
    
    sTsk = sTsk & CStr(Tipo.ItemData(Tipo.ListIndex))
    
    If Detalhado.Value Then
        sTsk = sTsk & "D"
    Else
        sTsk = sTsk & "R"
    End If
    
    gobjRelatorio.sNomeTsk = sTsk

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 37963

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171103)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 37964

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 37965

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 37966

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 43190
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 37964
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 37965, 37966, 43190
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171104)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String, sRepr As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If

   If sVend_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Promotor >= " & Forprint_ConvInt(CInt(sVend_I))
        
   End If

   If sVend_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Promotor <= " & Forprint_ConvInt(CInt(sVend_F))
        
    End If
    
   If sCliente_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ClienteVou >= " & Forprint_ConvLong(CLng(sCliente_I))
        
   End If

   If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ClienteVou <= " & Forprint_ConvLong(CLng(sCliente_F))

    End If
    
   If sRepr <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " ( Representante = " & Forprint_ConvLong(CLng(sRepr))
        
        If ReprExibeVendas.Value = vbChecked Then
            sExpressao = sExpressao & " OU ClienteVou = " & Forprint_ConvLong(CLng(sRepr))
        End If
            
        sExpressao = sExpressao & " ) "
        
   End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
    
    If Trim(DataVigInicial.ClipText) <> "" And Trim(DataVigFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " (( DataVigenciaDe >= " & Forprint_ConvData(CDate(DataVigInicial.Text))

        sExpressao = sExpressao & " E DataVigenciaDe <= " & Forprint_ConvData(CDate(DataVigFinal.Text))

        sExpressao = sExpressao & " ) OU ( DataVigenciaAte >= " & Forprint_ConvData(CDate(DataVigInicial.Text))

        sExpressao = sExpressao & " E DataVigenciaAte <= " & Forprint_ConvData(CDate(DataVigFinal.Text))

        sExpressao = sExpressao & " ) OU ( DataVigenciaDe <= " & Forprint_ConvData(CDate(DataVigInicial.Text))

        sExpressao = sExpressao & " E DataVigenciaAte >= " & Forprint_ConvData(CDate(DataVigFinal.Text)) & " )) "

    ElseIf Trim(DataVigFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVigenciaDe <= " & Forprint_ConvData(CDate(DataVigFinal.Text))
        
    ElseIf Trim(DataVigInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVigenciaAte >= " & Forprint_ConvData(CDate(DataVigInicial.Text))

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171105)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String, sRepr As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 37967

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 37968

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 37969

    End If
    
    'critica Vendedor Inicial e Final
    If VendedorInicial.Text <> "" Then
        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
    Else
        sVend_I = ""
    End If
    
    If VendedorFinal.Text <> "" Then
        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
    Else
        sVend_F = ""
    End If
            
    If sVend_I <> "" And sVend_F <> "" Then
        
        If CInt(sVend_I) > CInt(sVend_F) Then gError 37970
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 37971
    
    End If
        
    'critica Cliente Inicial e Final
    If ClienteInicial.Text <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_I = ""
    End If
    
    If ClienteFinal.Text <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_F = ""
    End If
            
    If sCliente_I <> "" And sCliente_F <> "" Then
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 20850
        
    End If
    
    'critica Cliente Inicial e Final
    If ReprInicial.Text <> "" Then
        sRepr = CStr(LCodigo_Extrai(ReprInicial.Text))
    Else
        sRepr = ""
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 20850
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                     
        Case 20851
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ReprInicial.SetFocus
                     
         Case 37967
            ProdutoInicial.SetFocus

        Case 37968
            ProdutoFinal.SetFocus

        Case 37969
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
       
        Case 37970
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
        
         Case 37971
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
                              
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171106)

    End Select

    Exit Function

End Function

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataVigFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVigFinal)

End Sub

Private Sub DataVigInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVigInicial)

End Sub

Private Sub LabelVendedorAte_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorDe_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        Call VendedorInicial_Validate(bSGECancelDummy)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        Call VendedorFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 37973

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37973
            'Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171107)

    End Select

End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 37974

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37974
             'Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171108)

    End Select

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 37975

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171109)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 37976

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37976

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171110)

    End Select

    Exit Sub

End Sub

Private Sub DataVigFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataVigFinal_Validate

    If Len(DataVigFinal.ClipText) > 0 Then

        sDataFim = DataVigFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 37975

    End If

    Exit Sub

Erro_DataVigFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 37975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171109)

    End Select

    Exit Sub

End Sub

Private Sub DataVigInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataVigInicial_Validate

    If Len(DataVigInicial.ClipText) > 0 Then

        sDataInic = DataVigInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 37976

    End If

    Exit Sub

Erro_DataVigInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 37976

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171110)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37977

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 37977
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171111)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37978

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 37978
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171112)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37979

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 37979
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171113)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37980

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 37980
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171114)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVigDe_DownClick

    lErro = Data_Up_Down_Click(DataVigInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37977

    Exit Sub

Erro_UpDownVigDe_DownClick:

    Select Case gErr

        Case 37977
            DataVigInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171111)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVigDe_UpClick

    lErro = Data_Up_Down_Click(DataVigInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37978

    Exit Sub

Erro_UpDownVigDe_UpClick:

    Select Case gErr

        Case 37978
            DataVigInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171112)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVigAte_DownClick

    lErro = Data_Up_Down_Click(DataVigFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37979

    Exit Sub

Erro_UpDownVigAte_DownClick:

    Select Case gErr

        Case 37979
            DataVigFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171113)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVigAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVigAte_UpClick

    lErro = Data_Up_Down_Click(DataVigFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37980

    Exit Sub

Erro_UpDownVigAte_UpClick:

    Select Case gErr

        Case 37980
            DataVigFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171114)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 126796

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 126796
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171115)

    End Select

End Sub


Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 126797

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 126797
             Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171116)

    End Select

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 37988
    
    If lErro <> SUCESSO Then gError 43251

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37988

        Case 43251
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171117)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 37989
    
    If lErro <> SUCESSO Then gError 43252

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37989

        Case 43252
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171118)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDO_VENDEDOR_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Vouchers Emitidos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTRPVouchers"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is VendedorInicial Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorFinal Then
            Call LabelVendedorAte_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
             Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is ReprInicial Then
            Call LabelReprDe_Click
        End If
    
    End If

End Sub

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

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

'Private Sub LabelReprAte_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelReprAte, Source, X, Y)
'End Sub
'
'Private Sub LabelReprAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelReprAte, Button, Shift, X, Y)
'End Sub

Private Sub LabelReprDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelReprDe, Source, X, Y)
End Sub

Private Sub LabelReprDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelReprDe, Button, Shift, X, Y)
End Sub

'Private Sub LabelReprAte_Click()
'
'Dim objRepr As New ClassCliente
'Dim colSelecao As Collection
'
'    giReprInicial = 0
'
'    If Len(Trim(ReprFinal.Text)) > 0 Then
'        'Preenche com o Repr da tela
'        objRepr.lCodigo = LCodigo_Extrai(ReprFinal.Text)
'    End If
'
'    'Chama Tela ReprsLista
'    Call Chama_Tela("ClientesLista", colSelecao, objRepr, objEventoRepr)
'
'End Sub

Private Sub LabelReprDe_Click()

Dim objRepr As New ClassCliente
Dim colSelecao As Collection

    giReprInicial = 1

    If Len(Trim(ReprInicial.Text)) > 0 Then
        'Preenche com o Repr da tela
        objRepr.lCodigo = LCodigo_Extrai(ReprInicial.Text)
    End If
    
    'Chama Tela ReprsLista
    Call Chama_Tela("ClientesLista", colSelecao, objRepr, objEventoRepr)

End Sub

Private Sub ReprInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRepr As New ClassCliente

On Error GoTo Erro_ReprInicial_Validate

    If Len(Trim(ReprInicial.Text)) > 0 Then
   
        'Tenta ler o Repr (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ReprInicial, objRepr, 0)
        If lErro <> SUCESSO Then gError 126796

    End If
    
    giReprInicial = 1
    
    Exit Sub

Erro_ReprInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 126796
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171115)

    End Select

End Sub

'Private Sub ReprFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objRepr As New ClassCliente
'
'On Error GoTo Erro_ReprFinal_Validate
'
'    If Len(Trim(ReprFinal.Text)) > 0 Then
'
'        'Tenta ler o Repr (NomeReduzido ou Código)
'        lErro = TP_Cliente_Le2(ReprFinal, objRepr, 0)
'        If lErro <> SUCESSO Then gError 126797
'
'    End If
'
'    giReprInicial = 0
'
'    Exit Sub
'
'Erro_ReprFinal_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 126797
'             Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objRepr.lCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171116)
'
'    End Select
'
'End Sub
