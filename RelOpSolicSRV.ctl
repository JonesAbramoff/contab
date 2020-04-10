VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpSolicSRV 
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   7920
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   1275
      Left            =   5625
      TabIndex        =   45
      Top             =   1695
      Width           =   2085
      Begin VB.OptionButton OpStatus 
         Caption         =   "Atendido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   270
         TabIndex        =   17
         Top             =   885
         Width           =   1470
      End
      Begin VB.OptionButton OpStatus 
         Caption         =   "Aberto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   16
         Top             =   555
         Width           =   1215
      End
      Begin VB.OptionButton OpStatus 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   270
         TabIndex        =   15
         Top             =   225
         Width           =   1110
      End
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Solicitações"
      Height          =   555
      Left            =   30
      TabIndex        =   42
      Top             =   1140
      Width           =   5505
      Begin MSMask.MaskEdBox SSDe 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   210
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SSAte 
         Height          =   300
         Left            =   3255
         TabIndex        =   6
         Top             =   195
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
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
         Index           =   3
         Left            =   225
         TabIndex        =   44
         Top             =   255
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
         Height          =   195
         Index           =   4
         Left            =   2805
         TabIndex        =   43
         Top             =   225
         Width           =   360
      End
   End
   Begin VB.Frame FrameAtendentes 
      Caption         =   "Atendentes"
      Height          =   645
      Left            =   45
      TabIndex        =   39
      Top             =   4035
      Width           =   7665
      Begin VB.ComboBox AtendenteAte 
         Height          =   315
         Left            =   3345
         TabIndex        =   14
         Top             =   225
         Width           =   2055
      End
      Begin VB.ComboBox AtendenteDe 
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   225
         Width           =   2055
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
         Index           =   6
         Left            =   2940
         TabIndex        =   41
         Top             =   285
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
         Index           =   5
         Left            =   195
         TabIndex        =   40
         Top             =   285
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Serviços"
      Height          =   1035
      Left            =   30
      TabIndex        =   34
      Top             =   2985
      Width           =   7680
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   615
         TabIndex        =   11
         Top             =   225
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   615
         TabIndex        =   12
         Top             =   615
         Width           =   1935
         _ExtentX        =   3413
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
         TabIndex        =   38
         Top             =   660
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
         TabIndex        =   37
         Top             =   255
         Width           =   360
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   36
         Top             =   225
         Width           =   4980
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   35
         Top             =   615
         Width           =   4995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      Height          =   615
      Left            =   30
      TabIndex        =   31
      Top             =   2355
      Width           =   5505
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   9
         Top             =   210
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3255
         TabIndex        =   10
         Top             =   210
         Width           =   2160
         _ExtentX        =   3810
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   255
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
         TabIndex        =   32
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpSolicSRV.ctx":0000
      Left            =   1380
      List            =   "RelOpSolicSRV.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   2730
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
      Left            =   5880
      Picture         =   "RelOpSolicSRV.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   870
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   570
      Left            =   30
      TabIndex        =   27
      Top             =   540
      Width           =   5505
      Begin MSComCtl2.UpDown UpDownDe 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   615
         TabIndex        =   1
         Top             =   180
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownAte 
         Height          =   315
         Left            =   4185
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   195
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   225
         Width           =   345
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
         Left            =   2805
         TabIndex        =   28
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vendedores"
      Height          =   630
      Left            =   30
      TabIndex        =   24
      Top             =   1695
      Width           =   5505
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   7
         Top             =   225
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   8
         Top             =   225
         Width           =   2160
         _ExtentX        =   3810
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   285
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   270
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5655
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpSolicSRV.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpSolicSRV.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpSolicSRV.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpSolicSRV.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   19
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
      Index           =   0
      Left            =   675
      TabIndex        =   30
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "RelOpSolicSRV"
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

Dim giClienteInicial As Integer
Dim giProdInicial As Integer
Dim giVendedorInicial As Integer

Const TELA_STATUS_TODOS = 0
Const TELA_STATUS_ABERTAS = 1
Const TELA_STATUS_ATENDIDAS = 2

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 202455

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 202456
    
    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 202457
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 202458
    
    AtendenteDe.Text = ""
    AtendenteAte.Text = ""

    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 202455 To 202458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202459)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    giVendedorInicial = 1
    giProdInicial = 1
    giClienteInicial = 1
    
    OpStatus(TELA_STATUS_TODOS).Value = True
           
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202460)

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

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 202461
   
    'pega Vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TVENDINIC", sParam)
    If lErro Then gError 202462
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  Vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TVENDFIM", sParam)
    If lErro Then gError 202463
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then gError 202464

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 202465

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then gError 202466

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 202467
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 202468

    Call DateParaMasked(DataDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 202469

    Call DateParaMasked(DataAte, StrParaDate(sParam))
       
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCLIENTEINIC", sParam)
    If lErro Then gError 202470
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("TCLIENTEFIM", sParam)
    If lErro Then gError 202471
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'pega a SS inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NSSDE", sParam)
    If lErro Then gError 202472
    
    SSDe.Text = sParam
    Call SSDe_Validate(bSGECancelDummy)
    
    'pega a SS final e exibe
    lErro = objRelOpcoes.ObterParametro("NSSATE", sParam)
    If lErro Then gError 202473
    
    SSAte.Text = sParam
    Call SSAte_Validate(bSGECancelDummy)
    
    'pega o Atendente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TATENDDE", sParam)
    If lErro Then gError 202474
    
    AtendenteDe.Text = sParam
    Call AtendenteDe_Validate(bSGECancelDummy)
    
    'pega o Atendente final e exibe
    lErro = objRelOpcoes.ObterParametro("TATENDATE", sParam)
    If lErro Then gError 202475
    
    AtendenteAte.Text = sParam
    Call AtendenteAte_Validate(bSGECancelDummy)
    
    'pega o status
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro Then gError 202476
    
    OpStatus(StrParaInt(sParam)).Value = True
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 202461 To 202476

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202477)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 202478
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 202479

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 202478
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 202479
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202480)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 202481

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 202482
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 202483

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 202481, 202483

        Case 202482
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202484)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 202485

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 202486

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 202487

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 202485, 202487

        Case 202486
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202488)

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
        If lErro <> SUCESSO Then gError 202489

        objProduto.sCodigo = sProdutoFormatado

    End If

    colSelecao.Add NATUREZA_PROD_SERVICO

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProdutoAte, "Natureza = ?")

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 202489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202490)

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
        If lErro <> SUCESSO Then gError 202491

        objProduto.sCodigo = sProdutoFormatado

    End If

    colSelecao.Add NATUREZA_PROD_SERVICO

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProdutoDe, "Natureza = ?")

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 202491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202492)

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
    If lErro <> SUCESSO Then gError 202493
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 202494
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 202493, 202494
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202495)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sVend_I As String
Dim sVend_F As String
Dim sProd_I As String
Dim sProd_F As String
Dim sCliente_I As String
Dim sCliente_F As String
Dim iStatus As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sVend_I, sVend_F, sCliente_I, sCliente_F, iStatus)
    If lErro <> SUCESSO Then gError 202496
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 202497
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 202498

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 202499
         
    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202500

    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202501
    
    lErro = objRelOpcoes.IncluirParametro("TATENDDE", AtendenteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202502

    lErro = objRelOpcoes.IncluirParametro("TATENDATE", AtendenteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202503
    
    lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(StrParaDate(DataDe.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202504

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(StrParaDate(DataAte.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202505
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202506

    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202507
    
    lErro = objRelOpcoes.IncluirParametro("NSSDE", SSDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202508
    
    lErro = objRelOpcoes.IncluirParametro("NSSATE", SSAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202509
    
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", CStr(iStatus))
    If lErro <> AD_BOOL_TRUE Then gError 202510
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, sVend_I, sVend_F, sCliente_I, sCliente_F, iStatus)
    If lErro <> SUCESSO Then gError 202511
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 202496 To 202511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202512)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 202513

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELANALISEVENDAS")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 202514

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 202513
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 202514

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202515)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 202516

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 202516

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202517)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 202518

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 202519

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 202520

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 202521
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 202518
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 202519 To 202521
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202522)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String, iStatus As Integer) As Long
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
        sExpressao = sExpressao & "Vendedor >= " & Forprint_ConvInt(CInt(sVend_I))
   End If

   If sVend_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(sVend_F))
    End If
    
   If iStatus <> TELA_STATUS_TODOS Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Status = " & Forprint_ConvInt(iStatus)
    End If

   If sCliente_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))
   End If

   If sCliente_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))
    End If
    
   If SSDe.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo >= " & Forprint_ConvLong(CLng(SSDe.Text))
   End If

   If SSAte.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvLong(CLng(SSAte.Text))
    End If
    
   If Codigo_Extrai(AtendenteDe.Text) <> 0 Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente >= " & Forprint_ConvInt(Codigo_Extrai(AtendenteDe.Text))
   End If

   If Codigo_Extrai(AtendenteAte.Text) <> 0 Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente <= " & Forprint_ConvInt(Codigo_Extrai(AtendenteAte.Text))
    End If

    If Trim(DataDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))
    End If

    If Trim(DataAte.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202523)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String, iStatus As Integer) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    If OpStatus(TELA_STATUS_ABERTAS) Then
        iStatus = TELA_STATUS_ABERTAS
    ElseIf OpStatus(TELA_STATUS_TODOS) Then
        iStatus = TELA_STATUS_TODOS
    Else
        iStatus = TELA_STATUS_ATENDIDAS
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 202524

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 202525

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
        If sProd_I > sProd_F Then gError 202526
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
        If StrParaInt(sVend_I) > StrParaInt(sVend_F) Then gError 202527
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 202528
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
        If StrParaInt(sCliente_I) > StrParaInt(sCliente_F) Then gError 202529
    End If
    
    If SSDe.Text <> "" And SSAte.Text <> "" Then
        If StrParaDbl(SSDe.Text) > StrParaDbl(SSAte.Text) Then gError 202530
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 202524
            ProdutoInicial.SetFocus

        Case 202525
            ProdutoFinal.SetFocus
                     
        Case 202526
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
       
        Case 202527
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
        
        Case 202528
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
        
        Case 202529
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
            
        Case 202530
            Call Rotina_Erro(vbOKOnly, "ERRO_SS_INICIAL_MAIOR", gErr)
            SSDe.SetFocus
                                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202531)

    End Select

    Exit Function

End Function

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
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
        If lErro <> SUCESSO Then gError 202532

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 202532

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202533)

    End Select

End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 202534

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 202534
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202535)

    End Select

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 202536

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 202536

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202537)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 202538

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 202538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202539)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202540

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 202540
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202541)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202542

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 202542
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202543)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202544

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 202544
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202545)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202546

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 202546
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202547)

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
        If lErro <> SUCESSO Then gError 202548

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 202548

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202549)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 202550

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 202550

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202551)

    End Select

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202552
    
    If lErro <> SUCESSO Then gError 202553

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 202552

        Case 202553
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202554)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202555
    
    If lErro <> SUCESSO Then gError 202556

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 202555

        Case 202556
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202557)

    End Select

    Exit Sub

End Sub

Private Sub SSDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SSDe_Validate

    If Len(Trim(SSDe.Text)) <> 0 Then
        
    End If

    Exit Sub

Erro_SSDe_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202558)

    End Select

    Exit Sub

End Sub

Private Sub SSAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SSAte_Validate

    If Len(Trim(SSAte.Text)) <> 0 Then
        
    End If

    Exit Sub

Erro_SSAte_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202559)

    End Select

    Exit Sub

End Sub
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDO_VENDEDOR_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Solicitação de Serviços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpSolicSRV"
    
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

Public Sub AtendenteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteDe_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteDe)
    If lErro <> SUCESSO Then gError 202560
        
    Exit Sub
    
Erro_AtendenteDe_Validate:

    Cancel = True
    
    Select Case gErr
                
        Case 202560
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202561)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 202562
        
    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202562
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202563)

    End Select

End Sub
