VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpOrdemSRV 
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   7920
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   1275
      Left            =   5625
      TabIndex        =   42
      Top             =   1830
      Width           =   2085
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data de Entrega"
      Height          =   570
      Left            =   30
      TabIndex        =   39
      Top             =   1215
      Width           =   5505
      Begin MSComCtl2.UpDown UpDownEntDe 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEntDe 
         Height          =   300
         Left            =   615
         TabIndex        =   5
         Top             =   225
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEntAte 
         Height          =   315
         Left            =   4185
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEntAte 
         Height          =   300
         Left            =   3240
         TabIndex        =   7
         Top             =   240
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
         Index           =   6
         Left            =   2805
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
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   270
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordens"
      Height          =   585
      Left            =   30
      TabIndex        =   36
      Top             =   1830
      Width           =   5505
      Begin MSMask.MaskEdBox OSDe 
         Height          =   300
         Left            =   600
         TabIndex        =   9
         Top             =   210
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OSAte 
         Height          =   300
         Left            =   3255
         TabIndex        =   10
         Top             =   195
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelOSAte 
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   225
         Width           =   360
      End
      Begin VB.Label LabelOSDe 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   255
         Width           =   315
      End
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Solicitações"
      Height          =   630
      Left            =   30
      TabIndex        =   33
      Top             =   2475
      Width           =   5505
      Begin MSMask.MaskEdBox SSDe 
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
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
         TabIndex        =   12
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
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
         Index           =   1
         Left            =   225
         TabIndex        =   35
         Top             =   300
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
         Index           =   2
         Left            =   2805
         TabIndex        =   34
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Serviços"
      Height          =   1035
      Left            =   30
      TabIndex        =   28
      Top             =   3165
      Width           =   7680
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   615
         TabIndex        =   13
         Top             =   225
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   615
         TabIndex        =   14
         Top             =   615
         Width           =   2175
         _ExtentX        =   3836
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   255
         Width           =   360
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2775
         TabIndex        =   30
         Top             =   225
         Width           =   4650
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2775
         TabIndex        =   29
         Top             =   615
         Width           =   4665
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrdemSRV.ctx":0000
      Left            =   1380
      List            =   "RelOpOrdemSRV.ctx":0002
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
      Picture         =   "RelOpOrdemSRV.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   870
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data de Emissão"
      Height          =   600
      Left            =   30
      TabIndex        =   24
      Top             =   555
      Width           =   5505
      Begin MSComCtl2.UpDown UpDownDe 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
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
         Top             =   240
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
         Top             =   240
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
         Top             =   255
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
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   285
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
         Index           =   7
         Left            =   2805
         TabIndex        =   25
         Top             =   300
         Width           =   360
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
         Picture         =   "RelOpOrdemSRV.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOrdemSRV.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrdemSRV.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOrdemSRV.ctx":0940
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
      TabIndex        =   27
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOrdemSRV"
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

Dim giProdInicial As Integer
Dim giOSInicial As Integer

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
Private WithEvents objEventoOS As AdmEvento
Attribute objEventoOS.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoOS = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 202350

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 202351

    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 202350, 202351

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202352)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    giProdInicial = 1
    giOSInicial = 1
    
    OpStatus(TELA_STATUS_TODOS).Value = True
           
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202353)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 202354
   
    'pega a OS inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOSDE", sParam)
    If lErro Then gError 202355
    
    OSDe.Text = sParam
    Call OSDe_Validate(bSGECancelDummy)
    
    'pega  a OS final e exibe
    lErro = objRelOpcoes.ObterParametro("TOSATE", sParam)
    If lErro Then gError 202356
    
    OSAte.Text = sParam
    Call OSAte_Validate(bSGECancelDummy)
    
    'pega a SS inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TSSDE", sParam)
    If lErro Then gError 202357
    
    SSDe.Text = sParam
    Call SSDe_Validate(bSGECancelDummy)
    
    'pega  a SS final e exibe
    lErro = objRelOpcoes.ObterParametro("TSSATE", sParam)
    If lErro Then gError 202358
    
    SSAte.Text = sParam
    Call SSAte_Validate(bSGECancelDummy)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then gError 202359

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 202360

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then gError 202361

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 202362
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 202363

    Call DateParaMasked(DataDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 202364

    Call DateParaMasked(DataAte, StrParaDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENTINIC", sParam)
    If lErro <> SUCESSO Then gError 202365

    Call DateParaMasked(DataEntDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DENTFIM", sParam)
    If lErro <> SUCESSO Then gError 202366

    Call DateParaMasked(DataEntAte, StrParaDate(sParam))
    
    'pega o status
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro Then gError 202367
    
    OpStatus(StrParaInt(sParam)).Value = True
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 202354 To 202367

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202368)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 202369
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 202370

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 202369
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 202370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202371)

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

Private Sub LabelOSAte_Click()

Dim objOS As New ClassOS
Dim colSelecao As New Collection

    giOSInicial = 2

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(SSAte.Text)) <> 0 Then objOS.sCodigo = SSAte.Text
    
    'lista as OP's
    Call Chama_Tela("OSLista", colSelecao, objOS, objEventoOS)
    
End Sub

Private Sub LabelOSDe_Click()

Dim objOS As New ClassOS
Dim colSelecao As New Collection

    giOSInicial = 1

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(SSDe.Text)) <> 0 Then objOS.sCodigo = SSDe.Text
    
    'lista as OP's
    Call Chama_Tela("OSLista", colSelecao, objOS, objEventoOS)
    
End Sub

Private Sub objEventoOS_evSelecao(obj1 As Object)

Dim objOS As ClassOS

    Set objOS = obj1

    If giOSInicial = 1 Then
        SSDe.Text = objOS.sCodigo
    Else
        SSAte.Text = objOS.sCodigo
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 202372

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 202373
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 202374

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 202372, 202374

        Case 202373
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202375)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 202376

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 202377

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 202378

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 202376, 202378

        Case 202377
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202379)

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
        If lErro <> SUCESSO Then gError 202380

        objProduto.sCodigo = sProdutoFormatado

    End If
    
    colSelecao.Add NATUREZA_PROD_SERVICO

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProdutoAte, "Natureza = ?")

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 202380

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202381)

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
        If lErro <> SUCESSO Then gError 202382

        objProduto.sCodigo = sProdutoFormatado

    End If
    
    colSelecao.Add NATUREZA_PROD_SERVICO

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProdutoDe, "Natureza = ?")

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 202382

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202383)

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
    If lErro <> SUCESSO Then gError 202384
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 202385
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 202384, 202385
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202386)

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
Dim sProd_I As String
Dim sProd_F As String
Dim iStatus As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, iStatus)
    If lErro <> SUCESSO Then gError 202387
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 202388
    
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 202389

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 202390
         
    lErro = objRelOpcoes.IncluirParametro("TOSDE", OSDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202391

    lErro = objRelOpcoes.IncluirParametro("TOSATE", OSAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202392
    
    lErro = objRelOpcoes.IncluirParametro("TSSDE", SSDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202393

    lErro = objRelOpcoes.IncluirParametro("TSSATE", SSAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 202394
    
    lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(StrParaDate(DataDe.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202395

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(StrParaDate(DataAte.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202396
    
    lErro = objRelOpcoes.IncluirParametro("DENTINIC", CStr(StrParaDate(DataEntDe.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202397

    lErro = objRelOpcoes.IncluirParametro("DENTFIM", CStr(StrParaDate(DataEntAte.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 202398
           
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", CStr(iStatus))
    If lErro <> AD_BOOL_TRUE Then gError 202399
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, iStatus)
    If lErro <> SUCESSO Then gError 202400
           
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 202387 To 202400

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202401)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 202402

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELANALISEVENDAS")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 202403

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 202402
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 202403

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202404)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 202405

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 202405

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202406)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 202407

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 202408

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 202409

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 202410
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 202407
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 202408 To 202410
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202411)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, iStatus As Integer) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    If sProd_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)
    End If
    
   If iStatus <> TELA_STATUS_TODOS Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Status = " & Forprint_ConvInt(iStatus)
    End If
    
   If OSDe.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo >= " & Forprint_ConvTexto(OSDe.Text)
   End If

   If OSAte.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvTexto(OSAte.Text)
    End If
    
   If SSDe.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "SS >= " & Forprint_ConvLong(StrParaLong(SSDe.Text))
   End If

   If SSAte.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "SS <= " & Forprint_ConvLong(StrParaLong(SSAte.Text))
    End If

    If Trim(DataDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))
    End If

    If Trim(DataAte.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))
    End If
    
    If Trim(DataEntDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Entrega >= " & Forprint_ConvData(CDate(DataEntDe.Text))
    End If

    If Trim(DataEntAte.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Entrega <= " & Forprint_ConvData(CDate(DataEntAte.Text))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202412)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, iStatus As Integer) As Long
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
    If lErro <> SUCESSO Then gError 202413

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 202414

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then
        If sProd_I > sProd_F Then gError 202415
    End If
    
    If StrParaDate(DataDe.Text) <> DATA_NULA And StrParaDate(DataAte.Text) <> DATA_NULA Then
        'data inicial não pode ser maior que a data final
        If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
             If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 202416
        End If
    End If
    
    If StrParaDate(DataEntDe.Text) <> DATA_NULA And StrParaDate(DataEntAte.Text) <> DATA_NULA Then
        'data inicial não pode ser maior que a data final
        If Trim(DataEntDe.ClipText) <> "" And Trim(DataEntAte.ClipText) <> "" Then
             If StrParaDate(DataEntDe.Text) > StrParaDate(DataEntAte.Text) Then gError 202417
        End If
    End If
    
    If OSDe.ClipText <> "" And OSAte.ClipText <> "" Then
        If UCase(OSDe.Text) > UCase(OSAte.Text) Then gError 202418
    End If
    
    If SSDe.ClipText <> "" And SSAte.ClipText <> "" Then
        If StrParaLong(SSDe.Text) > StrParaLong(SSAte.Text) Then gError 202419
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 202413
            ProdutoInicial.SetFocus

        Case 202414
            ProdutoFinal.SetFocus
                     
        Case 202415
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
       
        Case 202416
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 202417
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataEntDe.SetFocus
            
        Case 202418
            Call Rotina_Erro(vbOKOnly, "ERRO_OS_INICIAL_MAIOR", gErr)
            OSDe.SetFocus
            
        Case 202419
            Call Rotina_Erro(vbOKOnly, "ERRO_SS_INICIAL_MAIOR", gErr)
            SSDe.SetFocus
                                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202420)

    End Select

    Exit Function

End Function

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub DataEntAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEntAte)
End Sub

Private Sub DataEntDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEntDe)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 202421

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 202421

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202422)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 202423

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 202423

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202424)

    End Select

    Exit Sub

End Sub

Private Sub DataEntAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntAte_Validate

    If Len(DataEntAte.ClipText) > 0 Then

        lErro = Data_Critica(DataEntAte.Text)
        If lErro <> SUCESSO Then gError 202425

    End If

    Exit Sub

Erro_DataEntAte_Validate:

    Cancel = True

    Select Case gErr

        Case 202425

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202426)

    End Select

    Exit Sub

End Sub

Private Sub DataEntDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntDe_Validate

    If Len(DataEntDe.ClipText) > 0 Then

        lErro = Data_Critica(DataEntDe.Text)
        If lErro <> SUCESSO Then gError 202427

    End If

    Exit Sub

Erro_DataEntDe_Validate:

    Cancel = True

    Select Case gErr

        Case 202427

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202428)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202429

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 202429
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202430)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202431

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 202431
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202432)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202433

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 202433
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202434)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202435

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 202435
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202436)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202437
    
    If lErro <> SUCESSO Then gError 202438

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 202437

        Case 202438
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202439)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202440
    
    If lErro <> SUCESSO Then gError 202441

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 202440

        Case 202441
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202442)

    End Select

    Exit Sub

End Sub

Private Sub OSDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OSDe_Validate

    If Len(Trim(OSDe.Text)) <> 0 Then
        
    End If

    Exit Sub

Erro_OSDe_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202443)

    End Select

    Exit Sub

End Sub

Private Sub OSAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OSAte_Validate

    If Len(Trim(OSAte.Text)) <> 0 Then
        
    End If

    Exit Sub

Erro_OSAte_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202444)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202445)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202446)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDO_VENDEDOR_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Ordens de Serviço"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "RelOpOrdemSRV"
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
        
        If Me.ActiveControl Is OSDe Then
            Call LabelOSDe_Click
        ElseIf Me.ActiveControl Is OSAte Then
            Call LabelOSAte_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub UpDownEntDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntDe_DownClick

    lErro = Data_Up_Down_Click(DataEntDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202447

    Exit Sub

Erro_UpDownEntDe_DownClick:

    Select Case gErr

        Case 202447
            DataEntDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202448)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntDe_UpClick

    lErro = Data_Up_Down_Click(DataEntDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202449

    Exit Sub

Erro_UpDownEntDe_UpClick:

    Select Case gErr

        Case 202449
            DataEntDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202450)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntAte_DownClick

    lErro = Data_Up_Down_Click(DataEntAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 202451

    Exit Sub

Erro_UpDownEntAte_DownClick:

    Select Case gErr

        Case 202451
            DataEntAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202452)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntAte_UpClick

    lErro = Data_Up_Down_Click(DataEntAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 202453

    Exit Sub

Erro_UpDownEntAte_UpClick:

    Select Case gErr

        Case 202453
            DataEntAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202454)

    End Select

    Exit Sub

End Sub
