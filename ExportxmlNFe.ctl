VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ExportxmlNFeOCX 
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   ScaleHeight     =   6255
   ScaleWidth      =   7275
   Begin VB.CheckBox EnviarEmail 
      Caption         =   "Enviar por Email"
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
      Left            =   4635
      TabIndex        =   4
      Top             =   855
      Width           =   2130
   End
   Begin VB.Frame FrameEmail 
      Caption         =   "Envio por email"
      Enabled         =   0   'False
      Height          =   1680
      Left            =   135
      TabIndex        =   49
      Top             =   3210
      Width           =   7050
      Begin VB.CheckBox EmailMsgAuto 
         Caption         =   "Auto"
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
         Left            =   6210
         TabIndex        =   54
         Top             =   1005
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox EmailMensagem 
         Enabled         =   0   'False
         Height          =   660
         Left            =   960
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   945
         Width           =   5220
      End
      Begin VB.CheckBox EmailAssuntoAuto 
         Caption         =   "Auto"
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
         Left            =   6210
         TabIndex        =   52
         Top             =   645
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox EmailPara 
         Height          =   315
         Left            =   945
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   5220
      End
      Begin VB.TextBox EmailAssunto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         MaxLength       =   250
         TabIndex        =   12
         Top             =   585
         Width           =   5220
      End
      Begin VB.Label Label1 
         Caption         =   "Msg:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   465
         TabIndex        =   53
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Para:"
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
         Index           =   3
         Left            =   420
         TabIndex        =   51
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Assunto:"
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
         Left            =   135
         TabIndex        =   50
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4905
      ScaleHeight     =   495
      ScaleWidth      =   2220
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   15
      Width           =   2280
      Begin VB.CommandButton BotaoExecutar 
         Height          =   360
         Left            =   150
         Picture         =   "ExportxmlNFe.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Exportar Arquivos"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1170
         Picture         =   "ExportxmlNFe.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1665
         Picture         =   "ExportxmlNFe.ctx":05CC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   660
         Picture         =   "ExportxmlNFe.ctx":074A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "ExportxmlNFe.ctx":08A4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bordero de Cobrança"
      Height          =   570
      Left            =   135
      TabIndex        =   46
      Top             =   2130
      Width           =   7035
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   2220
      End
      Begin MSMask.MaskEdBox NumBordero 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumBordero 
         AutoSize        =   -1  'True
         Caption         =   "No. Borderô:"
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
         Left            =   3660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   48
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cobrador:"
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
         Left            =   30
         TabIndex        =   47
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.CheckBox optEmpresaToda 
      Caption         =   "Empresa Toda"
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
      Left            =   4635
      TabIndex        =   3
      Top             =   600
      Width           =   1680
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clientes"
      Height          =   510
      Left            =   135
      TabIndex        =   40
      Top             =   2700
      Width           =   7035
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   945
         TabIndex        =   9
         Top             =   150
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3795
         TabIndex        =   10
         Top             =   150
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Left            =   3390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   210
         Width           =   360
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
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
         Top             =   195
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localização do arquivo"
      Height          =   1065
      Left            =   135
      TabIndex        =   38
      Top             =   1065
      Width           =   7050
      Begin VB.CheckBox optNomeAuto 
         Caption         =   "Nome do Arquivo Automático"
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
         Left            =   915
         TabIndex        =   22
         Top             =   810
         Value           =   1  'Checked
         Width           =   3195
      End
      Begin VB.TextBox NomeArquivo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   930
         TabIndex        =   6
         Top             =   495
         Width           =   5370
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
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
         Left            =   6345
         TabIndex        =   21
         Top             =   150
         Width           =   555
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   930
         TabIndex        =   5
         Top             =   195
         Width           =   5430
      End
      Begin VB.Label Label1 
         Caption         =   "Arquivo:"
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   495
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   ".zip"
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
         Height          =   255
         Index           =   5
         Left            =   6375
         TabIndex        =   43
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
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
         Index           =   1
         Left            =   75
         TabIndex        =   39
         Top             =   225
         Width           =   795
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período de emissão da NF"
      Height          =   570
      Left            =   135
      TabIndex        =   34
      Top             =   495
      Width           =   4410
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2070
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   930
         TabIndex        =   1
         Top             =   195
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4020
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   330
         Left            =   2895
         TabIndex        =   2
         Top             =   180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoDe 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   525
         TabIndex        =   36
         Top             =   225
         Width           =   390
      End
      Begin VB.Label LabelPeriodoAte 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2475
         TabIndex        =   35
         Top             =   225
         Width           =   450
      End
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   6210
      TabIndex        =   14
      Top             =   4965
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acompanhamento"
      Height          =   1365
      Left            =   135
      TabIndex        =   26
      Top             =   4875
      Width           =   6045
      Begin MSComctlLib.ProgressBar PB 
         Height          =   405
         Left            =   60
         TabIndex        =   27
         Top             =   645
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Texto 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   45
         Top             =   1020
         Width           =   5820
      End
      Begin VB.Label Label5 
         Caption         =   "Copiando....."
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
         Left            =   195
         TabIndex        =   37
         Top             =   405
         Width           =   2100
      End
      Begin VB.Label RegTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   5160
         TabIndex        =   33
         Top             =   180
         Width           =   510
      End
      Begin VB.Label RegAtual 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2280
         TabIndex        =   32
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
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
         Index           =   0
         Left            =   4425
         TabIndex        =   31
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Registros processados:"
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
         Left            =   180
         TabIndex        =   30
         Top             =   180
         Width           =   2100
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5820
         TabIndex        =   29
         Top             =   420
         Width           =   120
      End
      Begin VB.Label perccompleto 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   4395
         TabIndex        =   28
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "ExportxmlNFe.ctx":0DD6
      Left            =   1050
      List            =   "ExportxmlNFe.ctx":0DD8
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   3375
   End
   Begin VB.Label Label4 
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
      Left            =   345
      TabIndex        =   25
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "ExportxmlNFeOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Dim iCancelar As Integer
Dim iExecutando As Integer

Private WithEvents objEventoClienteInicial As AdmEvento
Attribute objEventoClienteInicial.VB_VarHelpID = -1
Private WithEvents objEventoClienteFinal As AdmEvento
Attribute objEventoClienteFinal.VB_VarHelpID = -1
Private WithEvents objEventoBorderoCobranca As AdmEvento
Attribute objEventoBorderoCobranca.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoClienteInicial = New AdmEvento
    Set objEventoClienteFinal = New AdmEvento
    Set objEventoBorderoCobranca = New AdmEvento
    
    lErro = Carrega_Cobradores()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iCancelar = DESMARCADO
    iExecutando = DESMARCADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209216)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError ERRO_SEM_MENSAGEM
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataFinal, CDate(sParam))
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaLong(sParam) > 0 Then
        ClienteInicial.Text = sParam
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteInicial.Text = ""
    End If
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaLong(sParam) > 0 Then
        ClienteFinal.Text = sParam
        Call ClienteFinal_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = ""
    End If
        
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("TNOMEDIR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NomeDiretorio.Text = sParam
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("TNOMEARQ", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NomeArquivo.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NEMPTODA", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        optEmpresaToda.Value = vbChecked
    Else
        optEmpresaToda.Value = vbUnchecked
    End If

    lErro = objRelOpcoes.ObterParametro("NNOMEAUTO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        optNomeAuto.Value = vbChecked
        Call Calcula_NomeArq
    Else
        optNomeAuto.Value = vbUnchecked
    End If
    
    
    lErro = objRelOpcoes.ObterParametro("TEMAILPARA", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    EmailPara.Text = sParam
    
    
    lErro = objRelOpcoes.ObterParametro("TEMAILASSUNTO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    EmailAssunto.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("TEMAILMSG", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    EmailMensagem.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NENVIAREMAIL", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        EnviarEmail.Value = vbChecked
    Else
        EnviarEmail.Value = vbUnchecked
    End If
    Call EnviarEmail_Click
    
    lErro = objRelOpcoes.ObterParametro("NEMAILASSUNTOAUTO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        EmailAssuntoAuto.Value = vbChecked
        Call Calcula_NomeArq
    Else
        EmailAssuntoAuto.Value = vbUnchecked
    End If
    
    lErro = objRelOpcoes.ObterParametro("NEMAILMSGAUTO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        EmailMsgAuto.Value = vbChecked
        Call Calcula_NomeArq
    Else
        EmailMsgAuto.Value = vbUnchecked
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209217)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInicial = Nothing
    Set objEventoClienteFinal = Nothing
    Set objEventoBorderoCobranca = Nothing
    
End Sub

Function Trata_Parametros(Optional objRelatorio As AdmRelatorio, Optional objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 209218
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 209218
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209219)

    End Select

    Exit Function

End Function

Private Sub BotaoCancelar_Click()
    iCancelar = MARCADO
    If iExecutando = DESMARCADO Then Call BotaoFechar_Click
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCliente_De As String, sCliente_Ate As String

On Error GoTo Erro_Critica_Parametros

    If Len(Trim(NumBordero.Text)) = 0 Then
    
        If StrParaDate(DataInicial.Text) = DATA_NULA Then gError 209220
        If StrParaDate(DataFinal.Text) = DATA_NULA Then gError 209221
    
    End If
    
    If StrParaDate(DataInicial.Text) <> DATA_NULA And StrParaDate(DataFinal.Text) <> DATA_NULA And StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 209222

    'Verifica se o Cliente inicial foi preenchido
    If ClienteInicial.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteFinal.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        If CLng(sCliente_De) > CLng(sCliente_Ate) Then gError 209223
    End If
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 209248
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 209249
    
    If EnviarEmail.Value = vbChecked And Len(Trim(EmailPara.Text)) = 0 Then gError 213230
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr

        Case 209220
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
        
        Case 209221
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case 209222
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 209223
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus

        Case 209248
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
       
        Case 209249
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
            
        Case 213230
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCH_EMAIL", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209224)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    Cobrador.ListIndex = -1
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209225)

    End Select

    Exit Sub
    
End Sub

Private Sub ClienteFinal_Change()
    Call Calcula_NomeArq
End Sub

Private Sub ClienteInicial_Change()
    Call Calcula_NomeArq
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim objTela As Object
Dim iFilialEmpresa As Integer, iEmpToda As Integer, iNomeAuto As Integer
Dim objSerie As New ClassSerie, iAux As Integer
Dim objBorderoCobranca As New ClassBorderoCobranca, lNumBordero As Long
Dim objEnvioEmail As ClassEnvioDeEmail
Dim colEnvioEmail As New Collection
Dim sNomeArqParam As String

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If optEmpresaToda.Value = vbChecked Then
        iEmpToda = MARCADO
    Else
        iEmpToda = DESMARCADO
    End If
    
    If optEmpresaToda.Value = vbChecked Then
        iEmpToda = MARCADO
    Else
        iEmpToda = DESMARCADO
    End If
        
    If optNomeAuto.Value = vbChecked Then
        iNomeAuto = MARCADO
    Else
        iNomeAuto = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NEMPTODA", CStr(iEmpToda))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NNOMEAUTO", CStr(iNomeAuto))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", CStr(LCodigo_Extrai(ClienteInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", CStr(LCodigo_Extrai(ClienteFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEDIR", NomeDiretorio.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEARQ", NomeArquivo.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TEMAILPARA", EmailPara.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TEMAILASSUNTO", EmailAssunto.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TEMAILMSG", EmailMensagem.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    If EnviarEmail.Value = vbChecked Then
        iAux = MARCADO
    Else
        iAux = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NENVIAREMAIL", CStr(iAux))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If EmailAssuntoAuto.Value = vbChecked Then
        iAux = MARCADO
    Else
        iAux = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NEMAILASSUNTOAUTO", CStr(iAux))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If EmailMsgAuto.Value = vbChecked Then
        iAux = MARCADO
    Else
        iAux = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NEMAILMSGAUTO", CStr(iAux))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If optEmpresaToda.Value = vbChecked Then
        iFilialEmpresa = EMPRESA_TODA
    Else
        objSerie.iFilialEmpresa = giFilialEmpresa

        lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iFilialEmpresa = objSerie.iFilialEmpresa
    End If
    
    If Len(Trim(NumBordero.Text)) <> 0 Then
    
        objBorderoCobranca.lNumBordero = StrParaLong(NumBordero.Text)
        
        lErro = CF("BorderoCobranca_Le", objBorderoCobranca)
        If lErro <> SUCESSO And lErro <> 46366 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 46366 Then gError 66585
    
        lNumBordero = objBorderoCobranca.lNumBordero
    
    Else
    
        lNumBordero = 0
    
    End If
    
    If bExecutando Then
              
        lErro = Monta_Expressao_Selecao(objRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
        Set objTela = Me
        iExecutando = MARCADO

        RegAtual.Caption = 0
        perccompleto.Caption = Format(0, "#0.00")
        PB.Value = 0
        Texto.Caption = ""
    
        lErro = CF("NFe_Exporta_Xml", objTela, NomeDiretorio.Text, NomeArquivo.Text, iFilialEmpresa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), LCodigo_Extrai(ClienteInicial.Text), LCodigo_Extrai(ClienteFinal.Text), lNumBordero)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If EnviarEmail.Value = vbChecked Then

            Set objEnvioEmail = New ClassEnvioDeEmail
            colEnvioEmail.Add objEnvioEmail
            
            objEnvioEmail.sEmail = EmailPara.Text
            objEnvioEmail.sAssunto = EmailAssunto.Text
            objEnvioEmail.sMensagem = "<html><head><title>XMLS</title></head><body>" & Replace(EmailMensagem.Text, vbNewLine, "<br>") & "</body></html>"
            objEnvioEmail.sAnexo = NomeDiretorio.Text & NomeArquivo.Text & ".zip"
            objEnvioEmail.iConfirmacaoLeitura = MARCADO

            lErro = Sistema_Preparar_Batch(sNomeArqParam)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = CF("Rotina_Envia_Emails_Batch", sNomeArqParam, colEnvioEmail)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        End If
        
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 66585
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO", gErr, objBorderoCobranca.lNumBordero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209226)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 209227

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 209227
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209228)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Call gobjRelatorio.Executar_Prossegue2(Me)
    
    'Call BotaoFechar_Click
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ARQUIVO_GERADO", NomeDiretorio.Text & NomeArquivo.Text & ".zip")

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209229)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 209230

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 209230
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209231)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209232)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTAS_CORRENTES
    Set Form_Load_Ocx = Me
    Caption = "Exportação dos arquivos xml das NFes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExportxmlNFe"
    
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

Private Sub EmailAssuntoAuto_Click()
    If EmailAssuntoAuto.Value = vbChecked Then
        Call Calcula_NomeArq
        EmailAssunto.Enabled = False
    Else
        EmailAssunto.Enabled = True
    End If
End Sub

Private Sub EmailMsgAuto_Click()
    If EmailMsgAuto.Value = vbChecked Then
        Call Calcula_NomeArq
        EmailMensagem.Enabled = False
    Else
        EmailMensagem.Enabled = True
    End If
End Sub

Private Sub EnviarEmail_Click()
    If EnviarEmail.Value = vbChecked Then
        FrameEmail.Enabled = True
    Else
        FrameEmail.Enabled = False
    End If
End Sub


Private Sub NumBordero_Validate(Cancel As Boolean)
    Call Calcula_NomeArq
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NumBordero Then
            Call LabelNUmBordero_Click
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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Public Function Inicializa_Copia(ByVal lCount As Long) As Long

On Error GoTo Erro_Inicializa_Copia

    RegTotal.Caption = CStr(lCount)

    Inicializa_Copia = SUCESSO

    Exit Function

Erro_Inicializa_Copia:

    Inicializa_Copia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209233)

    End Select

    Exit Function
    
End Function

Public Function Atualiza_Texto(ByVal sTexto As String) As Long

On Error GoTo Erro_Atualiza_Texto

    Texto.Caption = sTexto
    
    DoEvents

    Atualiza_Texto = SUCESSO

    Exit Function

Erro_Atualiza_Texto:

    Atualiza_Texto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209234)

    End Select

    Exit Function
    
End Function


Public Function Processa_Registro() As Long

On Error GoTo Erro_Processa_Registro

    RegAtual.Caption = CLng(RegAtual.Caption) + 1
    perccompleto.Caption = Format((CLng(RegAtual.Caption) / CLng(RegTotal.Caption)) * 100, "#0.00")
    PB.Value = StrParaDbl(perccompleto.Caption)

    DoEvents
    
    If iCancelar = MARCADO Then gError ERRO_SEM_MENSAGEM
    
    DoEvents
    
    Processa_Registro = SUCESSO

    Exit Function

Erro_Processa_Registro:

    Processa_Registro = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209235)

    End Select

    Exit Function
    
End Function

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física do arquivo"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209236)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 209237

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 209237, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209238)

    End Select

    Exit Sub

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objcliente, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209239)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objcliente, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209240)

    End Select

End Sub

Private Sub LabelClienteDe_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteDe_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteInicial.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = ClienteInicial.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteInicial, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteDe_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209241)
    
    End Select
    
End Sub

Private Sub LabelClienteAte_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteAte_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteFinal.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = ClienteFinal.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteFinal, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteAte_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209242)
    
    End Select
    
End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Call Calcula_NomeArq

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209243)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Call Calcula_NomeArq

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209244)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209245)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209246)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209247)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209248)

    End Select

    Exit Sub

End Sub

Private Sub optNomeAuto_Click()
    If optNomeAuto.Value = vbChecked Then
        Call Calcula_NomeArq
        NomeArquivo.Enabled = False
    Else
        NomeArquivo.Enabled = True
    End If
End Sub

Private Sub Calcula_NomeArq()
Dim sNomeArq As String, sAux As String
    If optNomeAuto.Value = vbChecked Then
    
        sNomeArq = "XML" & Format(IIf(optEmpresaToda.Value = vbChecked, 0, giFilialEmpresa), "00") & "_DTE" & Format(Date, "YYYYMMDD") & Format(Now, "HHMMSS")
        
        If Len(Trim(NumBordero.Text)) = 0 Then
        
            If StrParaDate(DataInicial.Text) <> DATA_NULA Then
                sNomeArq = sNomeArq & "_DTI" & Format(StrParaDate(DataInicial.Text), "YYYYMMDD")
            End If
            If StrParaDate(DataFinal.Text) <> DATA_NULA Then
                sNomeArq = sNomeArq & "_DTF" & Format(StrParaDate(DataFinal.Text), "YYYYMMDD")
            End If
            If LCodigo_Extrai(ClienteInicial.Text) <> 0 Then
                sNomeArq = sNomeArq & "_CI" & CStr(LCodigo_Extrai(ClienteInicial.Text))
            End If
            If LCodigo_Extrai(ClienteFinal.Text) <> 0 Then
                sNomeArq = sNomeArq & "_CF" & CStr(LCodigo_Extrai(ClienteFinal.Text))
            End If
        
        Else
        
            sNomeArq = sNomeArq & "_BORD" & CStr(StrParaLong(NumBordero.Text)) & IIf(Cobrador.ListIndex = -1, "", "_" & Replace(Cobrador.Text, " ", "_"))
            
        End If
        
        NomeArquivo.Text = sNomeArq
    End If
    
    sAux = ""
    If StrParaDate(DataInicial.Text) <> DATA_NULA Then
        sAux = sAux & " de " & Format(StrParaDate(DataInicial.Text), "dd/mm/yyyy")
    End If
    If StrParaDate(DataFinal.Text) <> DATA_NULA Then
        sAux = sAux & " até " & Format(StrParaDate(DataFinal.Text), "dd/mm/yyyy")
    End If
    
    If EmailAssuntoAuto.Value = vbChecked Then
        EmailAssunto.Text = gsNomeEmpresa & ": Xmls " & sAux
    End If
    
    If EmailMsgAuto.Value = vbChecked Then
        EmailMensagem.Text = "Segue em anexo um zip contendo os arquivos xmls referente ao período: " & sAux
    End If
    
End Sub

Private Sub optEmpresaToda_Click()
    Call Calcula_NomeArq
End Sub

Private Sub objEventoClienteInicial_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteInicial_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteInicial.Text = objcliente.sNomeReduzido
    
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteInicial_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209257)
    
    End Select

End Sub

Private Sub objEventoClienteFinal_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteFinal_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteFinal.Text = objcliente.sNomeReduzido
    
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteFinal_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209258)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelNUmBordero_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_LabelNUmBordero_Click
    
    If Len(Trim(Cobrador.Text)) = 0 Then gError 66586
    
    If Len(Trim(NumBordero.Text)) > 0 Then objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    colSelecao.Add Codigo_Extrai(Cobrador.Text)
    
    'Chama Tela BorderoCobrancaLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBorderoCobranca, objEventoBorderoCobranca)

    Exit Sub

Erro_LabelNUmBordero_Click:

    Select Case gErr

        Case 66586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167364)

    End Select

    Exit Sub
    
End Sub

Private Sub NumBordero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumBordero)

End Sub

Private Sub objEventoBorderoCobranca_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim lErro As Long

On Error GoTo Erro_objEventoBorderoCobranca_evSelecao

    Set objBorderoCobranca = obj1
    
    NumBordero.PromptInclude = False
    NumBordero.Text = objBorderoCobranca.lNumBordero
    NumBordero.PromptInclude = True
    
    Call Calcula_NomeArq
    
'    LabelEmissao.Caption = Format(objBorderoCobranca.dtDataEmissao, "dd/mm/yy")
'    LabelValor = Format(objBorderoCobranca.dValor, "Standard")
'
'    objCarteiraCobranca.iCodigo = objBorderoCobranca.iCodCarteiraCobranca
'
'    'Lê a Carteira de Cobrança
'    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
'    If lErro <> SUCESSO And lErro <> 23413 Then gError 66571
'
'    'Se não achou, erro
'    If lErro = 23413 Then gError 66572
'
'    'Coloca na a Carteira de Cobrança
'    LabelCarteira.Caption = objCarteiraCobranca.sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoBorderoCobranca_evSelecao:

    Select Case gErr
        
        Case 66571
        
        Case 66572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", gErr, objCarteiraCobranca.iCodigo)
                    
        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167365)

    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Cobradores() As Long

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim colCobrador As New Collection

On Error GoTo Erro_Carrega_Cobradores

    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", colCobrador)
    If lErro <> SUCESSO Then gError 66582
    
    'Preenche a ComboBox Cobrador com os objetos da coleção de Cobradores
    For Each objCobrador In colCobrador

        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next

    Carrega_Cobradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Cobradores:

    Carrega_Cobradores = gErr
    
    Select Case gErr
    
        Case 66582
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167366)
            
    End Select
    
    Exit Function

End Function

Private Sub LabelNumBordero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumBordero, Source, X, Y)
End Sub

Private Sub LabelNumBordero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumBordero, Button, Shift, X, Y)
End Sub

