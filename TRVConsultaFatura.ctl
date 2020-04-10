VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRVConsultaFatura 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame5 
      Caption         =   "Baixa"
      Height          =   570
      Left            =   105
      TabIndex        =   59
      Top             =   4740
      Width           =   9315
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pagto:"
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
         Index           =   22
         Left            =   630
         TabIndex        =   66
         Top             =   195
         Width           =   570
      End
      Begin VB.Label DataPagto 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   65
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo em aberto:"
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
         Index           =   21
         Left            =   3360
         TabIndex        =   64
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label SaldoPagto 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4875
         TabIndex        =   63
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label DataRegPagto 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7260
         TabIndex        =   62
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label HoraRegPagto 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   8325
         TabIndex        =   61
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data\Hora:"
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
         Index           =   20
         Left            =   6285
         TabIndex        =   60
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.CommandButton BotaoConsultarContatos 
      Caption         =   "Consultar Contatos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7905
      TabIndex        =   6
      ToolTipText     =   "Lista todos contatos com o cliente ligados a fatura"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton BotaoAbrirFat 
      Caption         =   "Abrir Fatura .html"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6351
      TabIndex        =   5
      ToolTipText     =   "Abre o documento html com o detalhamento da fatura"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton BotaoGerarFat 
      Caption         =   "Regerar Fatura .html"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4797
      TabIndex        =   4
      ToolTipText     =   "Regera o documento html com o detalhamento da fatura"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton BotaoConsultarAp 
      Caption         =   "Consultar Aportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3243
      TabIndex        =   3
      ToolTipText     =   "Lista os aportes que fazem parte da fatura"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton BotaoConsultarItens 
      Caption         =   "Consultar Itens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1689
      TabIndex        =   2
      ToolTipText     =   "Lista os itens que fazem parte da fatura"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton BotaoConsultarDoc 
      Caption         =   "Consultar Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   1
      ToolTipText     =   "Abre a tela de consulta específica para o tipo de documento"
      Top             =   5385
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8310
      ScaleHeight     =   495
      ScaleWidth      =   1035
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   15
      Width           =   1095
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "TRVConsultaFatura.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "TRVConsultaFatura.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nota Fiscal"
      Height          =   1275
      Left            =   105
      TabIndex        =   26
      Top             =   3450
      Width           =   9315
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
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
         Index           =   18
         Left            =   2940
         TabIndex        =   58
         Top             =   225
         Width           =   435
      End
      Begin VB.Label ItemNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3390
         TabIndex        =   57
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Verif.:"
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
         Index           =   55
         Left            =   6750
         TabIndex        =   56
         Top             =   210
         Width           =   975
      End
      Begin VB.Label CodVerificacao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7755
         TabIndex        =   55
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NFe:"
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
         Index           =   56
         Left            =   4440
         TabIndex        =   54
         Top             =   240
         Width           =   420
      End
      Begin VB.Label NumNFe 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4890
         TabIndex        =   53
         Top             =   195
         Width           =   1275
      End
      Begin VB.Label ClienteNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   52
         Top             =   525
         Width           =   7935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Index           =   16
         Left            =   555
         TabIndex        =   51
         Top             =   555
         Width           =   660
      End
      Begin VB.Label EmissaoNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4905
         TabIndex        =   38
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label ValorNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7755
         TabIndex        =   37
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   19
         Left            =   7215
         TabIndex        =   36
         Top             =   915
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Index           =   1
         Left            =   4095
         TabIndex        =   35
         Top             =   915
         Width           =   765
      End
      Begin VB.Label FilialEmpresaNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   34
         Top             =   870
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   17
         Left            =   705
         TabIndex        =   33
         Top             =   900
         Width           =   465
      End
      Begin VB.Label NumNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   32
         Top             =   195
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Index           =   15
         Left            =   465
         TabIndex        =   31
         Top             =   225
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cancelamento"
      Height          =   570
      Left            =   105
      TabIndex        =   25
      Top             =   2850
      Width           =   9315
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
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
         Left            =   540
         TabIndex        =   68
         Top             =   225
         Width           =   645
      End
      Begin VB.Label MotivoCanc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   67
         Top             =   180
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data\Hora:"
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
         Index           =   14
         Left            =   6315
         TabIndex        =   50
         Top             =   225
         Width           =   975
      End
      Begin VB.Label HoraCanc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   8325
         TabIndex        =   49
         Top             =   180
         Width           =   900
      End
      Begin VB.Label DataCanc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   48
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label UsuarioCanc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4875
         TabIndex        =   47
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   4095
         TabIndex        =   46
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Geração da Fatura"
      Height          =   975
      Left            =   105
      TabIndex        =   24
      Top             =   1845
      Width           =   9315
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%:"
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
         Index           =   11
         Left            =   2895
         TabIndex        =   45
         Top             =   630
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data\Hora:"
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
         Index           =   13
         Left            =   6285
         TabIndex        =   44
         Top             =   630
         Width           =   975
      End
      Begin VB.Label HoraGerFat 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   8325
         TabIndex        =   43
         Top             =   585
         Width           =   900
      End
      Begin VB.Label DataGerFat 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7260
         TabIndex        =   42
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label PercAporte 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3105
         TabIndex        =   41
         Top             =   585
         Width           =   795
      End
      Begin VB.Label UsuarioGerFat 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4875
         TabIndex        =   40
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Index           =   12
         Left            =   4095
         TabIndex        =   39
         Top             =   630
         Width           =   720
      End
      Begin VB.Label ValorAporte 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   30
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Aporte:"
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
         Index           =   10
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label ClienteVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1260
         TabIndex        =   28
         Top             =   240
         Width           =   7965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Vou:"
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
         Index           =   9
         Left            =   165
         TabIndex        =   27
         Top             =   285
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fatura"
      Height          =   1260
      Left            =   105
      TabIndex        =   10
      Top             =   540
      Width           =   9315
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   1275
         TabIndex        =   0
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "#######"
         PromptChar      =   " "
      End
      Begin VB.Label SiglaDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7275
         TabIndex        =   69
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   23
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   22
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Index           =   0
         Left            =   4065
         TabIndex        =   21
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   7
         Left            =   765
         TabIndex        =   20
         Top             =   915
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   570
         TabIndex        =   19
         Top             =   555
         Width           =   660
      End
      Begin VB.Label LblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   8
         Left            =   6765
         TabIndex        =   17
         Top             =   900
         Width           =   510
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
         Height          =   195
         Index           =   6
         Left            =   6810
         TabIndex        =   16
         Top             =   210
         Width           =   450
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1275
         TabIndex        =   15
         Top             =   525
         Width           =   7965
      End
      Begin VB.Label TipoDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7995
         TabIndex        =   14
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   13
         Top             =   870
         Width           =   1950
      End
      Begin VB.Label FilialEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1275
         TabIndex        =   12
         Top             =   870
         Width           =   2160
      End
      Begin VB.Label Emissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4875
         TabIndex        =   11
         Top             =   870
         Width           =   1305
      End
   End
End
Attribute VB_Name = "TRVConsultaFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giTipoDocDestino As Integer
Dim gobjDestino As Object

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

'Variáveis globais
Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoNumero = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192673)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(Optional objFat As ClassFaturaTRV) As Long
    
    If Not (objFat Is Nothing) Then
        Numero.PromptInclude = False
        Numero.Text = CStr(objFat.lNumFat)
        Numero.PromptInclude = False
        Call Numero_Validate(bSGECancelDummy)
    End If
    
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objEventoNumero = Nothing
    'Libera os objetos e coleções globais

End Sub
'*** FECHAMENTO DA TELA - FIM ***

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_CancelarFatura

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192674)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Limpa_Tela_CancelarFatura()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CancelarFatura

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    TipoDoc.Caption = ""
    Status.Caption = ""
    Cliente.Caption = ""
    Emissao.Caption = ""
    Valor.Caption = ""
    FilialEmpresa.Caption = ""
    
    ClienteVou.Caption = ""
    ValorAporte.Caption = ""
    PercAporte.Caption = ""
    UsuarioGerFat.Caption = ""
    DataGerFat.Caption = ""
    HoraGerFat.Caption = ""
    DataPagto.Caption = ""
    SaldoPagto.Caption = ""
    DataRegPagto.Caption = ""
    HoraRegPagto.Caption = ""
    NumNF.Caption = ""
    NumNFe.Caption = ""
    CodVerificacao.Caption = ""
    ClienteNF.Caption = ""
    FilialEmpresaNF.Caption = ""
    EmissaoNF.Caption = ""
    ValorNF.Caption = ""
    ItemNF.Caption = ""
    MotivoCanc.Caption = ""
    UsuarioCanc.Caption = ""
    DataCanc.Caption = ""
    HoraCanc.Caption = ""
    SiglaDoc.Caption = ""
    
    giTipoDocDestino = 0
    Set gobjDestino = Nothing
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_CancelarFatura:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192688)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta de Faturas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVCancelarFatura"

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

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDestino As Object
Dim iTipoDocDestino As Integer
Dim objTitRec As ClassTituloReceber
Dim objTitPag As ClassTituloPagar
Dim objNFsPag As ClassNFsPag
Dim objcliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objFilialEmpresa As New AdmFiliais
Dim lNumero As Long
Dim iItemNF As Integer
Dim objNF As New ClassNFiscal
Dim objFaturaTRV As ClassFaturaTRV

On Error GoTo Erro_Numero_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Numero.ClipText)) <> 0 Then

        'Critica a Codigo
        lErro = Long_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 196501
        
        lNumero = StrParaLong(Numero.Text)
        
        lErro = CF("TRVFaturas_Le", lNumero, objDestino, iTipoDocDestino, True, True)
        If lErro <> SUCESSO Then gError 196502
               
    End If
               
    Call Limpa_Tela_CancelarFatura
    
    If lNumero <> 0 Then
        Numero.PromptInclude = False
        Numero.Text = CStr(lNumero)
        Numero.PromptInclude = True
    End If

    If Len(Trim(Numero.ClipText)) <> 0 Then

        giTipoDocDestino = iTipoDocDestino
        Set gobjDestino = objDestino

        Select Case iTipoDocDestino
        
            Case TRV_TIPO_DOC_DESTINO_TITREC
            
                Set objFaturaTRV = objDestino.objInfoUsu
            
                Set objTitRec = objDestino
                objcliente.lCodigo = objTitRec.lCliente
                
                lErro = CF("Cliente_Le", objcliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 196504
                
                objFilialEmpresa.iCodFilial = objTitRec.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196505
                            
                SiglaDoc.Caption = objTitRec.sSiglaDocumento
                TipoDoc.Caption = objFaturaTRV.sTipoDoc
                If objTitRec.iStatus = STATUS_BAIXADO Then
                    Status.Caption = "Baixado"
                ElseIf objTitRec.iStatus = STATUS_EXCLUIDO Then
                    Status.Caption = "Cancelado"
                Else
                    Status.Caption = "Aberto"
                End If
                Cliente.Caption = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido
                Emissao.Caption = Format(objTitRec.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objTitRec.dValor, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
            
                lErro = CF("TitulosRecTRV_Le_NF", objDestino.lNumIntDoc, objNF, iItemNF)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196765
                
                If lErro = SUCESSO Then
                
                    If objNF.lNumNotaFiscal <> 0 Then NumNF.Caption = CStr(objNF.lNumNotaFiscal)
                    If objNF.lNumNFe <> 0 Then NumNFe.Caption = CStr(objNF.lNumNFe)
                    CodVerificacao.Caption = objNF.sCodVerificacaoNFe
                    
                    Set objcliente = New ClassCliente
                    
                    objcliente.lCodigo = objNF.lCliente
                    
                    lErro = CF("Cliente_Le", objcliente)
                    If lErro <> SUCESSO And lErro <> 12293 Then gError 196504
                    
                    Set objFilialEmpresa = New AdmFiliais
                    
                    objFilialEmpresa.iCodFilial = objNF.iFilialEmpresa
                    
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 196505
                    
                    ClienteNF.Caption = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido
                    EmissaoNF.Caption = Format(objNF.dtDataEmissao, "dd/mm/yyyy")
                    ValorNF.Caption = Format(objNF.dValorTotal, "STANDARD")
                    FilialEmpresaNF.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
                    ItemNF.Caption = iItemNF
                    
                End If
                
                SaldoPagto.Caption = Format(objTitRec.dSaldo, "STANDARD")
            
            Case TRV_TIPO_DOC_DESTINO_TITPAG
            
                Set objFaturaTRV = objDestino.objInfoUsu
            
                Set objTitPag = objDestino
                objForn.lCodigo = objTitPag.lFornecedor
                
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 196506
                
                objFilialEmpresa.iCodFilial = objTitPag.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196507
                            
                SiglaDoc.Caption = objTitPag.sSiglaDocumento
                TipoDoc.Caption = objFaturaTRV.sTipoDoc
                If objTitPag.iStatus = STATUS_BAIXADO Then
                    Status.Caption = "Baixado"
                ElseIf objTitPag.iStatus = STATUS_EXCLUIDO Then
                    Status.Caption = "Cancelado"
                Else
                    Status.Caption = "Aberto"
                End If
                Cliente.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
                Emissao.Caption = Format(objTitPag.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objTitPag.dValorTotal, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
    
                SaldoPagto.Caption = Format(objTitPag.dSaldo, "STANDARD")
    
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
            
                Set objFaturaTRV = objDestino.objInfoUsu
                
                Set objNFsPag = objDestino
                objForn.lCodigo = objNFsPag.lFornecedor
                
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 196508
                
                objFilialEmpresa.iCodFilial = objNFsPag.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196509
                            
                SiglaDoc.Caption = "NC"
                TipoDoc.Caption = objFaturaTRV.sTipoDoc
            
                If objNFsPag.iStatus = STATUS_BAIXADO Then
                    Status.Caption = "Baixado"
                    SaldoPagto.Caption = Format(0, "STANDARD")
                ElseIf objNFsPag.iStatus = STATUS_EXCLUIDO Then
                    Status.Caption = "Cancelado"
                    SaldoPagto.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                Else
                    Status.Caption = "Aberto"
                    SaldoPagto.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                End If

                Cliente.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
                Emissao.Caption = Format(objNFsPag.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
                
                If objNFsPag.iStatus <> STATUS_EXCLUIDO Then
                    
                    If objNFsPag.lNumIntTitPag = 0 Then
                    
                        Status.Caption = "Aberto"
                        SaldoPagto.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                    
                    Else
                    
                        Set objTitPag = New ClassTituloPagar
                        
                        objTitPag.lNumIntDoc = objNFsPag.lNumIntTitPag
                        
                        'Lê em Títulos a Pagar
                        lErro = CF("TituloPagar_Le", objTitPag)
                        If lErro <> SUCESSO And lErro <> 18372 Then gError ERRO_SEM_MENSAGEM
                        
                        'Se não encontrou
                        If lErro = 18372 Then
                        
                            'Procura em Títulos a Receber Baixados
                            lErro = CF("TituloPagarBaixado_Le", objTitPag)
                            If lErro <> SUCESSO And lErro <> 56661 Then gError ERRO_SEM_MENSAGEM
    
                            Status.Caption = "Baixado"
                            SaldoPagto.Caption = Format(0, "STANDARD")
                            
                        Else
                            Status.Caption = "Aberto"
                            SaldoPagto.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                        End If
                        
                        If objTitPag.lNumTitulo <> 0 Then NumNF.Caption = CStr(objTitPag.lNumTitulo)
                        
                        Set objForn = New ClassFornecedor
                        
                        objForn.lCodigo = objTitPag.lFornecedor
                        
                        lErro = CF("Fornecedor_Le", objForn)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
                        
                        Set objFilialEmpresa = New AdmFiliais
                        
                        objFilialEmpresa.iCodFilial = objTitPag.iFilialEmpresa
                        
                        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        ClienteNF.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
                        EmissaoNF.Caption = Format(objTitPag.dtDataEmissao, "dd/mm/yyyy")
                        ValorNF.Caption = Format(objTitPag.dValorTotal, "STANDARD")
                        FilialEmpresaNF.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
                        ItemNF.Caption = ""
                        
                    End If
                    
                End If
                       
            Case Else
                gError 196503
            
        End Select
                  
        If objFaturaTRV.lClienteVou <> 0 And iTipoDocDestino <> TRV_TIPO_DOC_DESTINO_NFSPAG Then
        
            Set objcliente = New ClassCliente
            
            objcliente.lCodigo = objFaturaTRV.lClienteVou
            
            lErro = CF("Cliente_Le", objcliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 196504
            
            ClienteVou.Caption = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido
        
        End If
    
        ValorAporte.Caption = Format(objFaturaTRV.dValorAporte, "STANDARD")
        PercAporte.Caption = Format(objFaturaTRV.dPercAporte, "PERCENT")
        UsuarioGerFat.Caption = objFaturaTRV.sUsuarioGerFat
        If objFaturaTRV.dtDataGerFat <> DATA_NULA Then DataGerFat.Caption = Format(objFaturaTRV.dtDataGerFat, "dd/mm/yyyy")
        If objFaturaTRV.dHoraGerFat <> 0 Then HoraGerFat.Caption = Format(objFaturaTRV.dHoraGerFat, "hh:mm:ss")
        If objFaturaTRV.dtDataPagto <> DATA_NULA Then DataPagto.Caption = Format(objFaturaTRV.dtDataPagto, "dd/mm/yyyy")
        If objFaturaTRV.dtDataRegPagto <> DATA_NULA Then DataRegPagto.Caption = Format(objFaturaTRV.dtDataRegPagto, "dd/mm/yyyy")
        If objFaturaTRV.dHoraRegPagto <> 0 Then HoraRegPagto.Caption = Format(objFaturaTRV.dHoraRegPagto, "hh:mm:ss")
        MotivoCanc.Caption = objFaturaTRV.sMotivo
        UsuarioCanc.Caption = objFaturaTRV.sUsuarioCanc
        If objFaturaTRV.dtDataCanc <> DATA_NULA Then DataCanc.Caption = Format(objFaturaTRV.dtDataCanc, "dd/mm/yyyy")
        If objFaturaTRV.dHoraCanc <> 0 Then HoraCanc.Caption = Format(objFaturaTRV.dHoraCanc, "hh:mm:ss")
    
    End If
    
    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 196501, 196502, 196504, 196505, 196506, 196507, 196508, 196509, 196765
        
        Case 196503
            'Call Rotina_Erro(vbOKOnly, "ERRO_TRV_DESTINO_NAO_CADASTRADO", gErr, iTipoDocDestino)
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_DESTINO_NAO_CADASTRADO2", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196510)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultarDoc_Click()

Dim sTela As String

On Error GoTo Erro_BotaoConsultarDoc_Click

    Select Case giTipoDocDestino
    
        Case TRV_TIPO_DOC_DESTINO_CREDFORN
            sTela = TRV_TIPO_DOC_DESTINO_CREDFORN_TELA
            
        Case TRV_TIPO_DOC_DESTINO_DEBCLI
            sTela = TRV_TIPO_DOC_DESTINO_DEBCLI_TELA
    
        Case TRV_TIPO_DOC_DESTINO_TITPAG
            sTela = TRV_TIPO_DOC_DESTINO_TITPAG_TELA
    
        Case TRV_TIPO_DOC_DESTINO_TITREC
            sTela = TRV_TIPO_DOC_DESTINO_TITREC_TELA
    
        Case TRV_TIPO_DOC_DESTINO_NFSPAG
            sTela = TRV_TIPO_DOC_DESTINO_NFSPAG_TELA
    
    End Select
    
    If Not (gobjDestino Is Nothing) Then
            
        Call Chama_Tela(sTela, gobjDestino)
        
    End If

    Exit Sub

Erro_BotaoConsultarDoc_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196776)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultarItens_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoConsultarItens_Click

    If Not (gobjDestino Is Nothing) Then
    
        colSelecao.Add giTipoDocDestino
        colSelecao.Add gobjDestino.lNumIntDoc

        Call Chama_Tela("DocFaturadosLista", colSelecao, Nothing, Nothing, "TipoDocDestino = ? AND NumIntDocDestino = ?")

    End If
    
    Exit Sub

Erro_BotaoConsultarItens_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196775)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultarAp_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoConsultarAp_Click

    If Not (gobjDestino Is Nothing) Then
    
        colSelecao.Add gobjDestino.objInfoUsu.lNumFat

        Call Chama_Tela("TRVAPortesPagtoFatHistLista", colSelecao, Nothing, Nothing, "NumFat = ?")

    End If
    
    Exit Sub

Erro_BotaoConsultarAp_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196774)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerarFat_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sNomeDiretorio As String
Dim sModeloFat As String
Dim sSiglaDoc As String

On Error GoTo Erro_BotaoGerarFat_Click

    If Not (gobjDestino Is Nothing) Then
    
        GL_objMDIForm.MousePointer = vbHourglass
        
        lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sNomeDiretorio)
        If lErro <> SUCESSO Then gError 196770
        
        lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML, EMPRESA_TODA, sModeloFat)
        If lErro <> SUCESSO Then gError 196771
           
        lErro = CF("TRVFaturas_Regera_Html", gobjDestino.objInfoUsu.lNumFat, gobjDestino.objInfoUsu.lNumFat, sModeloFat, sNomeDiretorio, sSiglaDoc)
        If lErro <> SUCESSO Then gError 196772
        
        GL_objMDIForm.MousePointer = vbDefault

        Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")
    
    End If
    
    Exit Sub

Erro_BotaoGerarFat_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 196770 To 196772

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196773)

    End Select

    Exit Sub

End Sub

Public Sub BotaoAbrirFat_Click()

Dim lErro As Long
Dim sNomeArquivo As String
Dim sConteudo As String

On Error GoTo Erro_BotaoAbrirFat_Click

    If Not (gobjDestino Is Nothing) Then

        lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
        If lErro <> SUCESSO Then gError 194199
        
        If StrParaLong(Numero.ClipText) > 999999 Then
            sNomeArquivo = sConteudo & gsEmpresaTRVHTML & String(8 - Len(Numero.ClipText), "0") & CStr(Numero.ClipText) & ".html"
        Else
            sNomeArquivo = sConteudo & gsEmpresaTRVHTML & String(6 - Len(Numero.ClipText), "0") & CStr(Numero.ClipText) & ".html"
        End If
        
        Call Shell("explorer.exe " & sNomeArquivo, vbMaximizedFocus)
    
    End If
    
    Exit Sub

Erro_BotaoAbrirFat_Click:

    Select Case gErr
    
        Case 194199

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194200)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultarContatos_Click()

Dim colSelecao As New Collection
Dim objRelacionamentoCli As New ClassRelacClientes

On Error GoTo Erro_BotaoConsultarContatos_Click

    If Not (gobjDestino Is Nothing) Then
    
        If giTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC Then

            colSelecao.Add gobjDestino.lNumIntDoc
            
            Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, Nothing, "NumIntParcRec IN (SELECT NumIntDoc FROM ParcelasRecTodas WHERE NumIntTitulo = ?) ")

        End If

    End If
    
    Exit Sub

Erro_BotaoConsultarContatos_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194200)

    End Select

    Exit Sub

End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objFat As New ClassFaturaTRV
Dim colSelecao As New Collection

On Error GoTo Erro_LblNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Numero.Text)) <> 0 Then

        objFat.lNumFat = Numero.Text

    End If

    Call Chama_Tela("TRVFaturasLista", colSelecao, objFat, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197010)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFat As New ClassFaturaTRV

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objFat = obj1
    
    Numero.PromptInclude = False
    Numero.Text = CStr(objFat.lNumFat)
    Numero.PromptInclude = True

    Call Numero_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197009)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Numero Then Call LblNumero_Click
    
    End If
    
End Sub

