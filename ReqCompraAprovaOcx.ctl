VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ReqCompraAprovaOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8340
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   675
      Visible         =   0   'False
      Width           =   16680
      Begin VB.Frame FrameItens 
         Caption         =   "Itens"
         Height          =   3420
         Left            =   60
         TabIndex        =   48
         Top             =   4845
         Width           =   16545
         Begin VB.TextBox ObsItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3960
            MaxLength       =   255
            TabIndex        =   55
            Top             =   1425
            Width           =   6795
         End
         Begin VB.TextBox DescProd 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1920
            MaxLength       =   255
            TabIndex        =   51
            Top             =   1065
            Width           =   4000
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   3015
            Left            =   60
            TabIndex        =   12
            Top             =   285
            Width           =   16350
            _ExtentX        =   28840
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   240
            Left            =   795
            TabIndex        =   50
            Top             =   1050
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UM 
            Height          =   225
            Left            =   4470
            TabIndex        =   52
            Top             =   1065
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Qtde 
            Height          =   225
            Left            =   5085
            TabIndex        =   53
            Top             =   1080
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLimC 
            Height          =   225
            Left            =   6285
            TabIndex        =   54
            Top             =   1035
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Requisições não enviadas"
         Height          =   4725
         Left            =   60
         TabIndex        =   33
         Top             =   15
         Width           =   16545
         Begin VB.CommandButton BotaoReqCompras 
            Caption         =   "Requisição de Compras..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7290
            TabIndex        =   11
            Top             =   4065
            Width           =   1830
         End
         Begin VB.CommandButton BotaoDesmarcarTodosReq 
            Caption         =   "Desmarcar Todos"
            Height          =   555
            Left            =   1680
            Picture         =   "ReqCompraAprovaOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4065
            Width           =   1425
         End
         Begin VB.CommandButton BotaoMarcarTodosReq 
            Caption         =   "Marcar Todos"
            Height          =   555
            Left            =   60
            Picture         =   "ReqCompraAprovaOcx.ctx":11E2
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4065
            Width           =   1425
         End
         Begin MSMask.MaskEdBox CodigoPV 
            Height          =   240
            Left            =   2475
            TabIndex        =   34
            Top             =   1500
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Requisitante 
            Height          =   240
            Left            =   6345
            TabIndex        =   30
            Top             =   1005
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.CheckBox Enviar 
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
            TabIndex        =   25
            Top             =   930
            Width           =   915
         End
         Begin VB.CheckBox Urgente 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5910
            TabIndex        =   29
            Top             =   975
            Width           =   735
         End
         Begin VB.TextBox ObsReq 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   31
            Top             =   1770
            Width           =   6795
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   3585
            TabIndex        =   27
            Top             =   975
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Req 
            Height          =   225
            Left            =   2745
            TabIndex        =   26
            Top             =   975
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   225
            Left            =   4755
            TabIndex        =   28
            Top             =   990
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRequisicoes 
            Height          =   705
            Left            =   60
            TabIndex        =   8
            Top             =   210
            Width           =   16290
            _ExtentX        =   28734
            _ExtentY        =   1244
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox CodigoOP 
            Height          =   240
            Left            =   4425
            TabIndex        =   49
            Top             =   1380
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8370
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   16650
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   3960
         Left            =   870
         TabIndex        =   35
         Top             =   390
         Width           =   7665
         Begin VB.Frame Frame9 
            Caption         =   "Número"
            Height          =   1425
            Left            =   4380
            TabIndex        =   45
            Top             =   2190
            Width           =   2385
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   315
               Left            =   780
               TabIndex        =   6
               Top             =   390
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   315
               Left            =   780
               TabIndex        =   7
               Top             =   960
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label Label14 
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
               Left            =   375
               TabIndex        =   47
               Top             =   450
               Width           =   315
            End
            Begin VB.Label Label12 
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
               Left            =   375
               TabIndex        =   46
               Top             =   1020
               Width           =   360
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data Limite - Compra"
            Height          =   1425
            Left            =   990
            TabIndex        =   42
            Top             =   2190
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataLimCDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   20
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataLimCAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   21
               Top             =   870
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimCDe 
               Height          =   300
               Left            =   735
               TabIndex        =   4
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataLimCAte 
               Height          =   300
               Left            =   720
               TabIndex        =   5
               Top             =   870
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label2 
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
               Left            =   285
               TabIndex        =   44
               Top             =   960
               Width           =   360
            End
            Begin VB.Label Label11 
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
               Left            =   255
               TabIndex        =   43
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data"
            Height          =   1425
            Left            =   990
            TabIndex        =   39
            Top             =   480
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   16
               Top             =   345
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   17
               Top             =   870
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   735
               TabIndex        =   0
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   720
               TabIndex        =   1
               Top             =   870
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label3 
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
               Left            =   285
               TabIndex        =   41
               Top             =   960
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
               Height          =   195
               Left            =   255
               TabIndex        =   40
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Data Limite - Entrega"
            Height          =   1425
            Left            =   4380
            TabIndex        =   36
            Top             =   480
            Width           =   2385
            Begin MSComCtl2.UpDown UpDownDataLimDe 
               Height          =   300
               Left            =   1905
               TabIndex        =   18
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataLimAte 
               Height          =   300
               Left            =   1890
               TabIndex        =   19
               Top             =   885
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataLimiteDe 
               Height          =   300
               Left            =   735
               TabIndex        =   2
               Top             =   360
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataLimiteAte 
               Height          =   300
               Left            =   720
               TabIndex        =   3
               Top             =   885
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label13 
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
               Left            =   285
               TabIndex        =   38
               Top             =   960
               Width           =   360
            End
            Begin VB.Label Label17 
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
               Left            =   255
               TabIndex        =   37
               Top             =   420
               Width           =   315
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   15225
      ScaleHeight     =   480
      ScaleWidth      =   1575
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   75
      Width           =   1635
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   45
         Picture         =   "ReqCompraAprovaOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "ReqCompraAprovaOcx.ctx":2356
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "ReqCompraAprovaOcx.ctx":2888
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8790
      Left            =   45
      TabIndex        =   22
      Top             =   330
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   15505
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requisições"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ReqCompraAprovaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer

Dim gobjReqCompraEnvio As ClassReqCompraEnvio

'GridRequisicoes
Dim objGridRequisicoes As AdmGrid
Dim iGrid_Enviar_Col As Integer
Dim iGrid_Req_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Urgente_Col As Integer
Dim iGrid_Requisitante_Col As Integer
Dim iGrid_CodigoOP_Col As Integer
Dim iGrid_ObsReq_Col As Integer
Dim iGrid_CodigoPV_Col As Integer

'GridItens
Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProd_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Qtde_Col As Integer
Dim iGrid_DataLimC_Col As Integer
Dim iGrid_ObsItem_Col As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Set objGridRequisicoes = New AdmGrid
    Set objGridItens = New AdmGrid
    Set gobjReqCompraEnvio = New ClassReqCompraEnvio

    'Inicializa o GridRequisicoes
    lErro = Inicializa_Grid_Requisicoes(objGridRequisicoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa o GridItens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211135)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    'libera as variaveis globais
    Set objGridRequisicoes = Nothing
    Set objGridItens = Nothing
    Set gobjReqCompraEnvio = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211136)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Limite Compra")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProd.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Qtde.Name)
    objGridInt.colCampo.Add (DataLimC.Name)
    objGridInt.colCampo.Add (ObsItem.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProd_Col = 2
    iGrid_UM_Col = 3
    iGrid_Qtde_Col = 4
    iGrid_DataLimC_Col = 5
    iGrid_ObsItem_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161194)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodosReq_Click()
    Call Grid_Marca_Desmarca(objGridRequisicoes, iGrid_Enviar_Col, DESMARCADO)
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Sub Limpa_Tela_ReqCompra()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ReqCompra

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridRequisicoes)
    Call Grid_Limpa(objGridItens)

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Limpa_Tela_ReqCompra:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211137)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela
    Call Limpa_Tela_ReqCompra

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 211138)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa o restante da tela
    Call Limpa_Tela_ReqCompra

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211139)

    End Select

    Exit Sub

End Sub

Private Sub CodigoAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodigoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)
End Sub

Private Sub CodigoDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodigoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)
End Sub

Private Sub DataAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub

Private Sub DataDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub DataLimiteAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataLimiteAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataLimiteAte, iAlterado)
End Sub

Private Sub DataLimiteDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataLimiteDe, iAlterado)
End Sub

Private Sub DataLimCAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataLimCAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataLimCAte, iAlterado)
End Sub

Private Sub DataLimCDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataLimCDe, iAlterado)
End Sub

Private Sub Enviar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub Enviar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRequisicoes)
End Sub

Private Sub Enviar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRequisicoes.objControle = Enviar
    lErro = Grid_Campo_Libera_Foco(objGridRequisicoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se o frame anterior foi o de Seleção e ele foi alterado
    If iFrameAtual <> 1 And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

        'Traz os dados das requisicoes e seus itens para a tela
        lErro = Traz_Requisicoes_Tela()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iFrameSelecaoAlterado = 0

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211140)

    End Select

    Exit Sub

End Sub

Private Function Traz_Requisicoes_Tela() As Long

Dim lErro As Long, lCodigoPV As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim iIndice As Integer, iLinha As Integer
Dim objRequisitante As ClassRequisitante

On Error GoTo Erro_Traz_Requisicoes_Tela

    lErro = Move_TabSelecao_Memoria()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ReqComprasEnvio_Le", gobjReqCompraEnvio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gobjReqCompraEnvio.colRequisicao.Count = 0 Then gError 211141
    
    Call Grid_Limpa(objGridRequisicoes)
    
    If gobjReqCompraEnvio.colRequisicao.Count >= objGridRequisicoes.objGrid.Rows Then
        Call Refaz_Grid(objGridRequisicoes, gobjReqCompraEnvio.colRequisicao.Count)
    End If
    
    iLinha = 0
    For Each objReqCompra In gobjReqCompraEnvio.colRequisicao

        iLinha = iLinha + 1

        GridRequisicoes.TextMatrix(iLinha, iGrid_Req_Col) = CStr(objReqCompra.lCodigo)

        'Verifica se DataLimite é diferente de Data Nula
        If objReqCompra.dtDataLimite <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_DataLimite_Col) = Format(objReqCompra.dtDataLimite, "dd/mm/yyyy")

        'Verifica se Data é diferente de Data Nula
        If objReqCompra.dtData <> DATA_NULA Then GridRequisicoes.TextMatrix(iLinha, iGrid_Data_Col) = Format(objReqCompra.dtData, "dd/mm/yyyy")

        GridRequisicoes.TextMatrix(iLinha, iGrid_Urgente_Col) = objReqCompra.lUrgente

        Set objRequisitante = New ClassRequisitante
        
        objRequisitante.lCodigo = objReqCompra.lRequisitante

        'Lê o requisitante
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError ERRO_SEM_MENSAGEM

        'Preenche o Requisitante com o código e o nome reduzido
        GridRequisicoes.TextMatrix(iLinha, iGrid_Requisitante_Col) = CStr(objRequisitante.lCodigo) & SEPARADOR & objRequisitante.sNome

        'Preenche a Observacao
        GridRequisicoes.TextMatrix(iLinha, iGrid_ObsReq_Col) = objReqCompra.sObservacao
                       
        If Len(Trim(objReqCompra.sOPCodigo)) > 0 Then

            lErro = Preenche_CodigoPV(objReqCompra, lCodigoPV)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            If lCodigoPV <> 0 Then
                GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoPV_Col) = CStr(lCodigoPV)
            End If
            
            GridRequisicoes.TextMatrix(iLinha, iGrid_CodigoOP_Col) = objReqCompra.sOPCodigo

        End If
    
    Next
    
    objGridRequisicoes.iLinhasExistentes = gobjReqCompraEnvio.colRequisicao.Count
    
    Call Grid_Refresh_Checkbox(objGridRequisicoes)
    
    Traz_Requisicoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Requisicoes_Tela:

    Traz_Requisicoes_Tela = gErr

    Select Case gErr
    
        Case 211141
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_INEXISTENTE", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211142)

    End Select

    Exit Function

End Function

Private Function Traz_ItensReq_Tela(ByVal iLinha As Integer) As Long

Dim lErro As Long, sProdMask As String
Dim objReqCompra As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim iIndice As Integer

On Error GoTo Erro_Traz_ItensReq_Tela

    FrameItens.Caption = "Itens"

    If objGridItens.iLinhasExistentes <> 0 Then Call Grid_Limpa(objGridItens)
    
    If Not (gobjReqCompraEnvio Is Nothing) Then

        If iLinha > 0 And iLinha <= gobjReqCompraEnvio.colRequisicao.Count Then
        
            Set objReqCompra = gobjReqCompraEnvio.colRequisicao.Item(iLinha)
        
            FrameItens.Caption = "Itens - " & CStr(objReqCompra.lCodigo)
        
            iIndice = 0
            For Each objItemRC In objReqCompra.colItens
                iIndice = iIndice + 1
               
                Call Mascara_RetornaProdutoTela(objItemRC.sProduto, sProdMask)
           
                GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdMask
                GridItens.TextMatrix(iIndice, iGrid_DescProd_Col) = objItemRC.sDescProduto
                If objReqCompra.dtDataLimite <> DATA_NULA Then
                    GridItens.TextMatrix(iIndice, iGrid_DataLimC_Col) = Format(DateAdd("d", -objItemRC.iTempoRessup, objReqCompra.dtDataLimite), "dd/mm/yyyy")
                End If
                GridItens.TextMatrix(iIndice, iGrid_ObsItem_Col) = objItemRC.sObservacao
                GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemRC.sUM
                GridItens.TextMatrix(iIndice, iGrid_Qtde_Col) = Formata_Estoque(objItemRC.dQuantidade)
            
            Next
            
            objGridItens.iLinhasExistentes = objReqCompra.colItens.Count
            
        End If
        
    End If
    
    Traz_ItensReq_Tela = SUCESSO

    Exit Function

Erro_Traz_ItensReq_Tela:

    Traz_ItensReq_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211143)

    End Select

    Exit Function
    
End Function

Function Move_TabSelecao_Memoria() As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    gobjReqCompraEnvio.dtDataDe = StrParaDate(DataDe.Text)
    gobjReqCompraEnvio.dtDataAte = StrParaDate(DataAte.Text)
    gobjReqCompraEnvio.dtDataLimiteDe = StrParaDate(DataLimiteDe.Text)
    gobjReqCompraEnvio.dtDataLimiteAte = StrParaDate(DataLimiteAte.Text)
    gobjReqCompraEnvio.dtDataLimCDe = StrParaDate(DataLimCDe.Text)
    gobjReqCompraEnvio.dtDataLimCAte = StrParaDate(DataLimCAte.Text)
    gobjReqCompraEnvio.lCodigoDe = StrParaLong(CodigoDe.Text)
    gobjReqCompraEnvio.lCodigoAte = StrParaLong(CodigoAte.Text)
    gobjReqCompraEnvio.iTipo = 2
    
    If gobjReqCompraEnvio.dtDataDe <> DATA_NULA And gobjReqCompraEnvio.dtDataAte <> DATA_NULA Then
        If gobjReqCompraEnvio.dtDataDe > gobjReqCompraEnvio.dtDataAte Then gError 211144
    End If
    If gobjReqCompraEnvio.dtDataLimiteDe <> DATA_NULA And gobjReqCompraEnvio.dtDataLimiteAte <> DATA_NULA Then
        If gobjReqCompraEnvio.dtDataLimiteDe > gobjReqCompraEnvio.dtDataLimiteAte Then gError 211145
    End If
    If gobjReqCompraEnvio.dtDataLimCDe <> DATA_NULA And gobjReqCompraEnvio.dtDataLimCAte <> DATA_NULA Then
        If gobjReqCompraEnvio.dtDataLimCDe > gobjReqCompraEnvio.dtDataLimCAte Then gError 211146
    End If
    If gobjReqCompraEnvio.lCodigoDe <> 0 And gobjReqCompraEnvio.lCodigoAte <> 0 Then
        If gobjReqCompraEnvio.lCodigoDe > gobjReqCompraEnvio.lCodigoAte Then gError 211147
    End If

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 211144
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 211145, 211146
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITEDE_MAIOR", gErr)

        Case 211147
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211148)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcarTodosReq_Click()
    Call Grid_Marca_Desmarca(objGridRequisicoes, iGrid_Enviar_Col, MARCADO)
End Sub

Private Sub GridRequisicoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRequisicoes, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
    
    Exit Sub

End Sub

Private Sub GridRequisicoes_GotFocus()
    Call Grid_Recebe_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_EnterCell()
    Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
End Sub

Private Sub GridRequisicoes_LeaveCell()
    Call Saida_Celula(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRequisicoes)
        
End Sub

Private Sub GridRequisicoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRequisicoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRequisicoes, iAlterado)
    End If
    
    Exit Sub
    
End Sub

Private Sub GridRequisicoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridRequisicoes)
End Sub

Private Sub GridRequisicoes_RowColChange()
    Call Grid_RowColChange(objGridRequisicoes)
    Call Traz_ItensReq_Tela(GridRequisicoes.Row)
End Sub

Private Sub GridRequisicoes_Scroll()
    Call Grid_Scroll(objGridRequisicoes)
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim lErro As Long

On Error GoTo Erro_GridItens_Click

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    Exit Sub

Erro_GridItens_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211149)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
        
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
        
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 211150

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 211150
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211151)

    End Select

    Exit Function

End Function

Private Sub UpDownDataAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimCAte_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimCDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimiteDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211152)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimiteAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211153)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimCDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimCDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimCDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimCDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211154)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimCAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimCAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataLimCAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimCAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211155)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211156)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a Data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211157)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimiteDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211158)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimiteAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211159)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimCDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimCDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimCDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimCDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211160)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimCAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataLimCAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataLimCAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataLimCAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211161)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211162)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211163)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Requisicoes(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Requisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Requisicoes

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Aprovar")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("OP")
    objGridInt.colColuna.Add ("PV")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("Requisitante")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (Enviar.Name)
    objGridInt.colCampo.Add (Req.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (CodigoOP.Name)
    objGridInt.colCampo.Add (CodigoPV.Name)
    objGridInt.colCampo.Add (Urgente.Name)
    objGridInt.colCampo.Add (Requisitante.Name)
    objGridInt.colCampo.Add (ObsReq.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Enviar_Col = 1
    iGrid_Req_Col = 2
    iGrid_Data_Col = 3
    iGrid_DataLimite_Col = 4
    iGrid_CodigoOP_Col = 5
    iGrid_CodigoPV_Col = 6
    iGrid_Urgente_Col = 7
    iGrid_Requisitante_Col = 8
    iGrid_ObsReq_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRequisicoes

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REQUISICOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12
    
    'Largura da primeira coluna
    GridRequisicoes.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Requisicoes = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Requisicoes:

    Inicializa_Grid_Requisicoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211164)

    End Select

    Exit Function

End Function

Private Sub BotaoReqCompras_Click()
'Chama a tela ReqComprasEnv

Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoReqCompras_Click

    'Verifica se existe alguma linha selecionada no GridRequisicoes
    If GridRequisicoes.Row = 0 Then gError 211165

    objRequisicaoCompras.lCodigo = StrParaLong(GridRequisicoes.TextMatrix(GridRequisicoes.Row, iGrid_Req_Col))
    objRequisicaoCompras.iFilialEmpresa = giFilialEmpresa

    'Chama a tela ReqComprasEnv
    Call Chama_Tela("ReqComprasCons", objRequisicaoCompras)

    Exit Sub

Erro_BotaoReqCompras_Click:

    Select Case gErr
    
        Case 211165
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211166)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataLimiteDe_Change()

    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteDe_Validate

    'Verifica se  DataLimiteDe foi preenchida
    If Len(Trim(DataLimiteDe.Text)) = 0 Then Exit Sub

    'Critica DataLimiteDe
    lErro = Data_Critica(DataLimiteDe.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataLimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211167)

    End Select

    Exit Sub

End Sub

Private Sub DataLimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimiteAte_Validate

    'Verifica se  DataLimiteAte foi preenchida
    If Len(Trim(DataLimiteAte.Text)) = 0 Then Exit Sub

    'Critica DataLimiteAte
    lErro = Data_Critica(DataLimiteAte.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataLimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211168)

    End Select

    Exit Sub

End Sub

Private Sub DataLimCDe_Change()
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataLimCDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimCDe_Validate

    'Verifica se  DataLimiteDe foi preenchida
    If Len(Trim(DataLimCDe.Text)) = 0 Then Exit Sub

    'Critica DataLimiteDe
    lErro = Data_Critica(DataLimCDe.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataLimCDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211169)

    End Select

    Exit Sub

End Sub

Private Sub DataLimCAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimCAte_Validate

    'Verifica se  DataLimiteAte foi preenchida
    If Len(Trim(DataLimCAte.Text)) = 0 Then Exit Sub

    'Critica DataLimiteAte
    lErro = Data_Critica(DataLimCAte.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataLimCAte_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211170)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se  DataDe foi preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica DataDe
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211171)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se  DataAte foi preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica DataAte
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211172)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Aprovação de Requisição para Compras"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ReqCompraAprova"

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

'**** fim do trecho a ser copiado *****

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Public Function Gravar_Registro() As Long
'Grava a Concorrencia

Dim lErro As Long
Dim objReqCompra As ClassRequisicaoCompras
Dim iCount As Integer, iLinha As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    iCount = 0
    iLinha = 0
    For Each objReqCompra In gobjReqCompraEnvio.colRequisicao
        iLinha = iLinha + 1
        If StrParaInt(GridRequisicoes.TextMatrix(iLinha, iGrid_Enviar_Col)) = MARCADO Then
            iCount = iCount + 1
            objReqCompra.iSelecionado = MARCADO
        Else
            objReqCompra.iSelecionado = DESMARCADO
        End If
    Next
    
    If iCount = 0 Then gError 211173
    
    lErro = CF("ReqCompraAprova_Grava", gobjReqCompraEnvio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 211173
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211174)

    End Select

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Function Preenche_CodigoPV(objRequisicaoCompra As ClassRequisicaoCompras, lCodigoPV As Long) As Long

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim iFilialPV As Integer

On Error GoTo Erro_Preenche_CodigoPV

    If Len(Trim(objRequisicaoCompra.sOPCodigo)) <> 0 Then
    
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objRequisicaoCompra.sOPCodigo
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30401 Then gError ERRO_SEM_MENSAGEM

        If lErro <> SUCESSO Then
        
            lErro = CF("ItensOP_Baixada_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 178689 Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        If lErro = SUCESSO Then
        
            For Each objItemOP In objOrdemProducao.colItens
                
                If objItemOP.lCodPedido <> 0 Then
                    lCodigoPV = objItemOP.lCodPedido
                    Exit For
                End If
                
                If objItemOP.lNumIntDocPai <> 0 Then
                
                    lErro = CF("ItensOP_Le_PV", objItemOP.lNumIntDocPai, lCodigoPV, iFilialPV)
                    If lErro <> SUCESSO And lErro <> 178696 And lErro <> 178697 Then gError ERRO_SEM_MENSAGEM
            
                End If
            
                If lCodigoPV <> 0 Then
                    Exit For
                End If
            
            Next
    
        End If
    
    End If

    Preenche_CodigoPV = SUCESSO
    
    Exit Function
    
Erro_Preenche_CodigoPV:

    Preenche_CodigoPV = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211175)

    End Select

    Exit Function

End Function

