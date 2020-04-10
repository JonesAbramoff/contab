VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl BorderoBoleto 
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   9060
   Begin VB.Frame FrameBoleto 
      BorderStyle     =   0  'None
      Caption         =   "Pagamentos em cartões não especificados"
      Height          =   2880
      Index           =   2
      Left            =   210
      TabIndex        =   30
      Top             =   2490
      Visible         =   0   'False
      Width           =   8505
      Begin VB.Frame Frame5 
         Caption         =   "Totais a Enviar"
         Height          =   975
         Left            =   4455
         TabIndex        =   46
         Top             =   1800
         Width           =   4005
         Begin VB.Label LabelTotalEnviarNCD 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2400
            TabIndex        =   50
            Top             =   570
            Width           =   1470
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cartão de Débito"
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
            Left            =   2400
            TabIndex        =   49
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label LabelTotalEnviarNCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   150
            TabIndex        =   48
            Top             =   570
            Width           =   1470
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cartão de Crédito"
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
            Left            =   150
            TabIndex        =   47
            Top             =   300
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Totais não enviados"
         Height          =   975
         Left            =   225
         TabIndex        =   41
         Top             =   1800
         Width           =   4005
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cartão de Débito"
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
            Left            =   2400
            TabIndex        =   45
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cartão de Crédito"
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
            Left            =   150
            TabIndex        =   44
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label LabelTotalNCC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   150
            TabIndex        =   43
            Top             =   570
            Width           =   1470
         End
         Begin VB.Label LabelTotalNCD 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2400
            TabIndex        =   42
            Top             =   570
            Width           =   1470
         End
      End
      Begin VB.ComboBox AdmCartaoN 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   2265
      End
      Begin VB.ComboBox ParcelamentoN 
         Height          =   315
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   660
         Width           =   2115
      End
      Begin MSMask.MaskEdBox ValorEnviarN 
         Height          =   270
         Left            =   5940
         TabIndex        =   33
         Top             =   900
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
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
      Begin MSFlexGridLib.MSFlexGrid GridBoletoN 
         Height          =   1515
         Left            =   225
         TabIndex        =   40
         Top             =   -15
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   2672
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrameBoleto 
      BorderStyle     =   0  'None
      Caption         =   "Valores cartões"
      Height          =   2880
      Index           =   1
      Left            =   270
      TabIndex        =   24
      Top             =   2430
      Width           =   8565
      Begin MSMask.MaskEdBox Parcelamento 
         Height          =   270
         Left            =   5910
         TabIndex        =   36
         Top             =   1740
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox AdmCartao 
         Height          =   270
         Left            =   3570
         TabIndex        =   35
         Top             =   1770
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorEnviar 
         Height          =   270
         Left            =   1980
         TabIndex        =   25
         Top             =   1770
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Standard"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Left            =   390
         TabIndex        =   26
         Top             =   1770
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
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
      Begin MSFlexGridLib.MSFlexGrid GridBoleto 
         Height          =   2370
         Left            =   300
         TabIndex        =   27
         Top             =   150
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   4180
         _Version        =   393216
         Rows            =   6
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelTotalEnviar 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6420
         TabIndex        =   39
         Top             =   2520
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4320
         TabIndex        =   38
         Top             =   2565
         Width           =   510
      End
      Begin VB.Label LabelTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4890
         TabIndex        =   37
         Top             =   2520
         Width           =   1470
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   4770
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   705
         Width           =   450
      End
   End
   Begin MSComCtl2.UpDown UpDownDataEnvio 
      Height          =   300
      Left            =   5910
      TabIndex        =   4
      Top             =   225
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2595
      Picture         =   "BorderoBoleto.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   570
      Left            =   6345
      ScaleHeight     =   510
      ScaleWidth      =   2535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2595
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   75
         Picture         =   "BorderoBoleto.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "BorderoBoleto.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1575
         Picture         =   "BorderoBoleto.ctx":0376
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2055
         Picture         =   "BorderoBoleto.ctx":08A8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   570
         Picture         =   "BorderoBoleto.ctx":0A26
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1395
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   645
      Width           =   8880
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7665
         TabIndex        =   12
         Top             =   570
         Width           =   1020
      End
      Begin VB.Frame Frame2 
         Caption         =   "Condições"
         Height          =   1140
         Left            =   3510
         TabIndex        =   21
         Top             =   165
         Width           =   3990
         Begin VB.OptionButton OptionAmbos 
            Caption         =   "Ambos"
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
            Left            =   270
            TabIndex        =   9
            Top             =   855
            Width           =   1350
         End
         Begin VB.OptionButton OptionParcelado 
            Caption         =   "Parcelado"
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
            Left            =   270
            TabIndex        =   8
            Top             =   570
            Width           =   1350
         End
         Begin VB.OptionButton OptionAVista 
            Caption         =   "À vista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   270
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.Frame Frame3 
            Caption         =   "Parcelado"
            Height          =   930
            Left            =   1710
            TabIndex        =   22
            Top             =   135
            Width           =   2025
            Begin VB.CheckBox CheckJurosLoja 
               BackColor       =   &H80000004&
               Caption         =   "Loja"
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
               Height          =   285
               Left            =   300
               TabIndex        =   10
               Top             =   225
               Width           =   1125
            End
            Begin VB.CheckBox CheckJurosAdm 
               Caption         =   "Administradora"
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
               Height          =   285
               Left            =   300
               TabIndex        =   11
               Top             =   570
               Width           =   1590
            End
         End
      End
      Begin VB.ComboBox AdmDefault 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   2355
      End
      Begin VB.ComboBox Rede 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   315
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartão:"
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   900
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rede:"
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
         Index           =   0
         Left            =   345
         TabIndex        =   19
         Top             =   360
         Width           =   525
      End
   End
   Begin MSMask.MaskEdBox DataEnvio 
      Height          =   300
      Left            =   4950
      TabIndex        =   3
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
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   225
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3315
      Left            =   180
      TabIndex        =   29
      Top             =   2085
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5847
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalhados"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Não Detalhados"
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
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   780
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   34
      Top             =   285
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Data de Envio:"
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
      Left            =   3615
      TabIndex        =   23
      Top             =   285
      Width           =   1290
   End
End
Attribute VB_Name = "BorderoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iAlterado As Integer
Dim giRedeAtual As Integer
Dim objGridBoleto As AdmGrid
Dim objGridBoletoN As AdmGrid
Dim gcolAdmMeioPagto As Collection
Dim iFrameAtual As Integer

Dim iGrid_AdmCartao_Col As Integer
Dim iGrid_Parcelamento_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_ValorEnviar_Col As Integer

Dim iGrid_AdmCartaoN_Col  As Integer
Dim iGrid_ParcelamentoN_Col  As Integer
Dim iGrid_ValorEnviarN_Col  As Integer

Private WithEvents objEventoBorderoBoleto As AdmEvento
Attribute objEventoBorderoBoleto.VB_VarHelpID = -1

Dim giDesativaBotaoTrazer As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub AdmDefault_Click()
    
    If Rede.ListIndex <> -1 Then Call BotaoTrazer_Click
        
End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objBordero As New ClassBorderoBoleto

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 120070

    'verifica se a data de envio está preenchida
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 120071

    'verifica se a rede está preenchida
    If Rede.ListIndex = -1 Then gError 120072
    
    Call Move_Tela_Memoria(objBordero)
    'If lErro <> SUCESSO Then gError 120073

    lErro = CF("BorderoBoleto_Le", objBordero)
    If lErro <> SUCESSO And lErro <> 107161 Then gError 120074

    If lErro = 107161 Then gError 120075
    
    '???? adaptar para bordero cheque
    'ver expr. selecao, nome tsk, etc..
    'aguardando tsk ficar pronto....
    'lErro = objRelatorio.ExecutarDireto("Borderô Boleto", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 120076

    'Limpa a Tela
    Call Limpa_Tela_BorderoBoleto

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 120070
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 120071
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 120072
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_SELECIONADA", gErr)
        
        Case 120073, 120074, 120076

        Case 120064
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROBOLETO_NAOENCONTRADO", gErr, objBordero.iFilialEmpresa, objBordero.lNumBordero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 143545)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Form_Load

    'instancia os grids globais
    Set objGridBoleto = New AdmGrid
    Set objGridBoletoN = New AdmGrid
    Set objEventoBorderoBoleto = New AdmEvento
    Set gcolAdmMeioPagto = New Collection

    iFrameAtual = 1

    'inicializa os grids
    Call Inicializa_GridBoleto(objGridBoleto)
    Call Inicializa_GridBoletoN(objGridBoletoN)

    'carrega a combo de redes
    lErro = Carrega_Rede_Cartao
    If lErro <> SUCESSO Then gError 107148

    'se estiver no bo a combo de redes é desativada
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then

        Label1(0).Enabled = False
        Rede.Enabled = False

    End If

    'preenche a data de envio com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa
 
    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107149

    'preenche o total da tela
    LabelTotalNCC.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa
    objTMPLojaFilial.dSaldo = 0
    
    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108222

    'preenche o total da tela
    LabelTotalNCD.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    LabelTotalEnviar.Caption = Format(0, "STANDARD")
    LabelTotalEnviarNCC.Caption = Format(0, "STANDARD")
    LabelTotalEnviarNCD.Caption = Format(0, "STANDARD")
    LabelTotal.Caption = Format(0, "STANDARD")
    
    'se estiver operando no bo-> desabilita o botaotrazer
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then BotaoTrazer.Enabled = False

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 107148, 107149, 108222

        Case 107150, 108223
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJAFILIAL_NAOENCONTRADO", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143546)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoBoleto As ClassBorderoBoleto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se objBorderoBoleto foi definido
    If Not objBorderoBoleto Is Nothing Then

        'traz o borderoboleto para a tela
        lErro = Traz_BorderoBoleto_Tela(objBorderoBoleto)
        If lErro <> SUCESSO And lErro <> 107168 Then gError 107173

        If lErro = 107168 Then

            'se não encontrou-> limpa a tela e preenche o código
            Call Limpa_Tela_BorderoBoleto
            Codigo.Text = objBorderoBoleto.lNumBordero

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 107173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143547)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a grava_registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107288

    'limpa a tela
    Call Limpa_Tela_BorderoBoleto

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143548)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objBorderoBoleto As New ClassBorderoBoleto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodigo_Click

    objBorderoBoleto.lNumBordero = StrParaLong(Codigo.Text)

    sSelecao = "ExibeTela = ?"

'    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
'        sSelecao = sSelecao & " AND DataBackOffice <> ?"
'    Else
'        sSelecao = sSelecao & " AND DataBackOffice = ?"
'    End If
'
    colSelecao.Add BOLETO_MANUAL
'    colSelecao.Add DATA_NULA

    Call Chama_Tela("BorderoBoletoLista", colSelecao, objBorderoBoleto, objEventoBorderoBoleto, sSelecao)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143549)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBorderoBoleto_evSelecao(obj1 As Object)

Dim objBorderoBoleto As ClassBorderoBoleto
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_objEventoBorderoBoleto_evSelecao

    'aponta para o objeto recebido por parâmetro
    Set objBorderoBoleto = obj1

    lErro = Traz_BorderoBoleto_Tela(objBorderoBoleto)
    If lErro <> SUCESSO Then gError 118006
    
    Me.Show

    Exit Sub

Erro_objEventoBorderoBoleto_evSelecao:

    Select Case gErr
        
        Case 118006
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143550)

    End Select

    Exit Sub

End Sub

Public Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim objBorderoBoleto As New ClassBorderoBoleto

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderoBoleto"

    'preenche o obj com os dados da tela
    Call Move_Tela_Memoria(objBorderoBoleto)

    'preenche a coleção de campos-valores
    colCampoValor.Add "FilialEmpresa", objBorderoBoleto.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumBordero", objBorderoBoleto.lNumBordero, 0, "NumBordero"
    colCampoValor.Add "Numero", objBorderoBoleto.sNumero, STRING_BORDEROBOLETO_NUMERO, "Numero"
    colCampoValor.Add "DataImpressao", objBorderoBoleto.dtDataImpressao, 0, "DataImpressao"
    colCampoValor.Add "DataEnvio", objBorderoBoleto.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataBackoffice", objBorderoBoleto.dtDataBackoffice, 0, "DataBackoffice"
    colCampoValor.Add "ExibeTela", objBorderoBoleto.iExibeTela, 0, "ExibeTela"

    'estabelece os filtros
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "ExibeTela", OP_IGUAL, BOLETO_MANUAL

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143551)

    End Select

    Exit Function

End Function

Public Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim objBorderoBoleto As New ClassBorderoBoleto
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'preenche os atributos necessários para chamar a traz tela
    objBorderoBoleto.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objBorderoBoleto.lNumBordero = colCampoValor.Item("NumBordero").vValor

    lErro = Traz_BorderoBoleto_Tela(objBorderoBoleto)
    If lErro <> SUCESSO And lErro <> 107168 Then gError 107172

    If lErro = 107168 Then gError 107174

    iAlterado = 0

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 107172

        Case 107174
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROBOLETO_NAOENCONTRADO", gErr, objBorderoBoleto.iFilialEmpresa, objBorderoBoleto.lNumBordero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143552)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'gera um código automático para o borderô
    lErro = BorderoBoleto_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 107176

    'preenche o código do borderô
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 107176

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143553)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'se não estiver preenchido-> sai
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'critica o campo
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 107177

    Cancel = False

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 107177

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143554)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)

End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvio_Validate

    'se a data não estiver preenchido-> sai da função
    If Len(Trim(DataEnvio.ClipText)) = 0 Then Exit Sub

    'critica a data
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then gError 107178

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True

    Select Case gErr

        Case 107178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143555)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Rede_Click()

Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Rede_Click

    'se não tem nada na rede selecionado->sai
    If Rede.ListIndex = -1 Then Exit Sub

    If Rede.ItemData(Rede.ListIndex) = giRedeAtual Then Exit Sub

    giRedeAtual = Rede.ItemData(Rede.ListIndex)

    Call GridBoleto_Limpa
    Call GridBoletoN_Limpa

    AdmDefault.Clear
    AdmCartaoN.Clear

    'verifica quais admmeiopagtos da coleção global existe pertence à rede selecionada
    For Each objAdmMeioPagto In gcolAdmMeioPagto

        'se pertencer à rede selecionada...
        If objAdmMeioPagto.iRede = Rede.ItemData(Rede.ListIndex) Then

            'adiciona a admmeiopagto nas combos
            AdmDefault.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
            AdmDefault.ItemData(AdmDefault.NewIndex) = objAdmMeioPagto.iCodigo
            AdmCartaoN.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
            AdmCartaoN.ItemData(AdmCartaoN.NewIndex) = objAdmMeioPagto.iCodigo

        End If

    Next

    Call BotaoTrazer_Click

    Exit Sub

Erro_Rede_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143556)

    End Select

    Exit Sub

End Sub

Private Sub Rede_GotFocus()

    If Rede.ListIndex <> -1 Then giRedeAtual = Rede.ItemData(Rede.ListIndex)

End Sub

Private Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'diminui a data de um dia
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 107179

    Exit Sub

Erro_UpDownDataEnvio_DownClick:

    Select Case gErr

        Case 107179

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143557)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'aumenta a data de um dia
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 107180

    Exit Sub

Erro_UpDownDataEnvio_UpClick:

    Select Case gErr

        Case 107180

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143558)

    End Select

    Exit Sub

End Sub

Private Sub OptionParcelado_Click()

    CheckJurosLoja.Enabled = True
    CheckJurosAdm.Enabled = True
    CheckJurosLoja.Value = MARCADO
    CheckJurosAdm.Value = MARCADO
    If Rede.ListIndex <> -1 Then Call BotaoTrazer_Click

End Sub

Private Sub OptionAVista_Click()

    CheckJurosLoja.Enabled = False
    CheckJurosAdm.Enabled = False
    If Rede.ListIndex <> -1 Then Call BotaoTrazer_Click

End Sub

Private Sub OptionAmbos_Click()

    CheckJurosLoja.Enabled = True
    CheckJurosAdm.Enabled = True
    CheckJurosLoja.Value = MARCADO
    CheckJurosAdm.Value = MARCADO
    If Rede.ListIndex <> -1 Then Call BotaoTrazer_Click

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameBoleto(TabStrip1.SelectedItem.index).Visible = True
        'Torna Frame atual visivel
        FrameBoleto(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.index

    End If

End Sub

Private Sub GridBoleto_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBoleto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridBoleto, iAlterado)
    End If

End Sub

Private Sub GridBoleto_EnterCell()

    Call Grid_Entrada_Celula(objGridBoleto, iAlterado)

End Sub

Private Sub GridBoleto_GotFocus()

    Call Grid_Recebe_Foco(objGridBoleto)

End Sub

Private Sub GridBoleto_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridBoleto)

End Sub

Private Sub GridBoleto_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBoleto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBoleto, iAlterado)
    End If

End Sub

Private Sub GridBoleto_LeaveCell()

    Call Saida_Celula(objGridBoleto)

End Sub

Private Sub GridBoleto_LostFocus()

    Call Grid_Libera_Foco(objGridBoleto)

End Sub

Private Sub GridBoleto_RowColChange()

    Call Grid_RowColChange(objGridBoleto)

End Sub

Private Sub GridBoleto_Scroll()

    Call Grid_Scroll(objGridBoleto)

End Sub

Private Sub GridBoleto_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBoleto)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelCodigo_Click
    End If

End Sub

Private Sub ValorEnviar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoleto)

End Sub

Private Sub ValorEnviar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoleto)

End Sub

Private Sub ValorEnviar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoleto.objControle = ValorEnviar
    lErro = Grid_Campo_Libera_Foco(objGridBoleto)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridBoletoN_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBoletoN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridBoletoN, iAlterado)
    End If

End Sub

Private Sub GridBoletoN_EnterCell()

    Call Grid_Entrada_Celula(objGridBoletoN, iAlterado)

End Sub

Private Sub GridBoletoN_GotFocus()

    Call Grid_Recebe_Foco(objGridBoletoN)

End Sub

Private Sub GridBoletoN_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iIndice As Integer
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim iTipoMeioPagtoLojaLinha As Integer
Dim dTotalCC As Double
Dim dTotalCD As Double

    Call Grid_Trata_Tecla1(KeyCode, objGridBoletoN)

    If KeyCode = vbKeyDelete Then

        For iIndice = 1 To objGridBoletoN.iLinhasExistentes

            'busca o tipomeiopagtoloja do cartão em questão
            For Each objAdmMeioPagto In gcolAdmMeioPagto
            
                'encontrando
                If objAdmMeioPagto.iCodigo = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice, iGrid_AdmCartaoN_Col)) Then
                    
                    'muda o indicador de tipomeiopagtoloja
                    iTipoMeioPagtoLojaLinha = objAdmMeioPagto.iTipoMeioPagto
                    Exit For
                
                End If
            
            Next
            
            'se for cartão de crédito, acumula o totalizador de cartão de crédito
            If iTipoMeioPagtoLojaLinha = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then dTotalCC = dTotalCC + StrParaDbl(GridBoletoN.TextMatrix(iIndice, iGrid_ValorEnviarN_Col))
            
            'se for cartão de débito, acumula o totalizador de cartão de débito
            If iTipoMeioPagtoLojaLinha = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then dTotalCD = dTotalCD + StrParaDbl(GridBoletoN.TextMatrix(iIndice, iGrid_ValorEnviarN_Col))
            
        Next

        LabelTotalEnviarNCC.Caption = Format(dTotalCC, "STANDARD")
        LabelTotalEnviarNCD.Caption = Format(dTotalCD, "STANDARD")

    End If

End Sub

Private Sub GridBoletoN_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBoletoN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBoletoN, iAlterado)
    End If



End Sub

Private Sub GridBoletoN_LeaveCell()

    Call Saida_Celula(objGridBoletoN)

End Sub

Private Sub GridBoletoN_LostFocus()

    Call Grid_Libera_Foco(objGridBoletoN)

End Sub

Private Sub GridBoletoN_RowColChange()

    Call Grid_RowColChange(objGridBoletoN)

End Sub

Private Sub GridBoletoN_Scroll()

    Call Grid_Scroll(objGridBoletoN)

End Sub

Private Sub GridBoletoN_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBoletoN)

End Sub

Private Sub ParcelamentoN_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelamentoN_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletoN)

End Sub

Private Sub ParcelamentoN_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletoN)

End Sub

Private Sub ParcelamentoN_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletoN.objControle = ParcelamentoN
    lErro = Grid_Campo_Libera_Foco(objGridBoletoN)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub AdmCartaoN_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AdmCartaoN_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletoN)

End Sub

Private Sub AdmCartaoN_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletoN)

End Sub

Private Sub AdmCartaoN_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletoN.objControle = AdmCartaoN
    lErro = Grid_Campo_Libera_Foco(objGridBoletoN)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorEnviarN_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviarN_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletoN)

End Sub

Private Sub ValorEnviarN_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletoN)

End Sub

Private Sub ValorEnviarN_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletoN.objControle = ValorEnviarN
    lErro = Grid_Campo_Libera_Foco(objGridBoletoN)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Name

            Case GridBoleto.Name
                lErro = Saida_Celula_GridBoleto(objGridInt)
                If lErro <> SUCESSO Then gError 107201

            Case GridBoletoN.Name
                lErro = Saida_Celula_GridBoletoN(objGridInt)
                If lErro <> SUCESSO Then gError 107202

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 107203

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 107201 To 107203

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143559)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridBoleto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridBoleto

    If objGridInt.objGrid.Col = iGrid_ValorEnviar_Col Then

            lErro = Saida_Celula_ValorEnviar(objGridInt)
            If lErro <> SUCESSO Then gError 107200

    End If

    Saida_Celula_GridBoleto = SUCESSO

    Exit Function

Erro_Saida_Celula_GridBoleto:

    Saida_Celula_GridBoleto = gErr

    Select Case gErr

        Case 107200

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143560)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorEnviar(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorEnviar

    Set objGridInt.objControle = ValorEnviar

    'se o valor estiver preenchido
    If Len(Trim(ValorEnviar.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ValorEnviar.Text))
        If lErro <> SUCESSO Then gError 107198

        'se a célula ultrapassar o valor da mesma linha-> erro
        If StrParaDbl(ValorEnviar.Text) > StrParaDbl(GridBoleto.TextMatrix(GridBoleto.Row, iGrid_Valor_Col)) Then gError 107197

    End If

    ValorEnviar.Text = Format(ValorEnviar.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107199

    dTotal = 0

    For iIndice = 1 To objGridBoleto.iLinhasExistentes

        dTotal = dTotal + StrParaDbl(GridBoleto.TextMatrix(iIndice, iGrid_ValorEnviar_Col))

    Next

    LabelTotalEnviar.Caption = Format(dTotal, "STANDARD")

    Saida_Celula_ValorEnviar = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorEnviar:

    Saida_Celula_ValorEnviar = gErr

    Select Case gErr

        Case 107197
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORENVIAR_MAIOR_VALORTOTAL", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 107198, 107199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143561)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridBoletoN(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridBoletoN

    Select Case objGridInt.objGrid.Col

        Case iGrid_AdmCartaoN_Col
            lErro = Saida_Celula_AdmCartaoN(objGridInt)
            If lErro <> SUCESSO Then gError 107194

        Case iGrid_ParcelamentoN_Col
            lErro = Saida_Celula_ParcelamentoN(objGridInt)
            If lErro <> SUCESSO Then gError 107195

        Case iGrid_ValorEnviarN_Col
            lErro = Saida_Celula_ValorEnviarN(objGridInt)
            If lErro <> SUCESSO Then gError 107196

    End Select

    Saida_Celula_GridBoletoN = SUCESSO

    Exit Function

Erro_Saida_Celula_GridBoletoN:

    Saida_Celula_GridBoletoN = gErr

    Select Case gErr

        Case 107194 To 107196

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143562)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AdmCartaoN(objGridInt As AdmGrid) As Long

Dim iCodigo As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_AdmCartaoN

    Set objGridInt.objControle = AdmCartaoN

    'verifica se a admcartao foi preenchido
    If Len(Trim(AdmCartaoN.Text)) <> 0 Then

        If GridBoletoN.Row - 1 = objGridBoletoN.iLinhasExistentes Then
            objGridBoletoN.iLinhasExistentes = objGridBoletoN.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107186

    Saida_Celula_AdmCartaoN = SUCESSO

    Exit Function

Erro_Saida_Celula_AdmCartaoN:

    Saida_Celula_AdmCartaoN = gErr

    Select Case gErr

        Case 107186
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143563)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcelamentoN(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_ParcelamentoN

    Set objGridInt.objControle = ParcelamentoN

    GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_ParcelamentoN_Col) = ParcelamentoN.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107190

    Saida_Celula_ParcelamentoN = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcelamentoN:

    Saida_Celula_ParcelamentoN = gErr

    Select Case gErr

        Case 107190
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143564)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorEnviarN(objGridInt As AdmGrid) As Long

Dim dTotalCC As Double
Dim dTotalCD As Double
Dim iIndice As Integer
Dim lErro As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim iTipoMeioPagtoLojaAtual As Integer
Dim iTipoMeioPagtoLojaLinha As Integer

On Error GoTo Erro_Saida_Celula_ValorEnviarN

    Set objGridInt.objControle = ValorEnviarN

    'busca o tipomeiopagtoloja do cartão em questão
    For Each objAdmMeioPagto In gcolAdmMeioPagto
    
        'encontrando
        If objAdmMeioPagto.iCodigo = Codigo_Extrai(GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_AdmCartaoN_Col)) Then
            
            'muda o indicador de tipomeiopagtoloja
            iTipoMeioPagtoLojaAtual = objAdmMeioPagto.iTipoMeioPagto
            Exit For
        
        End If
    
    Next
    
    'se o valor estiver preenchido
    If Len(Trim(ValorEnviarN.Text)) <> 0 Then

        For iIndice = 1 To objGridInt.iLinhasExistentes

            'busca o tipomeiopagtoloja do cartão em questão
            For Each objAdmMeioPagto In gcolAdmMeioPagto
            
                'encontrando
                If objAdmMeioPagto.iCodigo = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice, iGrid_AdmCartaoN_Col)) Then
                    
                    'muda o indicador de tipomeiopagtoloja
                    iTipoMeioPagtoLojaLinha = objAdmMeioPagto.iTipoMeioPagto
                    Exit For
                
                End If
            
            Next
            
            'acumula a soma de todas as linhas à exceção da linha atual
            If iIndice <> GridBoletoN.Row Then
                
                'se for cartão de crédito, acumula o totalizador de cartão de crédito
                If iTipoMeioPagtoLojaLinha = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then dTotalCC = dTotalCC + StrParaDbl(GridBoletoN.TextMatrix(iIndice, iGrid_ValorEnviarN_Col))
                
                'se for cartão de débito, acumula o totalizador de cartão de débito
                If iTipoMeioPagtoLojaLinha = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then dTotalCD = dTotalCD + StrParaDbl(GridBoletoN.TextMatrix(iIndice, iGrid_ValorEnviarN_Col))
            
            End If

        Next

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(ValorEnviarN.Text)
        If lErro <> SUCESSO Then gError 107192
        
        If iTipoMeioPagtoLojaAtual = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
            
            'se a célula ultrapassar o restante para o total-> erro
            If StrParaDbl(ValorEnviarN.Text) > StrParaDbl(LabelTotalNCC.Caption) - dTotalCC Then gError 107191
        
        End If
        
        If iTipoMeioPagtoLojaAtual = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then
        
            'se a célula ultrapassar o restante para o total-> erro
            If StrParaDbl(ValorEnviarN.Text) > StrParaDbl(LabelTotalNCD.Caption) - dTotalCD Then gError 108225
        
        End If
    
        ValorEnviarN.Text = Format(ValorEnviarN.Text, "STANDARD")
        
        If iTipoMeioPagtoLojaAtual = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then dTotalCC = dTotalCC + StrParaDbl(ValorEnviarN.Text)
        
        If iTipoMeioPagtoLojaAtual = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Then dTotalCD = dTotalCD + StrParaDbl(ValorEnviarN.Text)
        
        LabelTotalEnviarNCC.Caption = Format(dTotalCC, "STANDARD")
        LabelTotalEnviarNCD.Caption = Format(dTotalCD, "STANDARD")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107193

    Saida_Celula_ValorEnviarN = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorEnviarN:

    Saida_Celula_ValorEnviarN = gErr

    Select Case gErr

        Case 107191, 108225
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_LINHAS_MAIOR_TOTAL", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 107192, 107193
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143565)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, icaminho As Integer)

Dim sParcelamento As String
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case ParcelamentoN.Name

            'guarda o parcelamento atual
            sParcelamento = GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_ParcelamentoN_Col)

            If Len(Trim(GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_AdmCartaoN_Col))) = 0 Then
                objControl.Enabled = False
            Else

                objControl.Enabled = True

                'carrega um admmeiopagto com seus atributos chave
                objAdmMeioPagto.iCodigo = Codigo_Extrai(GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_AdmCartaoN_Col))
                objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa

                Call Carrega_Parcelamento(objAdmMeioPagto)

                'extrai o codigo do parcelamento armazenado anteriormente
                iCodigo = Codigo_Extrai(sParcelamento)

                'seleciona o item da combo com o codigo do parcelamento
                For iIndice = 0 To ParcelamentoN.ListCount - 1

                    If iCodigo = ParcelamentoN.ItemData(iIndice) Then

                        ParcelamentoN.ListIndex = iIndice
                        Exit For

                    End If

                Next

            End If

        Case ValorEnviarN.Name
            If Len(Trim(GridBoletoN.TextMatrix(GridBoletoN.Row, iGrid_AdmCartaoN_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143566)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()

Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iRede As Integer
Dim lErro As Long
Dim colAdmMeiopPagtoCondPagto As New Collection

On Error GoTo Erro_BotaoTrazer_Click

    If giDesativaBotaoTrazer = 0 Then

    'se não houver rede marcada-> erro
    If Rede.ListIndex = -1 Then gError 107227

    'se for marcado parcelado ou ambos...
    If OptionAmbos.Value = True Or OptionParcelado.Value = True Then

        '...e não tiver nenhuma das opções do tipo de parcelamento marcada-> erro
        If CheckJurosAdm.Value = DESMARCADO And CheckJurosLoja.Value = DESMARCADO Then gError 107228

        'se for juros de loja
        If CheckJurosLoja.Value = MARCADO Then objAdmMeioPagtoCondPagto.iJurosParcelamento = JUROSPARCELAMENTO_LOJA

        'se for juros da administradora
        If CheckJurosAdm.Value = MARCADO Then objAdmMeioPagtoCondPagto.iJurosParcelamento = objAdmMeioPagtoCondPagto.iJurosParcelamento + JUROSPARCELAMENTO_ADM

    End If

    'se admdefault estiver selecionada-> guardar o itemdata
    If AdmDefault.ListIndex <> -1 Then objAdmMeioPagtoCondPagto.iAdmMeioPagto = AdmDefault.ItemData(AdmDefault.ListIndex)

    'se for a vista
    If OptionAVista.Value = True Then objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AVISTA

    'se for parcelado
    If OptionParcelado.Value = True Then objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_PARCELADO

    'se for dos dois tipos
    If OptionAmbos.Value = True Then objAdmMeioPagtoCondPagto.iParcelamento = PARCELAMENTO_AMBOS

    'guarda o código da rede selecionada
    iRede = Rede.ItemData(Rede.ListIndex)

    'le as condpagto da adm em questao
    lErro = CF("AdmMeioPagtoCondPagto_Le1", iRede, objAdmMeioPagtoCondPagto, colAdmMeiopPagtoCondPagto)
    If lErro <> SUCESSO And lErro <> 107225 Then gError 107229

    If lErro = 107225 Then
        Call GridBoleto_Limpa
        gError 107230
    End If

    'preenche o grid
    Call GridBoleto_Preenche1(colAdmMeiopPagtoCondPagto)

    Call GridBoletoN_Limpa

    End If

    Exit Sub

Erro_BotaoTrazer_Click:

    Select Case gErr

        Case 107227
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_SELECIONADA", gErr)

        Case 107228
            Call Rotina_Erro(vbOKOnly, "ERRO_JUROSPARCELAMENTO_NAO_INFORMADO", gErr)

        Case 107229

        Case 107230
            Call Rotina_Aviso(vbOKOnly, "AVISO_RESULTADO_BUSCA_VAZIO")

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143567)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim objBorderoBoleto As New ClassBorderoBoleto
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se estiver no BO-> erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 107289

    'verifica se o código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107283

    'verifica se a data de envio está preenchida
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 107284

    'verifica se a rede está preenchida
    If Rede.ListIndex = -1 Then gError 107285

    'verifica se há alguma coisa para enviar
    If (StrParaDbl(LabelTotalEnviar.Caption) + StrParaDbl(LabelTotalEnviarNCC.Caption) + StrParaDbl(LabelTotalEnviarNCD.Caption)) = 0 Then gError 107286

    Call Move_Tela_Memoria(objBorderoBoleto)

    'verifica se está alterando
    lErro = Trata_Alteracao(objBorderoBoleto, objBorderoBoleto.iFilialEmpresa, objBorderoBoleto.lNumBordero)
    If lErro <> SUCESSO Then gError 107290

    'preenche a coleção de itens
    Call Move_Tela_Memoria_Boleto(objBorderoBoleto.colBorderoBoletoItem)

    'grava o borderoboleto
    lErro = CF("BorderoBoleto_Grava", objBorderoBoleto)
    If lErro <> SUCESSO Then gError 107312
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 107283
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 107284
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 107285
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_SELECIONADA", gErr)

        Case 107286
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTALENVIAR_ZERADO", gErr)

        Case 107290, 107312

        Case 107289
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_BORDEROBOLETO_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143568)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objBorderoBoleto As New ClassBorderoBoleto

On Error GoTo Erro_BotaoExcluir_Click

    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 107375
    
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_BORDEROBOLETO", giFilialEmpresa, Codigo.Text)
    
    If vbMsgResp = vbYes Then
    
        'preenche os atributos chave do borderoboleto
        objBorderoBoleto.lNumBordero = StrParaLong(Codigo.Text)
        objBorderoBoleto.iFilialEmpresa = giFilialEmpresa
        
        'chama a função de exclusão
        lErro = CF("BorderoBoleto_Exclui", objBorderoBoleto)
        If lErro <> SUCESSO Then gError 107387
        
        Call Limpa_Tela_BorderoBoleto
        
        iAlterado = 0
        
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:
    
    Select Case gErr
    
        Case 107375, 107387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143569)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim objBorderoBoleto As New ClassBorderoBoleto

On Error GoTo Erro_Botaolimpar_Click

    'testa se houve alteração
    lErro = Teste_Salva(objBorderoBoleto, iAlterado)
    If lErro <> SUCESSO Then gError 107231

    'limpa a tela
    Call Limpa_Tela_BorderoBoleto

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 107232

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 107231

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143570)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoBoleto()

Dim objTMPLojaFilial As New ClassTMPLojaFilial
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_BorderoBoleto

    giRedeAtual = 0

    Call Limpa_Tela(Me)

    Call GridBoleto_Limpa

    Call GridBoletoN_Limpa

    Rede.ListIndex = -1

    AdmDefault.Clear
    AdmCartaoN.Clear

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107165

    'preenche o total da tela
    LabelTotalNCC.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108224

    'preenche o total da tela
    LabelTotalNCD.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    'preenche a data de envio com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_BorderoBoleto:

    Select Case gErr

        Case 107165, 108224

        Case 107166, 108225
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJAFILIAL_NAOENCONTRADO", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143571)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

    'libera a memória
    Set objEventoBorderoBoleto = Nothing
    Set objGridBoleto = Nothing
    Set objGridBoletoN = Nothing
    Set gcolAdmMeioPagto = Nothing

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Carrega_Parcelamento(objAdmMeioPagto As ClassAdmMeioPagto)

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_Parcelamento

    'lê os tipos de parcelamento da admmeiopagto
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 107184

    'se não encontrar->erro
    If lErro = 104086 Then gError 107185

    'limpa a combo
    ParcelamentoN.Clear

    'preencher a combo de parcelamento
    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja

        ParcelamentoN.AddItem (objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento)
        ParcelamentoN.ItemData(ParcelamentoN.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

    Next

    Exit Sub

Erro_Carrega_Parcelamento:

    Select Case gErr

        Case 107184

        Case 107185
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_NAOENCONTRADO", gErr, objAdmMeioPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143572)

    End Select

    Exit Sub

End Sub

Private Function Traz_BorderoBoleto_Tela(objBorderoBoleto As ClassBorderoBoleto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBorderoBoletoItem As New ClassBorderoBoletoItem

On Error GoTo Erro_Traz_BorderoBoleto_Tela

    giDesativaBotaoTrazer = 1

    'limpa a tela
    Call Limpa_Tela_BorderoBoleto

    'busca o bordero no BD
    lErro = CF("BorderoBoleto_Le", objBorderoBoleto)
    If lErro <> SUCESSO And lErro <> 107161 Then gError 107167

    'se nãou encontrou-> erro
    If lErro = 107161 Then gError 107168
        
    'preenche a tela
    LabelTotalEnviar.Caption = Format(objBorderoBoleto.dValorEnviar, "STANDARD")
    LabelTotalEnviarNCC.Caption = Format(objBorderoBoleto.dValorEnviarNCC, "STANDARD")
    LabelTotalEnviarNCD.Caption = Format(objBorderoBoleto.dValorEnviarNCD, "STANDARD")
    Codigo.Text = objBorderoBoleto.lNumBordero

    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(objBorderoBoleto.dtDataEnvio, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    For iIndice = 0 To Rede.ListCount - 1

        If Rede.ItemData(iIndice) = objBorderoBoleto.iCodigoRede Then

            Rede.ListIndex = iIndice
            Exit For

        End If

    Next
    
    Set objBorderoBoletoItem = objBorderoBoleto.colBorderoBoletoItem.Item(1)
    
    If iIndice = Rede.ListCount Then gError 107310
    
    For iIndice = 0 To AdmDefault.ListCount - 1

        If AdmDefault.ItemData(iIndice) = objBorderoBoletoItem.iAdmMeioPagto Then

            AdmDefault.ListIndex = iIndice
            Exit For

        End If

    Next
   
    'preenche o grid de boletos
    Call GridBoleto_Preenche(objBorderoBoleto.colBorderoBoletoItem)

    giDesativaBotaoTrazer = 0

    Traz_BorderoBoleto_Tela = SUCESSO

    Exit Function

Erro_Traz_BorderoBoleto_Tela:

    giDesativaBotaoTrazer = 0

    Traz_BorderoBoleto_Tela = gErr

    Select Case gErr

        Case 107167, 107168

        Case 107310
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", objBorderoBoleto.iCodigoRede)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143573)

    End Select

    Exit Function

End Function

Private Function Carrega_Rede_Cartao() As Long

Dim colRedes As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim lErro As Long
Dim iIndice As Integer
Dim objRede As ClassRede

On Error GoTo Erro_Carrega_Rede_Cartao

    'lê as redes da tabela
    lErro = CF("Redes_Le_Todas", colRedes)
    If lErro <> SUCESSO Then gError 107145

    'se a coleção retornar vazia ->erro
    If colRedes.Count = 0 Then gError 107146

    'carrega a coleção de admmeiopagto cujo tipo=cartão de crédito
    lErro = CF("AdmMeioPagto_Le_TipoMeioPagto", TIPOMEIOPAGTOLOJA_CARTAO_CREDITO, gcolAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 107360 Then gError 107143
    
    'carrega a coleção de admmeiopagto cujo tipo=cartão de débito
    lErro = CF("AdmMeioPagto_Le_TipoMeioPagto", TIPOMEIOPAGTOLOJA_CARTAO_DEBITO, gcolAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 107360 Then gError 108220
    
    'Para cada elemento da Coleção
    For Each objAdmMeioPagto In gcolAdmMeioPagto

        'busca o codigo da admmeiopagto na combo de redes
        For iIndice = 0 To Rede.ListCount - 1

            If objAdmMeioPagto.iRede = Rede.ItemData(iIndice) Then Exit For

        Next

        'se não achou, é para adicionar
        If iIndice = Rede.ListCount Then

            'zera o indicador de sucesso na busca
            iIndice = 0

            'busca a rede o admmeiopagto na coleção de redes
            For Each objRede In colRedes

                'se encontrou, adiciona na combo os dados devidos
                If objAdmMeioPagto.iRede = objRede.iCodigo Then

                    Rede.AddItem (objRede.sNome)
                    Rede.ItemData(Rede.NewIndex) = objRede.iCodigo
                    Exit For

                End If

                iIndice = iIndice + 1

            Next

            'se o indice for igual ao contador da coleção de redes, não achou->erro
            If iIndice = colRedes.Count Then gError 107147

        End If

    Next

    Carrega_Rede_Cartao = SUCESSO

    Exit Function

Erro_Carrega_Rede_Cartao:

    Carrega_Rede_Cartao = gErr

    Select Case gErr

        Case 107143, 107145, 108220

        Case 107144, 108221
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_VAZIA", gErr)

        Case 107146
            Call Rotina_Erro(vbOKOnly, "ERRO_REDES_VAZIA", gErr)

        Case 107147
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, objAdmMeioPagto.iRede)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143574)

    End Select

    Exit Function

End Function

Private Sub Inicializa_GridBoleto(objGridInt As AdmGrid)

On Error GoTo Erro_Inicializa_GridBoleto

    'form do grid
    Set objGridInt.objForm = Me

    'Títulos das Colunas
    objGridInt.colColuna.Add ""
    objGridInt.colColuna.Add "Cartão"
    objGridInt.colColuna.Add "Parcelamento"
    objGridInt.colColuna.Add "Saldo"
    objGridInt.colColuna.Add "Valor a Enviar"

    'Controles que participam do Grid
    objGridInt.colCampo.Add AdmCartao.Name
    objGridInt.colCampo.Add Parcelamento.Name
    objGridInt.colCampo.Add Valor.Name
    objGridInt.colCampo.Add ValorEnviar.Name

    'Colunas do Grid
    iGrid_AdmCartao_Col = 1
    iGrid_Parcelamento_Col = 2
    iGrid_Valor_Col = 3
    iGrid_ValorEnviar_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridBoleto

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 10

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridBoleto.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'proibe incluir excluir linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Exit Sub

Erro_Inicializa_GridBoleto:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143575)

    End Select

    Exit Sub

End Sub

Private Sub Inicializa_GridBoletoN(objGridInt As AdmGrid)

On Error GoTo Erro_Inicializa_GridBoletoN

    'form do grid
    Set objGridInt.objForm = Me

    'Títulos das Colunas
    objGridInt.colColuna.Add ""
    objGridInt.colColuna.Add "Cartão"
    objGridInt.colColuna.Add "Parcelamento"
    objGridInt.colColuna.Add "Valor a Enviar"

    'Controles que participam do Grid
    objGridInt.colCampo.Add AdmCartaoN.Name
    objGridInt.colCampo.Add ParcelamentoN.Name
    objGridInt.colCampo.Add ValorEnviarN.Name

    'Colunas do Grid
    iGrid_AdmCartaoN_Col = 1
    iGrid_ParcelamentoN_Col = 2
    iGrid_ValorEnviarN_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridBoletoN

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 10

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridBoletoN.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'executa rotina enable para o grid
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Exit Sub

Erro_Inicializa_GridBoletoN:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143576)

    End Select

    Exit Sub

End Sub

Private Sub GridBoleto_Preenche(colBorderoBoletoItem As Collection)

Dim objBorderoBoletoItem As ClassBorderoBoletoItem
Dim dValorTotal As Double
Dim dValorEnviarTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_GridBoleto_Preenche

    'se a quantidade de itens da coleção superar a quantidade de linhas do grid
    If colBorderoBoletoItem.Count > 10 Then

        'acerta a quantidade de linhas do grid
        GridBoleto.Rows = colBorderoBoletoItem.Count + 1

        'reinicializa o grid
        Call Grid_Inicializa(objGridBoleto)

    End If

    iIndice = 1

    'preenche as linhas do grid
    For Each objBorderoBoletoItem In colBorderoBoletoItem

        'preenche as colunas do grid
        GridBoleto.TextMatrix(iIndice, iGrid_AdmCartao_Col) = objBorderoBoletoItem.iAdmMeioPagto & SEPARADOR & objBorderoBoletoItem.sNomeAdmMeioPagto
        GridBoleto.TextMatrix(iIndice, iGrid_Parcelamento_Col) = objBorderoBoletoItem.iParcelamento & SEPARADOR & objBorderoBoletoItem.sNomeParcelamento
        GridBoleto.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objBorderoBoletoItem.dSaldo + objBorderoBoletoItem.dValor, "STANDARD")
        GridBoleto.TextMatrix(iIndice, iGrid_ValorEnviar_Col) = Format(objBorderoBoletoItem.dValor, "STANDARD")

        dValorTotal = dValorTotal + objBorderoBoletoItem.dSaldo + objBorderoBoletoItem.dValor
        dValorEnviarTotal = dValorEnviarTotal + objBorderoBoletoItem.dValor
        iIndice = iIndice + 1

    Next

    'atualiza a quantidade de linhas existentes no grid
    objGridBoleto.iLinhasExistentes = colBorderoBoletoItem.Count

    'atualiza as checks do grid
    lErro = Grid_Refresh_Checkbox(objGridBoleto)
    If lErro <> SUCESSO Then gError 107169

    'preenche os totalizadores da tela
    LabelTotal.Caption = Format(dValorTotal, "STANDARD")
    LabelTotalEnviar.Caption = Format(dValorEnviarTotal, "STANDARD")

    Exit Sub

Erro_GridBoleto_Preenche:

    Select Case gErr

        Case 107169

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143577)

    End Select

    Exit Sub

End Sub

Private Sub GridBoleto_Preenche1(colAdmMeioPagtoCondPagto As Collection)

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice(0 To 1) As Integer
Dim dValorTotal As Double
Dim dValorEnviarTotal As Double
Dim vbMsgResp As VbMsgBoxResult

On Error GoTo Erro_GridBoleto_Preenche1

    'se já houver linhas no grid
    If objGridBoleto.iLinhasExistentes <> 0 Then

        'pergunta se deseja limpá-lo
        vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_LIMPAR_GRID")

        'se sim
        If vbMsgResp = vbYes Then

            'limpa o grid
            Call GridBoleto_Limpa

        'se não
        Else

            'varre o grid...
            For iIndice(0) = 1 To objGridBoleto.iLinhasExistentes

                iIndice(1) = 1

                '... verificando qual objeto da coleção já está no grid
                For Each objAdmMeioPagtoCondPagto In colAdmMeioPagtoCondPagto

                    'se ele for encontrado no grid
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = Codigo_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_AdmCartao_Col)) And objAdmMeioPagtoCondPagto.iParcelamento = Codigo_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_Parcelamento_Col)) Then

                        'o remove da coleção
                        colAdmMeioPagtoCondPagto.Remove (iIndice(1))
                        Exit For

                    End If

                    iIndice(1) = iIndice(1) + 1

                Next

                dValorEnviarTotal = dValorEnviarTotal + StrParaDbl(GridBoleto.TextMatrix(iIndice(0), iGrid_ValorEnviar_Col))

            Next

        End If

    End If

    If colAdmMeioPagtoCondPagto.Count + objGridBoleto.iLinhasExistentes > 10 Then

        GridBoleto.Rows = colAdmMeioPagtoCondPagto.Count + objGridBoleto.iLinhasExistentes + 1
        Call Grid_Inicializa(objGridBoleto)

    End If

    iIndice(0) = objGridBoleto.iLinhasExistentes + 1

    'preenche as linhas do grid
    For Each objAdmMeioPagtoCondPagto In colAdmMeioPagtoCondPagto

        'preenche as colunas do grid
        GridBoleto.TextMatrix(iIndice(0), iGrid_AdmCartao_Col) = objAdmMeioPagtoCondPagto.iAdmMeioPagto & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
        GridBoleto.TextMatrix(iIndice(0), iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
        GridBoleto.TextMatrix(iIndice(0), iGrid_Valor_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "STANDARD")

        iIndice(0) = iIndice(0) + 1

        dValorTotal = dValorTotal + objAdmMeioPagtoCondPagto.dSaldo

    Next

    LabelTotal.Caption = Format(dValorTotal + StrParaDbl(LabelTotal.Caption), "STANDARD")
    LabelTotalEnviar = Format(dValorEnviarTotal, "STANDARD")

    objGridBoleto.iLinhasExistentes = objGridBoleto.iLinhasExistentes + colAdmMeioPagtoCondPagto.Count

    Exit Sub

Erro_GridBoleto_Preenche1:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143578)

    End Select

    Exit Sub

End Sub

Private Sub GridBoleto_Limpa()

On Error GoTo Erro_GridBoleto_Limpa

    'limpa o grid
    Call Grid_Limpa(objGridBoleto)

    'limpa os totalizadores
    LabelTotal.Caption = Format(0, "STANDARD")
    LabelTotalEnviar.Caption = Format(0, "STANDARD")

    Exit Sub

Erro_GridBoleto_Limpa:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143579)

    End Select

    Exit Sub

End Sub

Private Sub GridBoletoN_Limpa()

Dim objTMPLojaFilial As New ClassTMPLojaFilial
Dim lErro As Long

On Error GoTo Erro_GridBoletoN_Limpa

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107228

    'preenche o total da tela
    LabelTotalNCC.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")

    'preenche a chave de busca para um meio de pagamento
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa
    objTMPLojaFilial.dSaldo = 0

    'busca o saldo do meio de pagamento
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108226

    'preenche o total da tela
    LabelTotalNCD.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    'limpa o grid
    Call Grid_Limpa(objGridBoletoN)

    'limpa os totalizadores
    LabelTotalEnviarNCC.Caption = Format(0, "STANDARD")
    LabelTotalEnviarNCD.Caption = Format(0, "STANDARD")

    Exit Sub

Erro_GridBoletoN_Limpa:

    Select Case gErr
        
        Case 108226, 108228

        Case 108227, 107291
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJAFILIAL_NAOENCONTRADO", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143580)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria(objBorderoBoleto As ClassBorderoBoleto)

On Error GoTo Erro_Move_Tela_Memoria

    'move os dados para a memória
    objBorderoBoleto.dtDataEnvio = StrParaDate(DataEnvio.Text)
    objBorderoBoleto.dValorEnviar = StrParaDbl(LabelTotalEnviar.Caption)
    objBorderoBoleto.dValorEnviarNCC = StrParaDbl(LabelTotalEnviarNCC.Caption)
    objBorderoBoleto.dValorEnviarNCD = StrParaDbl(LabelTotalEnviarNCD.Caption)
    objBorderoBoleto.iFilialEmpresa = giFilialEmpresa
    objBorderoBoleto.lNumBordero = StrParaLong(Codigo.Text)
    objBorderoBoleto.dtDataBackoffice = DATA_NULA
    objBorderoBoleto.dtDataImpressao = DATA_NULA
    If Rede.ListIndex <> -1 Then
        objBorderoBoleto.iCodigoRede = Rede.ItemData(Rede.ListIndex)
    End If
    
    Exit Sub

Erro_Move_Tela_Memoria:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143581)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria_Boleto(colBorderoBoletoItem As Collection)

Dim iAdmMeioPagto As Integer
Dim iParcelamento As Integer
Dim objBorderoBoletoItem As ClassBorderoBoletoItem
Dim iIndice(0 To 1) As Integer

On Error GoTo Erro_Move_Tela_Memoria_Boleto

    'varre o grid de boletos detalhados
    For iIndice(0) = 1 To objGridBoleto.iLinhasExistentes

        'se o valor estiver preenchido
        If Len(Trim(GridBoleto.TextMatrix(iIndice(0), iGrid_ValorEnviar_Col))) <> 0 Then

            'cria um novo item de borderoboleto
            Set objBorderoBoletoItem = New ClassBorderoBoletoItem

            'preenche seus atributos
            objBorderoBoletoItem.dSaldo = StrParaDbl(GridBoleto.TextMatrix(iIndice(0), iGrid_Valor_Col))
            objBorderoBoletoItem.dValor = StrParaDbl(GridBoleto.TextMatrix(iIndice(0), iGrid_ValorEnviar_Col))
            objBorderoBoletoItem.iAdmMeioPagto = Codigo_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_AdmCartao_Col))
            objBorderoBoletoItem.sNomeAdmMeioPagto = Nome_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_AdmCartao_Col))
            objBorderoBoletoItem.iFilialEmpresa = giFilialEmpresa
            objBorderoBoletoItem.iParcelamento = Codigo_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_Parcelamento_Col))
            objBorderoBoletoItem.lNumBordero = StrParaLong(Codigo.Text)
            objBorderoBoletoItem.sNomeParcelamento = Nome_Extrai(GridBoleto.TextMatrix(iIndice(0), iGrid_Parcelamento_Col))

            'adiciona à coleção
            colBorderoBoletoItem.Add objBorderoBoletoItem

        End If

    Next

    'varre o grid de boletos não detalhados
    For iIndice(0) = 1 To objGridBoletoN.iLinhasExistentes

        'se o valor estiver preenchido
        If Len(Trim(GridBoletoN.TextMatrix(iIndice(0), iGrid_ValorEnviarN_Col))) <> 0 Then

            'inicia o indexador da coleção de itens de borderoboleto
            iIndice(1) = 1

            'varre a coleção de borderoboletoitem...
            For Each objBorderoBoletoItem In colBorderoBoletoItem

                '... em busca do borderoboletoitem
                iAdmMeioPagto = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_AdmCartaoN_Col))
                iParcelamento = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_ParcelamentoN_Col))

                'se encontrar-> sai do loop
                If iAdmMeioPagto = objBorderoBoletoItem.iAdmMeioPagto And iParcelamento = objBorderoBoletoItem.iParcelamento Then Exit For

                'incrementa o indexador da coleção
                iIndice(1) = iIndice(1) + 1

            Next

            'se não encontrou(iIndice(1)= colborderoBoletoitem.count)
            If iIndice(1) > colBorderoBoletoItem.Count Then

                'cria um novo item de bordero
                Set objBorderoBoletoItem = New ClassBorderoBoletoItem

                'preenche os atributos do item
                objBorderoBoletoItem.dValorN = StrParaDbl(GridBoletoN.TextMatrix(iIndice(0), iGrid_ValorEnviarN_Col))
                objBorderoBoletoItem.iAdmMeioPagto = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_AdmCartaoN_Col))
                objBorderoBoletoItem.iFilialEmpresa = giFilialEmpresa
                objBorderoBoletoItem.iParcelamento = Codigo_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_ParcelamentoN_Col))
                objBorderoBoletoItem.lNumBordero = StrParaLong(Codigo.Text)
                objBorderoBoletoItem.sNomeParcelamento = Nome_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_ParcelamentoN_Col))
                objBorderoBoletoItem.sNomeAdmMeioPagto = Nome_Extrai(GridBoletoN.TextMatrix(iIndice(0), iGrid_AdmCartaoN_Col))
                colBorderoBoletoItem.Add objBorderoBoletoItem

            'se encontrou
            Else

                'aponta para o item da coleção
                Set objBorderoBoletoItem = colBorderoBoletoItem.Item(iIndice(1))

                'atualiza seu saldo
                objBorderoBoletoItem.dValorN = objBorderoBoletoItem.dValorN + StrParaDbl(GridBoletoN.TextMatrix(iIndice(0), iGrid_ValorEnviarN_Col))

            End If

        End If

    Next

    Exit Sub

Erro_Move_Tela_Memoria_Boleto:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143582)

    End Select

    Exit Sub

End Sub

Private Function BorderoBoleto_Codigo_Automatico(lCodigo As Long) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoBoleto_Codigo_Automatico

    'gera um número automático para o borderoboleto
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_BORDEROBOLETO", "BorderoBoleto", "NumBordero", lCodigo)
    If lErro <> SUCESSO Then gError 107175

    BorderoBoleto_Codigo_Automatico = SUCESSO

    Exit Function

Erro_BorderoBoleto_Codigo_Automatico:

    BorderoBoleto_Codigo_Automatico = gErr

    Select Case gErr

        Case 107175

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143583)

    End Select

    Exit Function

End Function


Private Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Trim(Mid(sTexto, iPosicao + 1))

    Nome_Extrai = sString

    Exit Function

End Function



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Borderô Boleto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoBoleto"

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

Private Sub Parcelado_Click()
'''    JurosAdm.Enabled = True
'''    JurosLoja.Enabled = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
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
