VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ImportaDadosFisc 
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   2565
   ScaleWidth      =   4950
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
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
      Left            =   1185
      Picture         =   "ImportaDadosFisc.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
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
      Left            =   2595
      Picture         =   "ImportaDadosFisc.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox CheckSobrepoe 
      Caption         =   "Sobrepõe alterações feitas anteriormente neste módulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   7
      Top             =   1170
      Width           =   4995
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   300
      TabIndex        =   0
      Top             =   225
      Width           =   4380
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   315
         Left            =   1815
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicio 
         Height          =   300
         Left            =   810
         TabIndex        =   2
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   315
         Left            =   3735
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFim 
         Height          =   300
         Left            =   2730
         TabIndex        =   4
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fim:"
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
         Left            =   2295
         TabIndex        =   5
         Top             =   285
         Width           =   360
      End
   End
End
Attribute VB_Name = "ImportaDadosFisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'o default p/data inicial é obtido a partir de FISConfig FISC_DATA_FIM_IMPORT
'gera registros de entrada e saida tabelas LivRegES*
'se vai sobrepor registros entrados anteriormente...

'se a atualizacao é on-line nao permitir ativar esta rotina

'Tabela: LivRegES

'Semelahante o que é feito na geração do Arquivo de ICMS.

'IO
'Fazer select nas Notas Fiscais
'Filtro: NFiscal.FilialEmpresa = giFilialEmpresa AND ((TiposDocInfo.ModeloArqICMS = MODELO_ARQ_ICMS_NFISCAL AND NFiscal.DataEmissao >= DataInicial AND NFiscal.DataEmissao <= DataFinal) OR (TiposDocInfo.ModeloArqICMS = MODELO_ARQ_ICMS_NFISCAL_ENTRADA AND NFiscal.DataEntrada >= DataInicial AND NFiscal.DataEntrada <= DataFinal))
'Para cada Registro Lido tem que

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

