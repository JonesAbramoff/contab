VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl AcompanhamentoSRVOcx 
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   ScaleHeight     =   4785
   ScaleWidth      =   8235
   Begin VB.ComboBox Contato 
      Height          =   315
      Left            =   1170
      TabIndex        =   8
      ToolTipText     =   $"AcompanhamentoSRVOcx.ctx":0000
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Assunto 
      Height          =   1140
      Left            =   105
      MaxLength       =   510
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Width           =   8055
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1725
      Picture         =   "AcompanhamentoSRVOcx.ctx":0088
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Pressione esse botão para gerar um código automático para o relacionamento."
      Top             =   1230
      Width           =   300
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "AcompanhamentoSRVOcx.ctx":0172
      Left            =   4140
      List            =   "AcompanhamentoSRVOcx.ctx":0174
      TabIndex        =   7
      Text            =   "Tipo"
      ToolTipText     =   "Selecione o tipo de relacionamento com o cliente. Para cadastrar novos tipos, use a tela Campos Genéricos."
      Top             =   1725
      Width           =   4005
   End
   Begin VB.ComboBox Atendente 
      Height          =   315
      Left            =   1185
      TabIndex        =   6
      ToolTipText     =   "Digite o código, o nome do atendente ou aperte F3 para consulta. Para cadastrar novos tipos, use a tela Campos Genéricos."
      Top             =   1710
      Width           =   2040
   End
   Begin VB.TextBox CodigoOS 
      Height          =   285
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   0
      Top             =   255
      Width           =   1350
   End
   Begin VB.CheckBox ImprimeGravacao 
      Caption         =   "Imprimir ao gravar"
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
      Left            =   3285
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5355
      ScaleHeight     =   450
      ScaleWidth      =   2685
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   105
      Width           =   2745
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   120
         Picture         =   "AcompanhamentoSRVOcx.ctx":0176
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   630
         Picture         =   "AcompanhamentoSRVOcx.ctx":0278
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1140
         Picture         =   "AcompanhamentoSRVOcx.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1650
         Picture         =   "AcompanhamentoSRVOcx.ctx":055C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "AcompanhamentoSRVOcx.ctx":0A8E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   5145
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1170
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   4155
      TabIndex        =   4
      ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
      Top             =   1200
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1185
      TabIndex        =   3
      ToolTipText     =   "Informe o código do relacionamento."
      Top             =   1200
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Hora 
      Height          =   315
      Left            =   7230
      TabIndex        =   5
      Top             =   1200
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1185
      TabIndex        =   2
      Top             =   750
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label FilialLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4155
      TabIndex        =   31
      Top             =   2265
      Width           =   1800
   End
   Begin VB.Label ClienteLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1170
      TabIndex        =   30
      Top             =   2265
      Width           =   1800
   End
   Begin VB.Label DescricaoProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3030
      TabIndex        =   29
      Top             =   750
      Width           =   5085
   End
   Begin VB.Label ProdutoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Serviço:"
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
      Left            =   330
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   28
      Top             =   825
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contato:"
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
      Left            =   360
      TabIndex        =   27
      Top             =   2820
      Width           =   735
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   26
      Top             =   2310
      Width           =   660
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   25
      Top             =   2325
      Width           =   465
   End
   Begin VB.Label LabelAssunto 
      AutoSize        =   -1  'True
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
      Left            =   135
      TabIndex        =   24
      Top             =   3225
      Width           =   750
   End
   Begin VB.Label LabelSeq 
      AutoSize        =   -1  'True
      Caption         =   "Sequencial:"
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
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   23
      Top             =   1260
      Width           =   1020
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   3585
      TabIndex        =   22
      Top             =   1245
      Width           =   480
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   3630
      TabIndex        =   21
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora:"
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
      Height          =   195
      Index           =   3
      Left            =   6675
      TabIndex        =   20
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Atendente:"
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
      Index           =   4
      Left            =   150
      TabIndex        =   19
      Top             =   1785
      Width           =   945
   End
   Begin VB.Label CodigoOSLabel 
      AutoSize        =   -1  'True
      Caption         =   " O.S.:"
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
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   300
      Width           =   510
   End
End
Attribute VB_Name = "AcompanhamentoSRVOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub LabelCliente_Click()

End Sub

