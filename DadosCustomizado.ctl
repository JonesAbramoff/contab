VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl DadosCustomizados 
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4005
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   Begin VB.TextBox Texto 
      Height          =   315
      Index           =   5
      Left            =   3945
      MaxLength       =   255
      TabIndex        =   47
      Top             =   1530
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox Texto 
      Height          =   315
      Index           =   4
      Left            =   3945
      MaxLength       =   255
      TabIndex        =   46
      Top             =   1173
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox Texto 
      Height          =   315
      Index           =   2
      Left            =   3945
      MaxLength       =   255
      TabIndex        =   45
      Top             =   461
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox Texto 
      Height          =   315
      Index           =   3
      Left            =   3945
      MaxLength       =   255
      TabIndex        =   44
      Top             =   817
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox Texto 
      Height          =   315
      Index           =   1
      Left            =   3945
      MaxLength       =   255
      TabIndex        =   43
      Top             =   105
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.CommandButton BotaoDadosCustDel 
      Height          =   405
      Left            =   6720
      Picture         =   "DadosCustomizado.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3525
      Width           =   435
   End
   Begin VB.CommandButton BotaoDadosCustNovo 
      Height          =   405
      Left            =   6225
      Picture         =   "DadosCustomizado.ctx":04B6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3525
      Width           =   435
   End
   Begin VB.ComboBox Controles 
      Height          =   315
      ItemData        =   "DadosCustomizado.ctx":09C8
      Left            =   7185
      List            =   "DadosCustomizado.ctx":09D8
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3585
      Width           =   1770
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Index           =   1
      Left            =   2430
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Index           =   2
      Left            =   2430
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Index           =   3
      Left            =   2430
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   9
      Top             =   810
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Index           =   4
      Left            =   2430
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Index           =   4
      Left            =   1260
      TabIndex        =   12
      Top             =   1185
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Index           =   5
      Left            =   2430
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Index           =   5
      Left            =   1260
      TabIndex        =   15
      Top             =   1545
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   315
      Index           =   1
      Left            =   1260
      TabIndex        =   22
      Top             =   2010
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   315
      Index           =   2
      Left            =   1260
      TabIndex        =   24
      Top             =   2370
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   26
      Top             =   2730
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   315
      Index           =   4
      Left            =   1260
      TabIndex        =   28
      Top             =   3090
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   315
      Index           =   5
      Left            =   1260
      TabIndex        =   30
      Top             =   3465
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Index           =   1
      Left            =   3945
      TabIndex        =   32
      Top             =   2010
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Index           =   2
      Left            =   3945
      TabIndex        =   34
      Top             =   2370
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Index           =   3
      Left            =   3945
      TabIndex        =   36
      Top             =   2730
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Index           =   4
      Left            =   3945
      TabIndex        =   38
      Top             =   3090
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Index           =   5
      Left            =   3945
      TabIndex        =   40
      Top             =   3465
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor5:"
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
      Index           =   2005
      Left            =   2595
      TabIndex        =   41
      Top             =   3555
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor4:"
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
      Index           =   2004
      Left            =   2595
      TabIndex        =   39
      Top             =   3195
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor3:"
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
      Index           =   2003
      Left            =   2595
      TabIndex        =   37
      Top             =   2835
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor2:"
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
      Index           =   2002
      Left            =   2595
      TabIndex        =   35
      Top             =   2445
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor1:"
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
      Index           =   2001
      Left            =   2595
      TabIndex        =   33
      Top             =   2070
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número5:"
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
      Height          =   285
      Index           =   3005
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   31
      Top             =   3510
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número4:"
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
      Height          =   285
      Index           =   3004
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   3135
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número3:"
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
      Height          =   285
      Index           =   3003
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   27
      Top             =   2775
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número2:"
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
      Height          =   285
      Index           =   3002
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   25
      Top             =   2415
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número1:"
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
      Height          =   285
      Index           =   3001
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   23
      Top             =   2055
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Texto5:"
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
      Index           =   4005
      Left            =   3210
      TabIndex        =   21
      Top             =   1590
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Texto4:"
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
      Index           =   4004
      Left            =   3210
      TabIndex        =   20
      Top             =   1230
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Texto3:"
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
      Index           =   4003
      Left            =   3210
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Texto2:"
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
      Index           =   4002
      Left            =   3210
      TabIndex        =   18
      Top             =   510
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Texto1:"
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
      Index           =   4001
      Left            =   2895
      TabIndex        =   17
      Top             =   150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data5:"
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
      Index           =   1005
      Left            =   120
      TabIndex        =   16
      Top             =   1590
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data4:"
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
      Index           =   1004
      Left            =   120
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data3:"
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
      Index           =   1003
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data2:"
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
      Index           =   1002
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data1:"
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
      Index           =   1001
      Left            =   105
      TabIndex        =   2
      Top             =   165
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "DadosCustomizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub BotaoNovo_Click()

End Sub

Private Sub Data_Change(Index As Integer)
'
End Sub

Private Sub Data_GotFocus(Index As Integer)
'
End Sub

Private Sub Data_Validate(Index As Integer, Cancel As Boolean)
'
End Sub
