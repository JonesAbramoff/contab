VERSION 5.00
Begin VB.UserControl TributoFilial 
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   ScaleHeight     =   4935
   ScaleWidth      =   7005
   Begin VB.Frame Frame3 
      Caption         =   "Escrituração dos Livros"
      Height          =   1620
      Left            =   300
      TabIndex        =   15
      Top             =   3075
      Width           =   4740
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   255
         Width           =   1905
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periodo Atual"
         Height          =   765
         Left            =   180
         TabIndex        =   16
         Top             =   690
         Width           =   4230
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   2790
            TabIndex        =   24
            Top             =   225
            Width           =   1245
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   810
            TabIndex        =   23
            Top             =   255
            Width           =   1245
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   18
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2235
            TabIndex        =   17
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Periodicidade:"
         Height          =   210
         Left            =   210
         TabIndex        =   20
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4725
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TributoFilial.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TributoFilial.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TributoFilial.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TributoFilial.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apuração"
      Height          =   1635
      Left            =   300
      TabIndex        =   4
      Top             =   1335
      Width           =   4740
      Begin VB.Frame Frame2 
         Caption         =   "Periodo Atual"
         Height          =   795
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   4230
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   2715
            TabIndex        =   22
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   795
            TabIndex        =   21
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2205
            TabIndex        =   9
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   285
            Width           =   570
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   255
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   "Periodicidade:"
         Height          =   210
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   945
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   330
      Width           =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição:"
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   870
      Width           =   780
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   975
      TabIndex        =   2
      Top             =   810
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Tributo:"
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "TributoFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'carga:
'obtem dados da tabela de tributos (defaults de periodicidade) e tributofilial e atualiza tributofilial e

Private Sub Label7_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label7(Index), Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7(Index), Button, Shift, X, Y)
End Sub


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

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

