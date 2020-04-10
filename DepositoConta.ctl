VERSION 5.00
Begin VB.UserControl DepositoConta 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ScaleHeight     =   1560
   ScaleWidth      =   3510
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   1935
      Picture         =   "DepositoConta.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   630
      Picture         =   "DepositoConta.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.ComboBox CodContaCorrente 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   345
      Width           =   1695
   End
   Begin VB.Label LblConta 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
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
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   390
      Width           =   1350
   End
End
Attribute VB_Name = "DepositoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

