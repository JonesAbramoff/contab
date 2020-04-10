VERSION 5.00
Begin VB.Form Anotacoes 
   Caption         =   "Anotações"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3465
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   165
      Width           =   1650
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "Anotacoes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "Anotacoes.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "Anotacoes.frx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox Anotacao 
      Height          =   2955
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Anotacoes.frx":080A
      Top             =   915
      Width           =   4920
   End
   Begin VB.Label Label1 
      Caption         =   "Texto:"
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
      Left            =   255
      TabIndex        =   5
      Top             =   585
      Width           =   1260
   End
End
Attribute VB_Name = "Anotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Trata_Parametros(objFormAtiva As Object)

End Function
