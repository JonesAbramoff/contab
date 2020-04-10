VERSION 5.00
Begin VB.Form OpcoesGerais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções Gerais"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "RelOpGerais.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   3825
      Picture         =   "RelOpGerais.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   510
      Left            =   3810
      Picture         =   "RelOpGerais.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Imprimir Folha de Rosto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   180
      TabIndex        =   0
      Top             =   105
      Width           =   2955
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Guardar resultado para reimpressão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1215
      Width           =   3630
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Somente fontes da impressora"
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
      Index           =   2
      Left            =   180
      TabIndex        =   1
      Top             =   470
      Width           =   3150
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Não exibir gráficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   805
      Width           =   1935
   End
End
Attribute VB_Name = "OpcoesGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

