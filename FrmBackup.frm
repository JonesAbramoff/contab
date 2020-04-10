VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aguarde"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Status 
      Height          =   495
      Left            =   135
      TabIndex        =   1
      Top             =   1680
      Width           =   4725
   End
   Begin VB.Label Label1 
      Caption         =   "Aguarde enquanto o backup agendado está sendo feito."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   4680
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Fechar()
    Unload Me
End Sub

Public Sub Abrir()
    Me.Show
    DoEvents
End Sub
