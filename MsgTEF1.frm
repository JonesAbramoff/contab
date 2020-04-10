VERSION 5.00
Begin VB.Form MsgTEF1 
   Caption         =   "Mensagem TEF para o Operador"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1470
      Top             =   510
   End
   Begin VB.Label LabelMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   360
      TabIndex        =   0
      Top             =   135
      Width           =   4185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MsgTEF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim giContador As Integer

Private Sub Form_Load()
    Timer1.Interval = 1000
    giContador = 0
End Sub

Private Sub Timer1_Timer()
    
    giContador = giContador + 1
    
    If giContador > 4 Then Unload Me
    
End Sub

