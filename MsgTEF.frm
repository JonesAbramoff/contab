VERSION 5.00
Begin VB.Form MsgTEF 
   Caption         =   "Mensagem TEF para o Operador"
   ClientHeight    =   1515
   ClientLeft      =   3060
   ClientTop       =   3450
   ClientWidth     =   4680
   Icon            =   "MsgTEF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Left            =   1605
      Top             =   615
   End
   Begin VB.Label LabelMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   4185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MsgTEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim giContador As Integer

Private Sub Form_Load()
    Timer1.Interval = 1000
    giContador = 1
End Sub

Private Sub Timer1_Timer()
    
    Me.Refresh
    giContador = giContador + 1
    
    If giContador > 4 Then Unload Me
    
End Sub
