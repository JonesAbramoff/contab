VERSION 5.00
Begin VB.UserControl NFiscalFaturaConsultaOcx 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1470
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Caption         =   $"NFiscalFaturaConsultaOcx.ctx":0000
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4440
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "NFiscalFaturaConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property
