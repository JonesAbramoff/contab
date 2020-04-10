VERSION 5.00
Begin VB.UserControl NFiscalConsultaOcx 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   ScaleHeight     =   1950
   ScaleWidth      =   5190
   Begin VB.Label Label1 
      Caption         =   "Esta tela é semelhante a tela de nota fiscal com a diferença que todos os campos serão labels. "
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4440
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Vai ter o campo data de saida como editável. Deverá estar presente no primeiro tabstrip"
      Height          =   525
      Left            =   165
      TabIndex        =   0
      Top             =   1095
      Width           =   4755
   End
End
Attribute VB_Name = "NFiscalConsultaOcx"
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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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
