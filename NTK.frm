VERSION 5.00
Begin VB.Form NTK 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NTK 
      Caption         =   "NTK"
      Height          =   645
      Left            =   690
      TabIndex        =   0
      Top             =   540
      Width           =   2610
   End
End
Attribute VB_Name = "NTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub NTK_Click()
    Dim obj As Object
    Dim colPedidos As New Collection
    Dim lErro As Long
    
    Set obj = CreateObject("sgenfe.NTKIntegracao")
    lErro = obj.ObterPedidos("089d5ed56d2217c86f8136b99ea5892d", "71a9eb255ee997b446ffe57", "https://tokecompre-desenv.herokuapp.com", colPedidos)
    
End Sub
