VERSION 5.00
Begin VB.Form PopupMenusCTB 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuGrid 
      Caption         =   "Rateio"
      Begin VB.Menu menuRateio 
         Caption         =   "&Aplicar Rateio"
      End
      Begin VB.Menu menulimpar 
         Caption         =   "&Limpar Grid"
      End
      Begin VB.Menu divisao 
         Caption         =   "-"
      End
      Begin VB.Menu menucancel 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "PopupMenusCTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

Private Sub menulimpar_Click()
    If Not (objTela Is Nothing) Then
        Call objTela.menulimpar_Click
        Set objTela = Nothing
    End If
End Sub

Private Sub menuRateio_Click()
    If Not (objTela Is Nothing) Then
        Call objTela.menuRateio_Click
        Set objTela = Nothing
    End If
End Sub
