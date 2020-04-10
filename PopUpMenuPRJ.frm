VERSION 5.00
Begin VB.Form PopUpMenuPRJ 
   Caption         =   "PopUpMenuGrid"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGrid 
      Caption         =   "Etapa"
      Begin VB.Menu mnuTvwAbrirEtapa 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnuTvwDocRelacs 
         Caption         =   "Documentos Relacionados"
      End
      Begin VB.Menu mnuGridSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTvwCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "PopUpMenuPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

Private Sub mnuTvwAbrirEtapa_Click()
    
    If Not (objTela Is Nothing) Then
        Call objTela.mnuTvwAbrirEtapa_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuTvwDocRelacs_Click()
    
    If Not (objTela Is Nothing) Then
        Call objTela.mnuTvwDocRelacs_Click
        Set objTela = Nothing
    End If

End Sub
