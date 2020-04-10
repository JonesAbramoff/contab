VERSION 5.00
Begin VB.Form PopUpMenuGridMD 
   Caption         =   "PopUpMenuGrid"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGrid 
      Caption         =   "Grid"
      Begin VB.Menu mnuGridConsultaDocOriginal 
         Caption         =   "Documento Original"
      End
      Begin VB.Menu mnuGridSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridMarcarTodos 
         Caption         =   "Marcar Todos"
      End
      Begin VB.Menu mnuGridDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
      End
      Begin VB.Menu mnuGridSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "PopUpMenuGridMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

Private Sub mnuGridConsultaDocOriginal_Click()
    
    If Not (objTela Is Nothing) Then
        Call objTela.mnuGridConsultaDocOriginal_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuGridMarcarTodos_Click()
    
    If Not (objTela Is Nothing) Then
        Call objTela.mnuGridMarcarTodos_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuGridDesmarcarTodos_Click()
    
    If Not (objTela Is Nothing) Then
        Call objTela.mnuGridDesmarcarTodos_Click
        Set objTela = Nothing
    End If

End Sub
