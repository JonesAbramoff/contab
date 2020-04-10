VERSION 5.00
Begin VB.Form PopUpMenuPagtoAporteSF 
   Caption         =   "PopUpMenuPagtoAporteSF"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGrid 
      Caption         =   "Pagto Sobre Fatura"
      Begin VB.Menu mnuGridHistorico 
         Caption         =   "Histórico de utilização resumido"
         Index           =   1
      End
      Begin VB.Menu mnuGridHistorico 
         Caption         =   "Histórico de utilização detalhado"
         Index           =   2
      End
      Begin VB.Menu mnuGridSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "PopUpMenuPagtoAporteSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Const TRV_HIST_APORTE_RESUMIDO = 1
Const TRV_HIST_APORTE_DETALHADO = 2

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

Private Sub mnuGridHistorico_Click(Index As Integer)
    If Not (objTela Is Nothing) Then
        If Index = TRV_HIST_APORTE_RESUMIDO Then
            Call objTela.mnuGridHistorico_Click
        Else
            Call objTela.mnuGridHistorico_Click(True)
        End If
        Set objTela = Nothing
    End If
End Sub
