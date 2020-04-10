VERSION 5.00
Begin VB.Form PopUpMenuRelacClientes 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuRelacionamentosClientes 
      Caption         =   "Relacionamentos com Clientes"
      Begin VB.Menu mnuNovoOrcamento 
         Caption         =   "Novo Orçamento (F5)"
      End
      Begin VB.Menu mnuNovoPedido 
         Caption         =   "Novo Pedido (F6)"
      End
      Begin VB.Menu mnuNovoRelacionamento 
         Caption         =   "Novo Relacionamento (F7)"
      End
      Begin VB.Menu mnuConsultas 
         Caption         =   "Consultas (F8)"
      End
      Begin VB.Menu mnuEditarRelacionamento 
         Caption         =   "Editar Relacionamento (F9)"
      End
   End
End
Attribute VB_Name = "PopUpMenuRelacClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

Private Sub mnuNovoOrcamento_Click()

    If Not (objTela Is Nothing) Then
        Call objTela.mnuRelacClientes_NovoOrcamento_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuNovoPedido_Click()

    If Not (objTela Is Nothing) Then
        Call objTela.mnuRelacClientes_NovoPedido_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuNovoRelacionamento_Click()

    If Not (objTela Is Nothing) Then
        Call objTela.mnuRelacClientes_NovoRelacionamento_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuConsultas_Click()

    If Not (objTela Is Nothing) Then
        Call objTela.mnuRelacClientes_Consultas_Click
        Set objTela = Nothing
    End If

End Sub

Private Sub mnuEditarRelacionamento_Click()

    If Not (objTela Is Nothing) Then
        Call objTela.mnuRelacClientes_EditarRelacionamento_Click
        Set objTela = Nothing
    End If

End Sub

