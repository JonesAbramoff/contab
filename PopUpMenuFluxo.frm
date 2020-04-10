VERSION 5.00
Begin VB.Form PopUpMenuFluxo 
   Caption         =   "PopUpMenuFluxo"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu FluxoDeCaixa 
      Caption         =   "Fluxo de Caixa"
      Begin VB.Menu Pagamentos 
         Caption         =   "Pagamentos"
         Begin VB.Menu Pag_TipoFornecedor 
            Caption         =   "Por Tipo de Fornecedor"
         End
         Begin VB.Menu Pag_Fornecedor 
            Caption         =   "Por Fornecedor"
         End
         Begin VB.Menu Pag_Titulo 
            Caption         =   "Por Título"
         End
      End
      Begin VB.Menu Recebimentos 
         Caption         =   "Recebimentos"
         Begin VB.Menu Rec_TipoCliente 
            Caption         =   "Por Tipo de Cliente"
         End
         Begin VB.Menu Rec_Cliente 
            Caption         =   "Por Cliente"
         End
         Begin VB.Menu Rec_Titulo 
            Caption         =   "Por Título"
         End
      End
      Begin VB.Menu Resgates 
         Caption         =   "Resgates"
         Begin VB.Menu Aplic_TipoAplicacao 
            Caption         =   "Por Tipo de Aplicação"
         End
         Begin VB.Menu Aplic_Aplicacao 
            Caption         =   "Por Aplicação"
         End
      End
      Begin VB.Menu Saldos_Iniciais 
         Caption         =   "Saldos Iniciais"
      End
      Begin VB.Menu Sintetico 
         Caption         =   "Sintético"
         Begin VB.Menu Sint_Projecao 
            Caption         =   "Por Projeção"
         End
         Begin VB.Menu Sint_Revisao 
            Caption         =   "Por Revisão"
         End
      End
   End
End
Attribute VB_Name = "PopUpMenuFluxo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objTela As Object

Private Sub Form_Unload(Cancel As Integer)
    Set objTela = Nothing
End Sub

'*******************************************************
'eventos recebidos do filho
'*******************************************************
Private Sub Pag_Fornecedor_Click()

On Error GoTo Erro_Pag_Fornecedor_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Pag_Fornecedor_Click
    End If
    
    Exit Sub
    
Erro_Pag_Fornecedor_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165044)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Pag_TipoFornecedor_Click()

On Error GoTo Erro_Pag_TipoFornecedor_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Pag_TipoFornecedor_Click
    End If
    
    Exit Sub
    
Erro_Pag_TipoFornecedor_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165045)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Pag_Titulo_Click()

On Error GoTo Erro_Pag_Titulo_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Pag_Titulo_Click
    End If
    
    Exit Sub
    
Erro_Pag_Titulo_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165046)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Rec_Cliente_Click()

On Error GoTo Erro_Rec_Cliente_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Rec_Cliente_Click
    End If
    
    Exit Sub
    
Erro_Rec_Cliente_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165047)
     
    End Select

    Exit Sub

End Sub

Private Sub Rec_TipoCliente_Click()

On Error GoTo Erro_Rec_TipoCliente_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Rec_TipoCliente_Click
    End If
    
    Exit Sub
    
Erro_Rec_TipoCliente_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165048)
     
    End Select

    Exit Sub

End Sub

Private Sub Rec_Titulo_Click()

On Error GoTo Erro_Rec_Titulo_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Rec_Titulo_Click
    End If
    
    Exit Sub
    
Erro_Rec_Titulo_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165049)
     
    End Select

    Exit Sub

End Sub

Private Sub Saldos_Iniciais_Click()

On Error GoTo Erro_Saldos_Iniciais_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Saldos_Iniciais_Click
    End If
    
    Exit Sub
    
Erro_Saldos_Iniciais_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165050)
     
    End Select

    Exit Sub

End Sub

Private Sub Sint_Projecao_Click()

On Error GoTo Erro_Sint_Projecao_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Sint_Projecao_Click
    End If
    
    Exit Sub
    
Erro_Sint_Projecao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165051)
     
    End Select

    Exit Sub

End Sub

Private Sub Sint_Revisao_Click()

On Error GoTo Erro_Sint_Revisao_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Sint_Revisao_Click
    End If
    
    Exit Sub
    
Erro_Sint_Revisao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165052)
     
    End Select

    Exit Sub

End Sub

Private Sub Aplic_Aplicacao_Click()

On Error GoTo Erro_Aplic_Aplicacao_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Aplic_Aplicacao_Click
    End If
    
    Exit Sub
    
Erro_Aplic_Aplicacao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165053)
     
    End Select

    Exit Sub

End Sub

Private Sub Aplic_TipoAplicacao_Click()

On Error GoTo Erro_Aplic_TipoAplicacao_Click
    
    If Not (objTela Is Nothing) Then
        Call objTela.Aplic_TipoAplicacao_Click
    End If
    
    Exit Sub
    
Erro_Aplic_TipoAplicacao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165054)
     
    End Select

    Exit Sub

End Sub



