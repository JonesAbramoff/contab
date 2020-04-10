VERSION 5.00
Begin VB.Form PrincPopUp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "PrincPopUpECF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu FluxoDeCaixa 
      Caption         =   "Fluxo De Caixa"
      Visible         =   0   'False
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
Attribute VB_Name = "PrincPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public sNomeTelaOcx As String
Private giWindowState As Integer

'###################################
'Inserido por Wagner
'Compilar com 1 para Hicare e 0 Para os demais Clientes
Const TUDO_MAIUSCULO = 0
Const PULA_SEGMENTO = 1
'###################################

Private WithEvents extCtl As VBControlExtender
Attribute extCtl.VB_VarHelpID = -1

Private Sub extCtl_ObjectEvent(Info As EventInfo)

   ' Program the events of the control using Select Case.
   Select Case Info.Name
   
        Case "Unload"      ' Handle unload event here.
            Unload Me
            
        ' Other cases now shown
        
        Case Else ' Unknown Event
            ' Handle unknown events here.
            
    End Select
    
End Sub

Public Property Get objFormOcx() As Object

    If Not (extCtl Is Nothing) Then
        Set objFormOcx = extCtl.object
    Else
        Set objFormOcx = Nothing
    End If
    
End Property

Private Sub Adiciona_Licenca(sTela As String)
On Error GoTo Erro_Adiciona_Licenca
    Call Licenses.Add(sTela)
    Exit Sub
    
Erro_Adiciona_Licenca:
    '??? tratar erro especifico
    Exit Sub
End Sub

Public Function Iniciar(sTela As String) As Long

Dim obj As Object, lAux As Long, lErro As Long, obj2 As Object
Dim colBrowseCampo As New Collection
Dim sTela1 As String
     
On Error GoTo Erro_Iniciar

    If InStr(sTela, "Lista") <> 0 Then
    
        'mario. Codigo colocado enquanto todos os browsers não tiverem sido convertidos para o novo formato
        sTela1 = Mid(sTela, InStr(sTela, ".") + 1)
    
        If UCase(right(sTela1, 3)) = "OCX" Then sTela1 = left(sTela1, Len(sTela1) - 3)
    
        'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
        lErro = CF("BrowseCampo_Le", sTela1, colBrowseCampo)
        If lErro <> SUCESSO Then gError 89983
    
        If colBrowseCampo.Count = 0 Then GoTo Label_Tela_Tradicional
    
        Set extCtl = Controls.Add("SGEECF.Browser", "Tela")
    Else
Label_Tela_Tradicional:
        Call Adiciona_Licenca(sTela)
    
        Set extCtl = Controls.Add(sTela, "Tela")
    End If

'
'
'    Call Adiciona_Licenca(sTela)
'
'    Set extCtl = Controls.Add(sTela, "Tela")
    sNomeTelaOcx = sTela
    Set obj = extCtl
        
    With Me!Tela
        Me.left = 0
        Me.top = 0
        Me.Height = .Height + (1440 * (2 * GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYCAPTION)) / GetDeviceCaps(Me.hdc, LOGPIXELSY))
        Me.Width = .Width + (1440 * 2 * GetSystemMetrics(SM_CXFIXEDFRAME) / GetDeviceCaps(Me.hdc, LOGPIXELSX))
        .Visible = True
    End With

    Set obj2 = Me!Tela.object.Form_Load_Ocx

    Iniciar = lErro_Chama_Tela
    
    Exit Function
    
Erro_Iniciar:

    Iniciar = Err
     
    Select Case Err
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165279)
     
    End Select
     
    Exit Function
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)

Dim State As Integer
Dim Ctrl As Control
Dim FirstTab As Integer, LastTab As Integer
Dim FirstCtrl As Control, LastCtrl As Control

On Error GoTo Erro_Form_KeyPress

    '######################################
    'Inserido por Wagner
    If TUDO_MAIUSCULO = 1 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If PULA_SEGMENTO = 1 Then Call Pula_Prox_Segmento(objFormOcx.ActiveControl, KeyAscii, objFormOcx.Caption)
    '######################################

    'teclou TAB
    If KeyAscii = 9 Then

         If Not (objFormOcx Is Nothing) Then

            State = GetKeyState(vbKeyShift) And &H1110 'Eliminate 1st bit

            ' Loop though all the controls and find
            ' the last control in the tab order
            For Each Ctrl In objFormOcx.Controls

                If Ctrl.TabStop = True And Ctrl.Visible = True And Ctrl.Enabled = True And Ctrl.TabIndex >= 0 Then

                    If FirstCtrl Is Nothing Then
                        FirstTab = Ctrl.TabIndex
                        LastTab = Ctrl.TabIndex
                        Set FirstCtrl = Ctrl
                        Set LastCtrl = Ctrl
                    End If

                    If Ctrl.TabIndex < FirstTab Then
                        FirstTab = Ctrl.TabIndex
                        Set FirstCtrl = Ctrl
                    ElseIf Ctrl.TabIndex >= LastTab Then 'Maximum value.
                        LastTab = Ctrl.TabIndex
                        Set LastCtrl = Ctrl
                    End If

                End If

Proximo_Controle:

            Next Ctrl

            If Not (FirstCtrl Is Nothing) Then

                If State = 0 Then
                    FirstCtrl.SetFocus
                Else
                    LastCtrl.SetFocus
                End If

            End If

        End If

    End If

    Exit Sub

Erro_Form_KeyPress:

    Select Case Err

        Case 438
            Resume Proximo_Controle

    End Select

End Sub

Private Sub Form_Load()

    '??? nao fazer nada
    
End Sub

Private Sub Form_Resize()

Dim objUserControl As UserControl

    If giLocalOperacao <> LOCALOPERACAO_ECF And InStr(UCase(sNomeTelaOcx), "LISTA") <> 0 Then

        If giWindowState = Me.WindowState And (Me.WindowState = 1 Or Me.WindowState = 2) Then Exit Sub
        giWindowState = Me.WindowState
    
        'Alteracao Daniel em 13/11/2001
        'If Me.WindowState <> 1 And Me.WindowState <> 2 Then
        If Me.WindowState <> 1 Then
            If Me.Height < 3000 Then Me.Height = 4000
            If Me.Width < 2000 Then Me.Width = 2000
        
            If Me.Visible = True Then
                objFormOcx.Tamanho Me.Width - (1440 * 4 * GetSystemMetrics(SM_CXFIXEDFRAME) / GetDeviceCaps(Me.hdc, LOGPIXELSX)), Me.Height - (1440 * (6 * GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYCAPTION)) / GetDeviceCaps(Me.hdc, LOGPIXELSY))
            End If
        End If

    End If

    If giLocalOperacao = LOCALOPERACAO_ECF And (InStr(UCase(sNomeTelaOcx), "VENDAM") Or InStr(UCase(sNomeTelaOcx), "VENDAP")) Then
        objFormOcx.Form_Resize Me.Width - (1440 * 4 * GetSystemMetrics(SM_CXFIXEDFRAME) / GetDeviceCaps(Me.hdc, LOGPIXELSX)), Me.Height - (1440 * (6 * GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYCAPTION)) / GetDeviceCaps(Me.hdc, LOGPIXELSY))
    End If

End Sub

Private Sub Form_Terminate()
    
'    If sNomeTelaOcx <> "" Then
'
'        'Controls.Remove "Tela"
'        Call Licenses.Remove(sNomeTelaOcx)
'
'    End If
    
End Sub

'*******************************************************
'eventos repassados p/ocx filho
'*******************************************************
Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload

    If Not (extCtl Is Nothing) Then
        
        Call extCtl.object.Form_Unload(Cancel)
        Set extCtl = Nothing
    
    End If

    Exit Sub

Erro_Form_Unload:

    Select Case Err

        Case 438 'o tratador do evento nao foi implementado na tela

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165280)

    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo Erro_Form_QueryUnload
    
Dim objControl As Object

    If Not (objFormOcx Is Nothing) Then
    
        Me.ValidateControls
        
        Call objFormOcx.Form_QueryUnload(Cancel, UnloadMode, Forms(0).ActiveForm Is Me And Screen.ActiveForm Is Me)
            
    End If
    
    Exit Sub
    
Erro_Form_QueryUnload:

    Select Case Err
          
        Case 438 'o tratador do evento nao foi implementado na tela
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165281)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Form_Activate()

On Error GoTo Erro_Form_Activate
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Form_Activate
    
    Exit Sub
    
Erro_Form_Activate:

    Select Case Err
          
        Case 438 'o tratador do evento nao foi implementado na tela
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165282)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Form_Deactivate()

On Error GoTo Erro_Form_Deactivate
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Form_Deactivate

    Exit Sub
    
Erro_Form_Deactivate:

    Select Case Err
          
        Case 438 'o tratador do evento nao foi implementado na tela
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165283)
     
    End Select

    Exit Sub
    
End Sub

'*******************************************************
'eventos recebidos do filho
'*******************************************************
Public Sub UnloadDoFilho()

    Unload Me
    
End Sub

'**** eca
Private Sub Pag_Fornecedor_Click()

On Error GoTo Erro_Pag_Fornecedor_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Pag_Fornecedor_Click
    
    Exit Sub
    
Erro_Pag_Fornecedor_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165284)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Pag_TipoFornecedor_Click()

On Error GoTo Erro_Pag_TipoFornecedor_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Pag_TipoFornecedor_Click
    
    Exit Sub
    
Erro_Pag_TipoFornecedor_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165285)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Pag_Titulo_Click()

On Error GoTo Erro_Pag_Titulo_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Pag_Titulo_Click
    
    Exit Sub
    
Erro_Pag_Titulo_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165286)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Rec_Cliente_Click()

On Error GoTo Erro_Rec_Cliente_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Rec_Cliente_Click
    
    Exit Sub
    
Erro_Rec_Cliente_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165287)
     
    End Select

    Exit Sub

End Sub

Private Sub Rec_TipoCliente_Click()

On Error GoTo Erro_Rec_TipoCliente_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Rec_TipoCliente_Click
    
    Exit Sub
    
Erro_Rec_TipoCliente_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165288)
     
    End Select

    Exit Sub

End Sub

Private Sub Rec_Titulo_Click()

On Error GoTo Erro_Rec_Titulo_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Rec_Titulo_Click
    
    Exit Sub
    
Erro_Rec_Titulo_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165289)
     
    End Select

    Exit Sub

End Sub

Private Sub Saldos_Iniciais_Click()

On Error GoTo Erro_Saldos_Iniciais_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Saldos_Iniciais_Click
    
    Exit Sub
    
Erro_Saldos_Iniciais_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165290)
     
    End Select

    Exit Sub

End Sub

Private Sub Sint_Projecao_Click()

On Error GoTo Erro_Sint_Projecao_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Sint_Projecao_Click
    
    Exit Sub
    
Erro_Sint_Projecao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165291)
     
    End Select

    Exit Sub

End Sub

Private Sub Sint_Revisao_Click()

On Error GoTo Erro_Sint_Revisao_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Sint_Revisao_Click
    
    Exit Sub
    
Erro_Sint_Revisao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165292)
     
    End Select

    Exit Sub

End Sub

Private Sub Aplic_Aplicacao_Click()

On Error GoTo Erro_Aplic_Aplicacao_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Aplic_Aplicacao_Click
    
    Exit Sub
    
Erro_Aplic_Aplicacao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165293)
     
    End Select

    Exit Sub

End Sub

Private Sub Aplic_TipoAplicacao_Click()

On Error GoTo Erro_Aplic_TipoAplicacao_Click
    
    If Not (objFormOcx Is Nothing) Then Call objFormOcx.Aplic_TipoAplicacao_Click
    
    Exit Sub
    
Erro_Aplic_TipoAplicacao_Click:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165294)
     
    End Select

    Exit Sub

End Sub

'####################################################################
'Inserido por Wagner
Function Pula_Prox_Segmento(ByVal objCampo1 As Control, KeyAscii As Integer, ByVal sNomeTela As String) As Long
'ao teclar . faz com que o cursor se dirija para o inicio do proximo segmento
'Recebe: O Controle, a tecla que foi apertada e o nome da tela

Dim iPos As Integer
Dim colSeg As New Collection
Dim vSeg As Variant

On Error GoTo Erro_Pula_Prox_Segmento

    If KeyAscii = Asc(".") Then

        'Se existe ponto na máscara
        If InStr(1, objCampo1.Mask, ".") <> 0 Then
            iPos = 1
            Do While iPos > 0 And iPos < Len(objCampo1.Mask)
                iPos = InStr(iPos, objCampo1.Mask, ".")
                If iPos > 0 And iPos <= Len(objCampo1.Mask) Then
                    colSeg.Add iPos
                    iPos = iPos + 1
                End If
            
            Loop
        
            If colSeg.Count > 0 Then
                For Each vSeg In colSeg
                    If objCampo1.SelStart + 1 <= vSeg Then
                        objCampo1.SelStart = vSeg - 1
                        'não deixa o caracter digitado prosseguir no seu processamento
                        KeyAscii = 0
                        Exit Function
                    End If
                Next
            End If
        End If

    End If

    Pula_Prox_Segmento = SUCESSO
    
    Exit Function
    
Erro_Pula_Prox_Segmento:

    Pula_Prox_Segmento = gErr

    Select Case gErr
    
        Case 438
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165271)
            
    End Select
    
    Exit Function

End Function
'#####################################################################
