VERSION 5.00
Begin VB.Form PrincMDIChild 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   HelpContextID   =   1000
   Icon            =   "PrincmdichildECF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   5340
End
Attribute VB_Name = "PrincMDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

#Const HABILITA_MODO_EDICAO = 0

'###################################
'Inserido por Wagner
'Compilar com 1 para Hicare e 0 Para os demais Clientes
Const TUDO_MAIUSCULO = 0
Const PULA_SEGMENTO = 1
'###################################

Public sNomeTelaOcx As String

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
    
'Dim obj As VBControlExtender
'
'    For Each obj In Controls
'
'        If obj.Name = "Tela" Then Set objFormOcx = obj.object
'        Exit Property
'
'    Next
'
'    Set objFormOcx = Nothing
    
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
Dim colBrowseParamSelecao As New Collection
Dim sTela1 As String
Dim iPos As Integer
Dim objTela As Object

On Error GoTo Erro_Iniciar

    Call Adiciona_Licenca(sTela)
    
    If InStr(sTela, "Lista") <> 0 Then
    
        'mario. Codigo colocado enquanto todos os browsers não tiverem sido convertidos para o novo formato
        iPos = InStr(sTela, ".")
        sTela1 = Mid(sTela, iPos + 1)
    
'        'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
'        lErro = CF("BrowseParamSelecao_Le", sTela1, colBrowseParamSelecao)
'        If lErro <> SUCESSO Then gError 89983
'
'        If colBrowseParamSelecao.Count = 0 Then GoTo Label_Tela_Tradicional
    
        Set extCtl = Controls.Add("SGEECF.Browser", "Tela")
    Else
Label_Tela_Tradicional:
        Set extCtl = Controls.Add(sTela, "Tela")
    End If
            
    sNomeTelaOcx = sTela
    Set obj = extCtl
        
    With Me!Tela
        Me.left = 0
        Me.top = 0
        Me.Height = .Height + (1440 * (2 * GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYCAPTION)) / GetDeviceCaps(Me.hdc, LOGPIXELSY))
        Me.Width = .Width + (1440 * 2 * GetSystemMetrics(SM_CXFIXEDFRAME) / GetDeviceCaps(Me.hdc, LOGPIXELSX))
        .Visible = True
    End With

#If HABILITA_MODO_EDICAO = 1 Then
    'codigo para Edicao Tela
    If giLocalOperacao <> LOCALOPERACAO_ECF Then Call Inicializa_Edicao(Me!Tela.object)
#End If

    Set obj2 = Me!Tela.object.Form_Load_Ocx
    
    Call CF("Telas_Trata_NomeExibicao", obj2)

'Inserido por Wagner
'######################
#If HABILITA_MODO_EDICAO = 1 Then
    'codigo para Edicao Tela
    If giLocalOperacao <> LOCALOPERACAO_ECF Then Call Inicializa_Edicao_zOrder(Me!Tela.object)
#End If
'######################

    '??? tratamento especial para daltonicos
    If InStr(UCase(gsNomeEmpresa), "CHAVE DIGITAL") <> 0 Then

        'Seta o Obj do tipo da tela
        Set objTela = Me!Tela.object
        
        lErro = MudaCor_LabelObrigatorio(objTela)
        If lErro <> SUCESSO Then Error 11307
    
    End If

'    extCtl.CausesValidation = True
    
    Iniciar = lErro_Chama_Tela
    
    Exit Function
    
Erro_Iniciar:

    Iniciar = Err
     
    Select Case Err
          
        Case 11307
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165266)
     
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

Private Sub Form_Terminate()
'    MsgBox ("form_terminate")

''    If sNomeTelaOcx <> "" Then
''
''        'Controls.Remove "Tela"
''        Call Licenses.Remove(sNomeTelaOcx)
''
''    End If
    
End Sub

'*******************************************************
'eventos repassados p/ocx filho
'*******************************************************
Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload

    If Not (extCtl Is Nothing) Then
        
#If HABILITA_MODO_EDICAO = 1 Then

        If giLocalOperacao <> LOCALOPERACAO_ECF Then
            'mario. Inicio Edicao de Telas
            If Not (gobjTelaAtiva Is Nothing) Then
                If extCtl.object.Name = gobjTelaAtiva.Name And (Not gobjPropriedades Is Nothing) Then
                    Call gobjPropriedades.Limpar
                Else
                    Set gobjTelaAtiva = Nothing
                End If
            End If
    
            Call Finaliza_Edicao(extCtl.object)
            'mario. Fim Edicao de Telas
            
        End If
#End If

        Call extCtl.object.Form_Unload(Cancel)
        
'        Controls.Remove "Tela"
    
'        If sNomeTelaOcx <> "" Then
'
'            Call Licenses.Remove(sNomeTelaOcx)
'
'        End If
'        MsgBox ("form_unload")
        Set extCtl = Nothing
    
    End If

    Exit Sub

Erro_Form_Unload:

    Select Case Err

        Case 438 'o tratador do evento nao foi implementado na tela

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165267)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165268)
     
    End Select

    Exit Sub
    
End Sub

Private Sub Form_Activate()

On Error GoTo Erro_Form_Activate
    
    If Not (objFormOcx Is Nothing) Then
        
#If HABILITA_MODO_EDICAO = 1 Then
        'Codigo que auxilia a Edicao Tela
        If giLocalOperacao <> LOCALOPERACAO_ECF Then Set gobjTelaAtiva = objFormOcx
#End If

        Call objFormOcx.Form_Activate
        
    End If
    
    Exit Sub
    
Erro_Form_Activate:

    Select Case Err
          
        Case 438 'o tratador do evento nao foi implementado na tela
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165269)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165270)
     
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

Function MudaCor_LabelObrigatorio(objTela As Object) As Long
'Função que muda a Cor dos Label's da Tela para abobora(laranja claro)

Dim objMudaCor As Object
Dim lErro As Long

On Error GoTo Erro_MudaCor_LabelObrigatorio

    For Each objMudaCor In objTela.Controls
    
        'Verifica se o controle é uma label se for
        If TypeOf objMudaCor Is Label Then
           'Se o controle for uma label e de cor vermelha então altera para laranja claro
            If objMudaCor.ForeColor = &H80& Then
           
                '??? ecolub objMudaCor.ForeColor = &H80FF&
                objMudaCor.ForeColor = &H8000000D
            
            End If
            
        End If
    
    Next

    MudaCor_LabelObrigatorio = SUCESSO

    Exit Function

Erro_MudaCor_LabelObrigatorio:
    
    MudaCor_LabelObrigatorio = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165272)
     
    End Select

    Exit Function

End Function

