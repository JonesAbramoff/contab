Attribute VB_Name = "SistGlobaisAdm"
Option Explicit

'Variáveis que armazenam independente de instância a property da classe
Public AdmGlob_iRetornoTela As Integer
Public AdmGlob_colFiliais As Collection
Public AdmGlob_colFiliaisEmpresa As Collection
Public AdmGlob_colUFs As Collection
Public Const WM_NCMOUSEMOVE = &HA0


'********** Inicio de Edicao Tela ***************

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    WindowProc = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc)
    
End Function

Function WindowProc0(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    WindowProc0 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc0)
    
End Function

Function WindowProc00(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    WindowProc00 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc00)
    
End Function

Function WindowProc1(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    WindowProc1 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc1)
    
End Function

Function WindowProc2(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc2 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc2)

End Function
        
Function WindowProc3(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc3 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc3)

End Function
        
Function WindowProc4(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc4 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc4)

End Function
        
Function WindowProc5(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc5 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc5)

End Function
        
Function WindowProc6(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc6 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc6)

End Function
        
Function WindowProc7(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc7 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc7)

End Function
        
Function WindowProc8(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc8 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc8)

End Function
        
Function WindowProc9(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc9 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc9)

End Function
        
Function WindowProc10(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc10 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc10)

End Function
        
Function WindowProc11(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc11 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc11)

End Function
        
Function WindowProc12(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc12 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc12)

End Function
        
Function WindowProc13(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc13 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc13)

End Function
        
Function WindowProc14(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc14 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc14)

End Function
        
Function WindowProc15(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc15 = My_WindowProc(hw, uMsg, wParam, lParam, glpPrevWndProc15)

End Function

Function My_WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal My_lpPrevWndProc As Long) As Long

Dim sParam As String
Dim lUnits As Long
Dim iNaoProcessa As Integer
Dim objControle As Control
Dim typePoint As POINTAPI

'On Error GoTo Erro_My_WindowProc

    Select Case uMsg
    
        Case WM_LBUTTONDOWN
                        
            If gobjmenuEdicao.Checked = True And gobjControleDrag Is Nothing Then
                If Not (gobjTelaAtiva Is Nothing) Then
                    For Each objControle In gobjTelaAtiva.Controls
                        If Not (TypeName(objControle) = "Label") And Not (TypeName(objControle) = "Menu") And Not (TypeName(objControle) = "Line") And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") And Not (TypeName(objControle) = "TabEndereco") And Not (TypeName(objControle) = "TabEnderecoAF") And Not (TypeName(objControle) = "TabEnderecoBol") And Not (TypeName(objControle) = "TabTributacaoFat") Then
                            If objControle.hWnd = hw Then
                                Set gobjControleDrag = objControle
                                Exit For
                            End If
                        End If
                    Next
                
                End If
                
                If Not (gobjControleDrag Is Nothing) Then
                
                    gsngEdicaoX = CSng("&H" & right(Hex(lParam), 4))
                    sParam = String(8 - Len(Hex(lParam)), "0") & Hex(lParam)
                    gsngEdicaoY = CSng("&H" & left(sParam, 4))
                    
                    gobjEstInicial.Timer1.Interval = 1
'                    If TypeName(objControle)= "CommandButton" Or TypeName(objControle)= "ComboBox" Or TypeName(objControle)= "TreeView" Or TypeName(objControle)= "UpDown" Or TypeName(objControle)= "CheckBox" Or TypeName(objControle)= "SSCheck" Then iNaoProcessa = 1
                    If TypeName(objControle) = "CommandButton" Or TypeName(objControle) = "ComboBox" Or TypeName(objControle) = "TreeView" Or TypeName(objControle) = "UpDown" Or TypeName(objControle) = "CheckBox" Then iNaoProcessa = 1
                End If
            
            End If
            
        Case WM_LBUTTONUP
        
            If gobjmenuEdicao.Checked = True Then
                    
                If giProxButtonUp = 1 And Not (gobjControleDrag Is Nothing) Then
                        giProxMouseMove = 1
                        giProxButtonUp = 0
                End If
            
            End If
            
        Case WM_MOUSEMOVE, WM_NCMOUSEMOVE

            If gobjmenuEdicao.Checked = True Then

                If giProxMouseMove = 1 Then

                    giProxMouseMove = 0
                    
                    If Not (gobjControleDrag Is Nothing) Then
                        If Not (gobjTelaAtiva Is Nothing) Then
                            If hw = gobjTelaAtiva.hWnd Then
                                Set gobjControleAlvo = gobjTelaAtiva
                            Else
        
                                For Each objControle In gobjTelaAtiva.Controls
                                    If Not (TypeName(objControle) = "Label") And Not (TypeName(objControle) = "Menu") And Not (TypeName(objControle) = "Line") And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") And Not (TypeName(objControle) = "TabEndereco") And Not (TypeName(objControle) = "TabEnderecoAF") And Not (TypeName(objControle) = "TabEnderecoBol") And Not (TypeName(objControle) = "TabTributacaoFat") Then
                                        If objControle.hWnd = hw Then
                                            Set gobjControleAlvo = objControle
                                            Exit For
                                        End If
                                    End If
                                Next
        
                            End If
    
                        End If
    
                        gsngEdicaoX = CSng("&H" & right(Hex(lParam), 4))
                        sParam = String(8 - Len(Hex(lParam)), "0") & Hex(lParam)
                        gsngEdicaoY = CSng("&H" & left(sParam, 4))
    
                        If uMsg = WM_NCMOUSEMOVE Then
    
                            typePoint.X = gsngEdicaoX
                            typePoint.Y = gsngEdicaoY
                            
                            Call ScreenToClient(hw, typePoint)
                            
                            gsngEdicaoX = typePoint.X
                            gsngEdicaoY = typePoint.Y
                            
                        End If
                            
                        gobjEstInicial.Timer2.Interval = 1

                    End If
                    
                End If

            End If
            
    End Select
    
    If iNaoProcessa <> 1 Then My_WindowProc = CallWindowProc(My_lpPrevWndProc, hw, uMsg, wParam, lParam)
    
    Exit Function
    
'Erro_My_WindowProc:
'
'    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 184218)
'
'    Exit Function

End Function


'********** Fim Edicao Tela ***************

