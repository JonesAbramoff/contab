VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmAdm"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvargsNomeTela As String
Private mvargcolBrowseUsuarioCampo As Collection
Private mvargiTelaOK As Integer

'Usada pelas ClassTipo nos vários Projetos
Public Enum enumTipo
    TIPO_VAZIO
    TIPO_STRING
    TIPO_DECIMAL
    TIPO_NAO_POSITIVO
    TIPO_CODIGO
    TIPO_CPF
    TIPO_CGC
    TIPO_OVERFLOW
End Enum

'Usada pela função Codigos_Le
Public Enum enumTipoCodigo
    TIPO_INT
    TIPO_LONG
    TIPO_STR
End Enum

'Serve para comunicação entre tela modal chamada e chamadora
Public Property Let giRetornoTela(ByVal vData As Integer)
    AdmGlob_iRetornoTela = vData
End Property

Public Property Get giRetornoTela() As Integer
    giRetornoTela = AdmGlob_iRetornoTela
End Property

Public Property Set gcolFiliais(ByVal vData As Object)
    Set AdmGlob_colFiliais = vData
End Property

Public Property Get gcolFiliais() As Collection
    Set gcolFiliais = AdmGlob_colFiliais
End Property

Public Property Set gcolFiliaisEmpresa(ByVal vData As Object)
    Set AdmGlob_colFiliaisEmpresa = vData
End Property

Public Property Get gcolFiliaisEmpresa() As Collection
    Set gcolFiliaisEmpresa = AdmGlob_colFiliaisEmpresa
End Property

Public Property Get gcolUFs() As Collection
        
Dim lErro As Long

On Error GoTo Erro_gcolUFs

    If AdmGlob_colUFs Is Nothing Then
    
        Set AdmGlob_colUFs = New Collection
    
        lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, AdmGlob_colUFs, STRING_ESTADOS_SIGLA)
        If lErro <> SUCESSO Then Error 46532
        
    End If
    
    Set gcolUFs = AdmGlob_colUFs
    
    Exit Property
    
Erro_gcolUFs:

    Set gcolUFs = Nothing
     
    Select Case Err
          
        Case 46532
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142038)
     
    End Select
     
    Exit Property

End Property

Public Property Let giTelaOK(ByVal vData As Integer)
    mvargiTelaOK = vData
End Property

Public Property Get giTelaOK() As Integer
    giTelaOK = mvargiTelaOK
End Property

Public Property Set gcolBrowseUsuarioCampo(ByVal vData As Object)
    Set mvargcolBrowseUsuarioCampo = vData
End Property

Public Property Get gcolBrowseUsuarioCampo() As Collection
    Set gcolBrowseUsuarioCampo = mvargcolBrowseUsuarioCampo
End Property

Public Property Let gsNomeTela(ByVal vData As String)
    mvargsNomeTela = vData
End Property

Public Property Get gsNomeTela() As String
    gsNomeTela = mvargsNomeTela
End Property

'******** Inicio Edicao de Telas **************************

Public Sub Inicializa_Edicao(objTela As Object)

Dim objControle As Object
    
    If Controle_ObterNomeClasse(objTela) = "Form" Then
        glpPrevWndProc0 = SetWindowLong(objTela.hWnd, GWL_WNDPROC, AddressOf WindowProc0)
    Else
        glpPrevWndProc00 = SetWindowLong(objTela.hWnd, GWL_WNDPROC, AddressOf WindowProc00)
    End If
    
'mario. colocar aqui os demais tipos de controle a serem tratados
'para cada controle criar lpPrevWndProc<numero> e WndProc<numero>
    
    For Each objControle In objTela.Controls
    
        If Controle_ObterNomeClasse(objControle) = "MaskEdBox" Then
            glpPrevWndProc = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc)
        ElseIf Controle_ObterNomeClasse(objControle) = "TextBox" Then
            glpPrevWndProc1 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc1)
        ElseIf Controle_ObterNomeClasse(objControle) = "CommandButton" Then
            glpPrevWndProc2 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc2)
        ElseIf Controle_ObterNomeClasse(objControle) = "Frame" Then
            glpPrevWndProc3 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc3)
''        ElseIf Controle_ObterNomeClasse(objControle) = "SSFrame" Then
''            glpPrevWndProc4 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc4)
        ElseIf Controle_ObterNomeClasse(objControle) = "TreeView" Then
            glpPrevWndProc5 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc5)
        ElseIf Controle_ObterNomeClasse(objControle) = "MSFlexGrid" Then
            glpPrevWndProc6 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc6)
        ElseIf Controle_ObterNomeClasse(objControle) = "CheckBox" Then
            glpPrevWndProc7 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc7)
        ElseIf Controle_ObterNomeClasse(objControle) = "PictureBox" Then
            glpPrevWndProc8 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc8)
        ElseIf Controle_ObterNomeClasse(objControle) = "UpDown" Then
            glpPrevWndProc9 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc9)
        ElseIf Controle_ObterNomeClasse(objControle) = "ComboBox" Then
            glpPrevWndProc10 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc10)
        ElseIf Controle_ObterNomeClasse(objControle) = "ListBox" Then
            glpPrevWndProc11 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc11)
'        ElseIf TypeOf objControle Is SSCheck Then
'            glpPrevWndProc12 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc12)
        ElseIf Controle_ObterNomeClasse(objControle) = "OptionButton" Then
            glpPrevWndProc13 = SetWindowLong(objControle.hWnd, GWL_WNDPROC, AddressOf WindowProc13)
        End If

    Next

    Call Preenche_Propriedades_EdicaoTela(objTela)

End Sub

Public Sub Finaliza_Edicao(objTela As Object)

Dim objControle As Object
Dim sNomeControle As String
    
    If Controle_ObterNomeClasse(objTela) = "Form" Then
        Call SetWindowLong(objTela.hWnd, GWL_WNDPROC, glpPrevWndProc0)
    Else
        Call SetWindowLong(objTela.hWnd, GWL_WNDPROC, glpPrevWndProc00)
    End If
    
'mario. colocar aqui os demais tipos de controle a serem tratados
'para cada controle criar lpPrevWndProc<numero> e WndProc<numero>
    
    For Each objControle In objTela.Controls
    
        'Apenas por Segurança ...
        sNomeControle = ""
        
        'Obtem o tipo do Controle
        sNomeControle = Controle_ObterNomeClasse(objControle)
    
        Select Case sNomeControle
        
            Case "MaskEdBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc)
                
            Case "TextBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc1)
                
            Case "CommandButton"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc2)
                
            Case "Frame"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc3)
                
    ''        ElseIf TypeOf objControle Is SSFrame Then
    ''            Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc4)
    
            Case "TreeView"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc5)
                
            Case "MSFlexGrid"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc6)
                
            Case "CheckBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc7)
                
            Case "PictureBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc8)
                
            Case "UpDown"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc9)
                
            Case "ComboBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc10)
                
            Case "ListBox"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc11)
                
    '        ElseIf TypeOf objControle Is SSCheck Then
    '            Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc12)
    
            Case "OptionButton"
                Call SetWindowLong(objControle.hWnd, GWL_WNDPROC, glpPrevWndProc13)
                
            End Select
            
    Next

End Sub

Private Function Preenche_Propriedades_EdicaoTela(objTela As Object) As Long
'Altera as propriedades dos Controles da Tela

Dim objEdicaoTela_Tela As New ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As New ClassEdicaoTela_Controle
Dim objControle As Object
Dim objControle2 As Object
Dim iAchou As Integer
Dim sNome As String
Dim sTabIndexBuffer As String
Dim iTabIndex As Integer
Dim iPosicaoPrimeiroEspaco As Integer
Dim iPosicaoSegundoEspaco As Integer
Dim colControles As New Collection
Dim iIndice As Integer
Dim sControle As String
Dim sControle2 As String
Dim iIndice2 As Integer

On Error GoTo Erro_Preenche_Propriedades_EdicaoTela

    'Procura a Tela
    For Each objEdicaoTela_Tela In gcolEdicaoTela
    
        If objEdicaoTela_Tela.sNomeTela = objTela.Name Then
            
            'Para cada controle da coleção
            For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
                
                'Procura o controle respectivo na Tela
                For Each objControle In objTela.Controls
                    
                    iIndice2 = -1
                    iIndice2 = objControle.Index
                    
                    If iIndice2 > -1 Then
                        sControle = objControle.Name & "(" & objControle.Index & ")"
                    Else
                        sControle = objControle.Name
                    End If

                    If sControle = objEdicaoTela_Controle.sNomeControle Then
                        
                        'Inserido por Wagner
                        '###############
                        If objEdicaoTela_Controle.izOrder <> -1 Then objControle.ZOrder objEdicaoTela_Controle.izOrder
                        
                        '###############
                        
                        'Altera as suas propriedades
                        If (Controle_ObterNomeClasse(objControle) = "Label") Or (Controle_ObterNomeClasse(objControle) = "CommandButton") Or (Controle_ObterNomeClasse(objControle) = "OptionButton") Or (Controle_ObterNomeClasse(objControle) = "Frame") Or (Controle_ObterNomeClasse(objControle) = "CheckBox") Then
                            objEdicaoTela_Controle.sTituloPadrao = objControle.Caption
                            objControle.Caption = objEdicaoTela_Controle.sTitulo
                        End If
                        
                        If objEdicaoTela_Controle.iAltura = -1 Then
                            objEdicaoTela_Controle.iAlturaPadrao = -1
                        Else
                            objEdicaoTela_Controle.iAlturaPadrao = objControle.Height
                            objControle.Height = objEdicaoTela_Controle.iAltura
                        End If
                        
                        If objControle.Enabled Then
                            objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO
                        Else
                            objEdicaoTela_Controle.iHabilitadoPadrao = DESMARCADO
                        End If
                        
''                        If (Not (TypeOf objControle Is Label)) And (Not (TypeOf objControle Is PictureBox)) And (Not (TypeOf objControle Is Frame)) And (Not (TypeOf objControle Is SSFrame)) And (Not (TypeOf objControle Is SSPanel)) Then
                        If (Not (Controle_ObterNomeClasse(objControle) = "Label")) And (Not (Controle_ObterNomeClasse(objControle) = "PictureBox")) And (Not (Controle_ObterNomeClasse(objControle) = "Frame")) And (Not (Controle_ObterNomeClasse(objControle) = "CommonDialog")) And (Not (Controle_ObterNomeClasse(objControle) = "Image")) Then
                            objEdicaoTela_Controle.iTabIndexPadrao = objControle.TabIndex
                        End If

                        objEdicaoTela_Controle.iEsquerdaPadrao = objControle.left
                        objControle.left = objEdicaoTela_Controle.iEsquerda
                        objEdicaoTela_Controle.iLarguraPadrao = objControle.Width
                        objControle.Width = objEdicaoTela_Controle.iLargura
                        
                        If objEdicaoTela_Controle.iTabStop = -1 Then
                            objEdicaoTela_Controle.iTabStopPadrao = -1
                        Else
                            If objControle.TabStop = True Then
                                objEdicaoTela_Controle.iTabStopPadrao = 1
                            Else
                                objEdicaoTela_Controle.iTabStopPadrao = 0
                            End If
                            
                            objControle.TabStop = objEdicaoTela_Controle.iTabStop
                        End If
                        
                        objEdicaoTela_Controle.iTopoPadrao = objControle.top
                        objControle.top = objEdicaoTela_Controle.iTopo
                        objEdicaoTela_Controle.iVisivelPadrao = CInt(objControle.Visible)
                        
'                        If objEdicaoTela_Controle.iVisivel = 1 Then
'                            objControle.Visible = True
'                        Else
'                            objControle.Visible = False
'                        End If
                        
                        If objTela.Name = objControle.Container.Name Then
                            objEdicaoTela_Controle.iIndiceContainerPadrao = -1
                            objEdicaoTela_Controle.sContainerPadrao = objControle.Container.Name
                        Else
                            objEdicaoTela_Controle.iIndiceContainerPadrao = -1
                            objEdicaoTela_Controle.iIndiceContainerPadrao = objControle.Container.Index
                            objEdicaoTela_Controle.sContainerPadrao = objControle.Container.Name
                        End If

                        iAchou = 0
                        
                        For Each objControle2 In objTela.Controls
                            
                            iIndice2 = -1
                            iIndice2 = objControle2.Index
                            
                            sControle2 = objControle2.Name

                            If sControle2 = objEdicaoTela_Controle.sContainer Then
                                
                                If objEdicaoTela_Controle.iIndiceContainer > -1 Then
                                    
                                    If objControle2.Index = objEdicaoTela_Controle.iIndiceContainer Then
                                        
                                        Set objControle.Container = objControle2
                                        iAchou = 1
                                        Exit For
                                    End If
                                Else
                                    Set objControle.Container = objControle2
                                    iAchou = 1
                                End If
                            End If
                        Next
    
                        If iAchou <> 1 Then
    
                            If objEdicaoTela_Controle.sContainer = objTela.Name Then
                                Set objControle.Container = objTela
                                iAchou = 1
                            End If
            
                        End If
                        
                        Exit For
                    End If
                Next
            Next
        
            If Len(Trim(objEdicaoTela_Tela.sTabIndex)) > 0 Then
                objEdicaoTela_Tela.sTabIndexPadrao = ""
                'Percorre Todos os controles para guardar os TabIndex Padrões
                For Each objControle In objTela.Controls
''                    If (Not (TypeOf objControle Is Label)) And (Not (TypeOf objControle Is PictureBox)) And (Not (TypeOf objControle Is Frame)) And (Not (TypeOf objControle Is SSFrame)) And (Not (TypeOf objControle Is SSPanel)) Then
                    If (Not (Controle_ObterNomeClasse(objControle) = "Label")) And (Not (Controle_ObterNomeClasse(objControle) = "PictureBox")) And (Not (Controle_ObterNomeClasse(objControle) = "Frame")) And (Not (Controle_ObterNomeClasse(objControle) = "CommonDialog")) And (Not (Controle_ObterNomeClasse(objControle) = "Image")) Then
                        objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " " & objControle.TabIndex
                    End If
                Next
                
                objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " "
                
                sTabIndexBuffer = objEdicaoTela_Tela.sTabIndex
                gobjEstInicial.List1.Clear
                Set colControles = New Collection
                iIndice = 0
                For Each objControle In objTela.Controls
''                    If (Not (TypeOf objControle Is Label)) And (Not (TypeOf objControle Is PictureBox)) And (Not (TypeOf objControle Is Frame)) And (Not (TypeOf objControle Is SSFrame)) And (Not (TypeOf objControle Is SSPanel)) Then
                    If (Not (Controle_ObterNomeClasse(objControle) = "Label")) And (Not (Controle_ObterNomeClasse(objControle) = "PictureBox")) And (Not (Controle_ObterNomeClasse(objControle) = "Frame")) And (Not (Controle_ObterNomeClasse(objControle) = "CommonDialog")) And (Not (Controle_ObterNomeClasse(objControle) = "Image")) Then
                        
                        If Len(Trim(sTabIndexBuffer)) <> 0 Then
                            iIndice = iIndice + 1
                            iPosicaoPrimeiroEspaco = InStr(1, sTabIndexBuffer, " ")
                            iPosicaoSegundoEspaco = InStr(2, sTabIndexBuffer, " ")
                            iTabIndex = CInt(Trim(Mid(sTabIndexBuffer, 2, (iPosicaoSegundoEspaco - iPosicaoPrimeiroEspaco) - 1)))
                        
                            sTabIndexBuffer = right(sTabIndexBuffer, Len(sTabIndexBuffer) - (iPosicaoSegundoEspaco - iPosicaoPrimeiroEspaco))
                        
                            colControles.Add objControle
                            Call gobjEstInicial.List1.AddItem(iTabIndex)
                            gobjEstInicial.List1.ItemData(gobjEstInicial.List1.NewIndex) = iIndice
                        End If
                    End If
                Next
                
                For iIndice = 0 To gobjEstInicial.List1.ListCount - 1
                    colControles.Item(gobjEstInicial.List1.ItemData(iIndice)).TabIndex = gobjEstInicial.List1.List(iIndice)
                Next
            End If
        
        End If
    Next
    
    Exit Function
    
Erro_Preenche_Propriedades_EdicaoTela:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142039)
    
    End Select
    
    Exit Function

End Function


Public Function Inicializa_Edicao_zOrder(objTela As Object) As Long
'Altera as propriedades dos Controles da Tela

Dim objEdicaoTela_Tela As New ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As New ClassEdicaoTela_Controle
Dim objControle As Object
Dim sControle As String
Dim iIndice2 As Integer

On Error GoTo Erro_Inicializa_Edicao_zOrder

    'Procura a Tela
    For Each objEdicaoTela_Tela In gcolEdicaoTela
    
        If objEdicaoTela_Tela.sNomeTela = objTela.Name Then
            
            'Para cada controle da coleção
            For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
                
                'Procura o controle respectivo na Tela
                For Each objControle In objTela.Controls
                    
                    iIndice2 = -1
                    iIndice2 = objControle.Index
                    
                    If iIndice2 > -1 Then
                        sControle = objControle.Name & "(" & objControle.Index & ")"
                    Else
                        sControle = objControle.Name
                    End If
                    
                    If sControle = objEdicaoTela_Controle.sNomeControle Then
                        
                        'Inserido por Wagner
                        '###############
                        If objEdicaoTela_Controle.izOrder <> -1 Then
                            objControle.ZOrder objEdicaoTela_Controle.izOrder
                        End If
                        '###############
                        
                        If objEdicaoTela_Controle.iHabilitado = MARCADO Then
                            objControle.Enabled = True
                        Else
                            objControle.Enabled = False
                        End If
                        
                    End If
                Next
            Next
        End If
    Next
    
    Exit Function
    
Erro_Inicializa_Edicao_zOrder:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142040)
    
    End Select
    
    Exit Function

End Function

'******** Fim Edicao de Telas **************************

Public Function Trata_Alteracao(objObjetoTela As Object, ParamArray avParametros() As Variant) As Long

Dim objObjetosBD As New ClassObjetoBD
Dim lErro As Long
Dim avParamFiltro() As Variant
Dim iIndMin As Integer, iIndMax As Integer
Dim iIndice As Integer
Dim Y As New AdmChamaTela

On Error GoTo Erro_Trata_Alteracao

    'Retorna o nome da Classe do Objeto em questão no objObjetosBD.sClasseObjeto
    objObjetosBD.sClasseObjeto = TypeName(objObjetoTela)
    
    If objObjetosBD.sClasseObjeto = "ClassTiposRegApuracao" Then objObjetosBD.sClasseObjeto = objObjetosBD.sClasseObjeto & objObjetoTela.sNomeTabela
    
    'Chama função de leitura da tabela ObjetosBD_Le
    lErro = CF("ObjetosBD_Le", objObjetosBD)
    If lErro <> SUCESSO And lErro <> 80430 Then gError 80431

    'Se a ClasseObjeto não foi encontrada
    If lErro = 80430 Then gError 80432
    
    'Verifica se a mensagem está configurada para informar aviso
    If objObjetosBD.iAvisaSobrePosicao <> DESMARCADO Then
    
        iIndMin = LBound(avParametros)
        iIndMax = UBound(avParametros)
        
        ReDim avParamFiltro(iIndMin To iIndMax)
        
        For iIndice = iIndMin To iIndMax
            avParamFiltro(iIndice) = avParametros(iIndice)
        Next
    
    
        'Verifica se é alteração ou inclusão
        lErro = CF("VerificaRegistro_ObjetosBD", objObjetosBD, avParamFiltro)
        If lErro <> SUCESSO And lErro <> 80446 Then gError 80447
    
        If lErro = SUCESSO Then
            lErro = Y.Chama_Tela("TrataAlteracao", objObjetosBD)
        Else
            giRetornoTela = vbOK
        End If
    
        'Verifica o retorno da tela TrataAlteracao
        If giRetornoTela = vbCancel Then gError 80433
    
    End If
    
    
Trata_Alteracao = SUCESSO

    Exit Function
    
Erro_Trata_Alteracao:

    Trata_Alteracao = gErr
    
    Select Case gErr
    
        Case 80431

        Case 80432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEOBJETO_INEXISTENTE", gErr, objObjetosBD.sClasseObjeto)

        Case 80433, 80447

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142041)

    End Select
    
    Exit Function

End Function
