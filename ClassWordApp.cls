VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWordApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim iPgmDoc As Integer

Dim objApp As Object
Dim objDispatcher As Object
Dim objWorkBook As Object
Dim objDoc As Object
Dim objFrame As Object
Dim objTable As Object, iLinhaTabela As Integer, iColunaTabela As Integer, sFonte As String, dFontH As Double, dFontW As Double, dColor As Double

'constantes do word - ini
Private Const wdCharacter = 1
Private Const wdGoToField = 7
Private Const wdWord9TableBehavior = 1
Private Const wdAutoFitContent = 1
Private Const wdGoToLine = 3
Private Const wdCell = 12
'constantes do word - fim

Private Function Inicializar() As Long

Dim lErro As Long
Dim sConteudo As String
Dim lRetorno As Long
Dim sPgm As String
Dim sUsaPgmPadrao As String

On Error GoTo Erro_Inicializar

    sUsaPgmPadrao = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "UsaPgmOfficePadrao", "1", sUsaPgmPadrao, 255, "ADM100.INI")
    sUsaPgmPadrao = left(sUsaPgmPadrao, lRetorno)
    
    If CInt(sUsaPgmPadrao) = MARCADO Then
    
        lErro = CF("Config_Le", "AdmConfig", "PGM_PADRAO_OFFICE", EMPRESA_TODA, sConteudo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iPgmDoc = CInt(sConteudo)
    
    Else

        sPgm = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "PgmOffice", "1", sPgm, 255, "ADM100.INI")
        sPgm = left(sPgm, lRetorno)
    
        Select Case UCase(sPgm)
        
            Case "OPENOFFICE", "OPEN OFFICE", "OO", CStr(PLANILHA_OO)
                iPgmDoc = PLANILHA_OO
            
            Case "LIBREOFFICE", "LIBRE OFFICE", "LO", CStr(PLANILHA_LO)
                iPgmDoc = PLANILHA_LO
                
            Case Else
                iPgmDoc = PLANILHA_MO
                   
        End Select
        
    End If
    
    Inicializar = SUCESSO
    
    Exit Function
    
Erro_Inicializar:
    
    Inicializar = gErr
    
    iPgmDoc = PLANILHA_MO
    
End Function

Private Function Finalizar() As Long

On Error GoTo Erro_Finalizar

    If Not (objApp Is Nothing) Then Set objApp = Nothing
    If Not (objDispatcher Is Nothing) Then Set objDispatcher = Nothing
    If Not (objWorkBook Is Nothing) Then Set objWorkBook = Nothing
    If Not (objDoc Is Nothing) Then Set objDoc = Nothing
    If Not (objFrame Is Nothing) Then Set objFrame = Nothing
    If Not (objTable Is Nothing) Then Set objTable = Nothing

    Finalizar = SUCESSO
    
    Exit Function
    
Erro_Finalizar:
    
    Finalizar = gErr
    
End Function

Public Function Abrir() As Long
'Cria uma instÂncia do Excel para uso pelo Corporator

On Error GoTo Erro_Abrir

    Call Inicializar

    'Se não existe uma instância do excel em aberto
    If objApp Is Nothing Then
        
        Select Case iPgmDoc
        
            Case PLANILHA_OO
                Set objApp = CreateObject("com.sun.star.ServiceManager")
                
                Set objDispatcher = OO_CreateUnoService("com.sun.star.frame.DispatchHelper")
                                
            Case PLANILHA_LO
            
            Case Else
                'Cria uma nova instância do Excel
                Set objApp = CreateObject("Word.Application")
    
        End Select
    
    End If
    
    Abrir = SUCESSO
    
    Exit Function
    
Erro_Abrir:
    
    Abrir = gErr
    
    Select Case gErr

        Case 429
            If iPgmDoc = PLANILHA_OO Then
                Call Rotina_Erro(vbOKOnly, "ERRO_OO_WRITER_NAO_INSTALADO", gErr, Error)
                Call OpenOffice_Trata_Config
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_WORD_NAO_INSTALADO", gErr, Error)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213513)
        
    End Select
    
    Call Fechar

End Function

Public Function Fechar() As Long

    'Se existe uma instância do excel em aberto
    If Not (objApp Is Nothing) Then
        
        Select Case iPgmDoc
        
            Case PLANILHA_OO
                If Not (objDoc Is Nothing) Then objDoc.Close (True)
                
            Case PLANILHA_LO
            
            Case Else
                If Not (objDoc Is Nothing) Then objDoc.Close (False)
    
        End Select
            
    End If
    
    Call Finalizar
    
    Fechar = SUCESSO
    
    Exit Function
    
Erro_Fechar:
    
    Fechar = gErr

On Error GoTo Erro_Fechar
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213514)
        
    End Select

End Function

Public Function Salvar(ByVal sNomeArq As String) As Long

On Error GoTo Erro_Salvar

    'Obtem a célula
    Select Case iPgmDoc
    
        Case PLANILHA_OO
        
            If UCase(right(sNomeArq, 3)) = "DOC" Or UCase(right(sNomeArq, 4)) = "DOC""" Then
                'Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq)), OO_MakePropertyValue("FilterName", "MS Word 97")))
                Call objDoc.storeAsURL(OO_ConvertToUrl(sNomeArq), Array(OO_MakePropertyValue("FilterName", "MS Word 97")))
            ElseIf UCase(right(sNomeArq, 3)) = "PDF" Or UCase(right(sNomeArq, 4)) = "PDF""" Then
                'Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq)), OO_MakePropertyValue("FilterName", "writer_pdf_Export")))
                Call objDoc.storeToURL(OO_ConvertToUrl(sNomeArq), Array(OO_MakePropertyValue("FilterName", "writer_pdf_Export")))
            Else
                'Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq))))
                Call objDoc.storeAsURL(OO_ConvertToUrl(sNomeArq), Array())
            End If
            
        Case PLANILHA_LO
    
        Case Else
            If UCase(right(sNomeArq, 3)) = "PDF" Or UCase(right(sNomeArq, 4)) = "PDF""" Then

                objDoc.ExportAsFixedFormat OutputFileName:=Replace(sNomeArq, """", ""), ExportFormat:=17
            Else
                objDoc.SaveAs FileName:=sNomeArq, FileFormat:=0
            End If
    
    End Select
    
    Salvar = SUCESSO
    
    Exit Function
    
Erro_Salvar:
    
    Salvar = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213561)
        
    End Select
    
    Exit Function
    
End Function

Public Function Mudar_Visibilidade(ByVal bVisivel As Boolean) As Long

On Error GoTo Erro_Mudar_Visibilidade

    'Obtem a célula
    Select Case iPgmDoc
    
        Case PLANILHA_OO
            objDoc.getCurrentController.getFrame.getContainerWindow.setVisible bVisivel
                        
        Case PLANILHA_LO
    
        Case Else
            objApp.Visible = bVisivel
    
    End Select
    
    Mudar_Visibilidade = SUCESSO
    
    Exit Function
    
Erro_Mudar_Visibilidade:
    
    Mudar_Visibilidade = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213560)
        
    End Select
    
End Function

Public Function Abrir_Doc(ByVal sNomeArq As String) As Long

Dim dVersaoWord As Double

On Error GoTo Erro_Abrir_Doc

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO
            Set objWorkBook = OO_CreateUnoService("com.sun.star.frame.Desktop") 'cria o objeto desktop
    
            If UCase(right(sNomeArq, 3)) = "DOC" Then
                Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("FilterName", "MS Word 97"), OO_MakePropertyValue("Hidden", True)))
            Else
                Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("Hidden", True)))
            End If
            
            Set objFrame = objDoc.CurrentController.Frame

        Case PLANILHA_LO
    
        Case Else
            
            dVersaoWord = 0
            If IsNumeric(objApp.Version) Then
                dVersaoWord = CDbl(Replace(objApp.Version, ".", ","))
            End If
        
            If dVersaoWord < 15 Then
                Set objDoc = objApp.Documents.Open(sNomeArq, , True)
            Else
                Set objDoc = objApp.Documents.Open(sNomeArq)
            End If
          
    End Select
    
    Abrir_Doc = SUCESSO
    
    Exit Function
    
Erro_Abrir_Doc:
    
    Abrir_Doc = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select
    
End Function

Public Function DOC_Cria_Tabela(ByVal lIndiceFF As Long, ByVal iNumLinhas As Integer, ByVal iNumColunas As Integer) As Long
    
Dim objCampoForm As Object
Dim vEnum, vVal
Dim lQtdeAux As Long
Dim objAnchor
Dim iIndice As Integer

On Error GoTo Erro_DOC_Cria_Tabela

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO

            Set vEnum = objDoc.getTextFields().createEnumeration()
            
            If Not IsNull(vEnum) Then
                Do While vEnum.hasMoreElements()
                    Set vVal = vEnum.nextElement()
                    lQtdeAux = lQtdeAux + 1
                    If lIndiceFF = lQtdeAux Then
                        If vVal.supportsService("com.sun.star.text.TextField.Input") Then
                            
                            Set objAnchor = vVal.getAnchor()
                            
                            dFontW = objAnchor.CharWeight
                            sFonte = objAnchor.CharFontName
                            dFontH = objAnchor.CharHeight
                            dColor = objAnchor.CharColor
                            
                            Call vVal.setPropertyValue("Content", "")
                            Call objDoc.TextFields.Refresh
                            Call vVal.dispose
                            Call objDoc.TextFields.Refresh
                            
                            Set objTable = objDoc.createInstance("com.sun.star.text.TextTable")
                            Call objTable.Initialize(iNumLinhas, iNumColunas)
                            
                            Call objAnchor.GetText().insertTextContent(objAnchor, objTable, False)
                            
                            iColunaTabela = 0
                            iLinhaTabela = 0
                            
                        End If
                    End If
                Loop
            End If
            
        Case PLANILHA_LO
    
        Case Else
            
            Set objCampoForm = objDoc.FormFields(lIndiceFF)
            Call objCampoForm.Select
            
            'objWord.selection.GoToNext wdGoToField
            objApp.selection.MoveRight wdCharacter, 1
        
            objDoc.Tables.Add objApp.selection.Range, iNumLinhas, iNumColunas, wdWord9TableBehavior, wdAutoFitContent
            objApp.selection.Tables(1).ApplyStyleHeadingRows = True
            objApp.selection.Tables(1).ApplyStyleLastRow = True
            objApp.selection.Tables(1).ApplyStyleFirstColumn = True
            objApp.selection.Tables(1).ApplyStyleLastColumn = True
            
    End Select
    
    DOC_Cria_Tabela = SUCESSO
    
    Exit Function
    
Erro_DOC_Cria_Tabela:
    
    DOC_Cria_Tabela = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select
    
    Exit Function

End Function

Public Function DOC_Insere_Cabec_Tabela(ParamArray avParams()) As Long

Dim iIndice As Integer
Dim bBold As Boolean
Dim objCampoForm As Object
Dim objRange As Object

On Error GoTo Erro_DOC_Insere_Cabec_Tabela

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO
        
            iLinhaTabela = 0
            
            For iIndice = 0 To UBound(avParams)
                Call objTable.getCellByPosition(iColunaTabela, iLinhaTabela).setString(avParams(iIndice))
                iColunaTabela = iColunaTabela + 1
            Next
            
            Set objRange = objTable.getCellRangeByName("A" & CStr(iLinhaTabela + 1) & ":" & Chr(64 + iColunaTabela) & CStr(iLinhaTabela + 1))
            objRange.CharWeight = 150
            objRange.CharFontName = sFonte
            objRange.CharHeight = dFontH
            objRange.CharColor = dColor
            iLinhaTabela = iLinhaTabela + 1

        Case PLANILHA_LO
    
        Case Else
            
            bBold = objApp.selection.Font.Bold
            For iIndice = 0 To UBound(avParams)
        
                objApp.selection.Font.Bold = True
                objApp.selection.TypeText avParams(iIndice)
                
                'If iIndice <> UBound(avParams) Then
                    objApp.selection.MoveRight wdCharacter, 1
                'End If
        
            Next
            objApp.selection.Font.Bold = bBold
            
    End Select
    
    DOC_Insere_Cabec_Tabela = SUCESSO
    
    Exit Function
    
Erro_DOC_Insere_Cabec_Tabela:
    
    DOC_Insere_Cabec_Tabela = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function

End Function

Public Function DOC_Insere_Valores_Tabela(ParamArray avParams()) As Long

Dim iIndice As Integer
Dim objCampoForm As Object
Dim objRange As Object

On Error GoTo Erro_DOC_Insere_Valores_Tabela

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO

            iColunaTabela = 0
            
            For iIndice = 0 To UBound(avParams)
                Call objTable.getCellByPosition(iColunaTabela, iLinhaTabela).setString(avParams(iIndice))
                iColunaTabela = iColunaTabela + 1
            Next
            Set objRange = objTable.getCellRangeByName("A" & CStr(iLinhaTabela + 1) & ":" & Chr(64 + iColunaTabela) & CStr(iLinhaTabela + 1))
            objRange.CharWeight = dFontW
            objRange.CharFontName = sFonte
            objRange.CharHeight = dFontH
            objRange.CharColor = dColor
            iLinhaTabela = iLinhaTabela + 1
                                                       
            If iLinhaTabela >= objTable.getrows().Count Then
                            
                Set objRange = objTable.getCellRangeByName("A1:" & Chr(64 + iColunaTabela) & CStr(iLinhaTabela))
                            
                Call objDoc.CurrentController.Select(objRange)
                
                Call objDispatcher.executeDispatch(objDoc.CurrentController.Frame, ".uno:SetOptimalColumnWidth", "", 0, Array())
                
            End If
            
        Case PLANILHA_LO
    
        Case Else
            
            'Set objCampoForm = objDoc.FormFields(lIndiceFF)
            'Call objCampoForm.Select
            
            objApp.selection.MoveRight wdCell
            'objWord.selection.GoToNext wdGoToLine
        
            For iIndice = 0 To UBound(avParams)
        
                objApp.selection.TypeText avParams(iIndice)
                
                'If iIndice <> UBound(avParams) Then
                    objApp.selection.MoveRight wdCharacter, 1
                'End If
                
            Next
            
    End Select
    
    DOC_Insere_Valores_Tabela = SUCESSO
    
    Exit Function
    
Erro_DOC_Insere_Valores_Tabela:
    
    DOC_Insere_Valores_Tabela = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function

End Function

Public Function DOC_Insere_Valores_Tabela2(ParamArray avParams()) As Long

Dim iIndice As Integer
Dim objCampoForm As Object
Dim objRange As Object

On Error GoTo Erro_DOC_Insere_Valores_Tabela2

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO
            
            For iIndice = 0 To UBound(avParams)
                Call objTable.getCellByPosition(iColunaTabela, iLinhaTabela - 1).setString(avParams(iIndice))
                iColunaTabela = iColunaTabela + 1
            Next
            Set objRange = objTable.getCellRangeByName("A" & CStr(iLinhaTabela) & ":" & Chr(64 + iColunaTabela) & CStr(iLinhaTabela))
            objRange.CharWeight = dFontW
            objRange.CharFontName = sFonte
            objRange.CharHeight = dFontH
            objRange.CharColor = dColor
            
        Case PLANILHA_LO
    
        Case Else
            
            For iIndice = 0 To UBound(avParams)
        
                objApp.selection.TypeText avParams(iIndice)
                
                'If iIndice <> UBound(avParams) Then
                    objApp.selection.MoveRight wdCharacter, 1
                'End If
                
            Next
            
    End Select
    
    DOC_Insere_Valores_Tabela2 = SUCESSO
    
    Exit Function
    
Erro_DOC_Insere_Valores_Tabela2:
    
    DOC_Insere_Valores_Tabela2 = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function

End Function

Public Function DOC_Insere_Figura(ByVal sNomeFigura As String) As Long

Dim iIndice As Integer
Dim objCampoForm As Object, iColunaTabela As Integer
Dim objRange As Object

On Error GoTo Erro_DOC_Insere_Figura

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO
            
        Case PLANILHA_LO
    
        Case Else
            objApp.selection.GoToNext wdGoToField
            objApp.selection.InlineShapes.AddPicture sNomeFigura, False, True
           
    End Select
    
    DOC_Insere_Figura = SUCESSO
    
    Exit Function
    
Erro_DOC_Insere_Figura:
    
    DOC_Insere_Figura = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function

End Function

Public Function Qtde_FormFields() As Long

Dim lQtdeAux As Long
Dim iIndice As Integer
Dim vEnum, vVal

On Error GoTo Erro_Qtde_FormFields

    lQtdeAux = 0

    Select Case iPgmDoc
    
        Case PLANILHA_OO
        
            Set vEnum = objDoc.getTextFields().createEnumeration()
            
            If Not IsNull(vEnum) Then
                Do While vEnum.hasMoreElements()
                    Set vVal = vEnum.nextElement()
                    lQtdeAux = lQtdeAux + 1
                Loop
            End If
            
        Case PLANILHA_LO
    
        Case Else
            lQtdeAux = objDoc.FormFields.Count
            
    End Select
    
    Qtde_FormFields = lQtdeAux
    
    Exit Function
    
Erro_Qtde_FormFields:
    
    Qtde_FormFields = 0
    
    Exit Function
    
End Function

Public Function FormField_Seleciona(ByVal lIndiceFF As Long) As Long

Dim objCampoForm As Object

On Error GoTo Erro_FormField_Seleciona

    Select Case iPgmDoc
    
        Case PLANILHA_OO

        Case PLANILHA_LO
    
        Case Else
            Set objCampoForm = objDoc.FormFields(lIndiceFF)
            Call objCampoForm.Select
            
    End Select
    
    FormField_Seleciona = SUCESSO
    
    Exit Function
    
Erro_FormField_Seleciona:
    
    FormField_Seleciona = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function
    
End Function

Public Function FormField_Obtem_Nome(ByVal lIndiceFF As Long, ByRef sNome As String) As Long

Dim objCampoForm As Object
Dim vEnum, vVal
Dim lQtdeAux As Long

On Error GoTo Erro_FormField_Obtem_Nome

    sNome = ""

    Select Case iPgmDoc
    
        Case PLANILHA_OO

            Set vEnum = objDoc.getTextFields().createEnumeration()
            
            If Not IsNull(vEnum) Then
                Do While vEnum.hasMoreElements()
                    Set vVal = vEnum.nextElement()
                    lQtdeAux = lQtdeAux + 1
                    If lIndiceFF = lQtdeAux Then
                        sNome = vVal.getPropertyValue("Hint")
                        Exit Do
                    End If
                Loop
            End If
            
        Case PLANILHA_LO
    
        Case Else
            Set objCampoForm = objDoc.FormFields(lIndiceFF)
            sNome = objCampoForm.Name
            
    End Select
    
    FormField_Obtem_Nome = SUCESSO
    
    Exit Function
    
Erro_FormField_Obtem_Nome:
    
    FormField_Obtem_Nome = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function
    
End Function

Public Function FormField_Preenche_Valor(ByVal lIndiceFF As Long, ByVal vValor As Variant) As Long

Dim objCampoForm As Object
Dim vEnum, vVal
Dim lQtdeAux As Long
Dim objAnchor, objCursor
Dim oTextField

On Error GoTo Erro_FormField_Preenche_Valor

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmDoc
    
        Case PLANILHA_OO

            Set vEnum = objDoc.getTextFields().createEnumeration()
            
            If Not IsNull(vEnum) Then
                Do While vEnum.hasMoreElements()
                    Set vVal = vEnum.nextElement()
                    lQtdeAux = lQtdeAux + 1
                    If lIndiceFF = lQtdeAux Then
                        If vVal.supportsService("com.sun.star.text.TextField.Input") Then
                            
                            Set objAnchor = vVal.getAnchor()

                            Set objCursor = objAnchor.GetText().createTextCursorByRange(objAnchor)
                            
                            Call vVal.setPropertyValue("Content", "")
                            Call objDoc.TextFields.Refresh
                            Call vVal.dispose
                            Call objDoc.TextFields.Refresh
                            
                            objCursor.String = vValor
                            
                        End If
                    End If
                Loop
            End If
            
        Case PLANILHA_LO
    
        Case Else
            Set objCampoForm = objDoc.FormFields(lIndiceFF)
            Call objCampoForm.Select
            objCampoForm.Range.Text = vValor
            
    End Select
    
    FormField_Preenche_Valor = SUCESSO
    
    Exit Function
    
Erro_FormField_Preenche_Valor:
    
    FormField_Preenche_Valor = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select

    Exit Function
    
End Function

Private Sub OpenOffice_Trata_Config()

Dim lRet As Long, vbResult As VbMsgBoxResult
Dim sNomeArqTeste As String, sDummy As String
Dim sOOWriterExec As String

On Error GoTo Erro_OpenOffice_Trata_Config

    sNomeArqTeste = App.Path & "\OpenOffice.odt"
    
    If Dir(sNomeArqTeste) <> "" Then

        vbResult = Rotina_Aviso(vbYesNo, "AVISO_OPENOFFICE_CONFIG")
        If vbResult = vbYes Then

            'Localiza o exe do Writer do OpenOffice
            sOOWriterExec = Space(255)
            lRet = FindExecutable(sNomeArqTeste, sDummy, sOOWriterExec)
            sOOWriterExec = Trim(sOOWriterExec)
        
            'Abre o doc de controle
            lRet = ShellExecute(HWND_TOPMOST, "open", sOOWriterExec, sNomeArqTeste, sDummy, SW_NORMAL)
            
        End If

    End If

    Exit Sub
    
Erro_OpenOffice_Trata_Config:
    
    Exit Sub

End Sub

Private Function OO_CreateUnoService(ByVal strServiceName As String) As Object
    Set OO_CreateUnoService = objApp.createInstance(strServiceName)
End Function

Private Function OO_MakePropertyValue(ByVal cName As String, ByVal uValue As Variant) As Object
    
  Dim objPropertyValue As Object

  Set objPropertyValue = objApp.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  
  objPropertyValue.Name = cName
  objPropertyValue.Value = uValue
      
  Set OO_MakePropertyValue = objPropertyValue

End Function

Private Function OO_ConvertToUrl(ByVal strFile As String) As String
    strFile = Replace(strFile, "\", "/")
    strFile = Replace(strFile, ":", "|")
    strFile = Replace(strFile, " ", "%20")
    strFile = "file:///" + strFile
    OO_ConvertToUrl = strFile
End Function
