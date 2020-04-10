VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassExcelApp"
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

Private Const EXCEL_XLCATEGORY = 1

Dim iPgmPlanilha As Integer

Dim objApp As Object
Dim objDispatcher As Object
Dim objWorkBook As Object
Dim objDoc As Object
Dim objFrame As Object
Dim objSheet As Object
Dim objSheetGraf As Object
'Dim objRange As Object
'Dim objCell As Object
Dim objGrafico As Object
Dim objLocalSettings As Object

Dim asColFormat(1 To 100) As String
Dim alColFormatNumber(1 To 100) As Long

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
        
        iPgmPlanilha = CInt(sConteudo)
    
    Else

        sPgm = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "PgmOffice", "1", sPgm, 255, "ADM100.INI")
        sPgm = left(sPgm, lRetorno)
    
        Select Case UCase(sPgm)
        
            Case "OPENOFFICE", "OPEN OFFICE", "OO", CStr(PLANILHA_OO)
                iPgmPlanilha = PLANILHA_OO
            
            Case "LIBREOFFICE", "LIBRE OFFICE", "LO", CStr(PLANILHA_LO)
                iPgmPlanilha = PLANILHA_LO
                
            Case Else
                iPgmPlanilha = PLANILHA_MO
                   
        End Select
        
    End If
    
    Inicializar = SUCESSO
    
    Exit Function
    
Erro_Inicializar:
    
    Inicializar = gErr
    
    iPgmPlanilha = PLANILHA_MO
    
End Function

Private Function Finalizar() As Long

On Error GoTo Erro_Finalizar

    If Not (objGrafico Is Nothing) Then Set objGrafico = Nothing
    If Not (objSheetGraf Is Nothing) Then Set objSheetGraf = Nothing
    
'    If Not (objRange Is Nothing) Then Set objRange = Nothing
'    If Not (objCell Is Nothing) Then Set objCell = Nothing
    
    If Not (objSheet Is Nothing) Then Set objSheet = Nothing
    If Not (objFrame Is Nothing) Then Set objFrame = Nothing
    If Not (objDoc Is Nothing) Then Set objDoc = Nothing
    If Not (objWorkBook Is Nothing) Then Set objWorkBook = Nothing
    
    If Not (objLocalSettings Is Nothing) Then Set objLocalSettings = Nothing
    If Not (objDispatcher Is Nothing) Then Set objDispatcher = Nothing
    If Not (objApp Is Nothing) Then Set objApp = Nothing

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
        
        Select Case iPgmPlanilha
        
            Case PLANILHA_OO
                Set objApp = CreateObject("com.sun.star.ServiceManager")
                
                Set objDispatcher = OO_CreateUnoService("com.sun.star.frame.DispatchHelper")
                
                Set objLocalSettings = objApp.Bridge_GetStruct("com.sun.star.lang.Locale")

                objLocalSettings.Language = "pt"
                objLocalSettings.Country = "br"
                
            Case PLANILHA_LO
            
            Case Else
                'Cria uma nova instância do Excel
                Set objApp = CreateObject("Excel.Application")
    
        End Select
    
    End If
    
    Abrir = SUCESSO
    
    Exit Function
    
Erro_Abrir:
    
    Abrir = gErr
    
    Select Case gErr

        Case 429
            If iPgmPlanilha = PLANILHA_OO Then
                Call Rotina_Erro(vbOKOnly, "ERRO_OO_CALC_NAO_INSTALADO", gErr, Error)
                Call OpenOffice_Trata_Config
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_NAO_INSTALADO", gErr, Error)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213513)
        
    End Select
    
    Call Fechar

End Function

Public Function Fechar() As Long

On Error GoTo Erro_Fechar

    'Se existe uma instância do excel em aberto
    If Not (objApp Is Nothing) Then
        
        Select Case iPgmPlanilha
        
            Case PLANILHA_OO
                If Not (objDoc Is Nothing) Then objDoc.Close (True)
                
            Case PLANILHA_LO
            
            Case Else
                objApp.Quit
    
        End Select
            
    End If
    
    Call Finalizar
    
    Fechar = SUCESSO
    
    Exit Function
    
Erro_Fechar:
    
    Fechar = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213514)
        
    End Select

End Function

Public Function Criar_Planilha() As Long

On Error GoTo Erro_Criar_Planilha

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objWorkBook = objApp.createInstance("com.sun.star.frame.Desktop") 'cria o objeto desktop
    
            Set objDoc = objWorkBook.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array(OO_MakePropertyValue("Hidden", True)))
            
            'objDoc.getCurrentController.getFrame.getContainerWindow.setVisible False
            
            Set objFrame = objDoc.CurrentController.Frame
            
            Set objSheet = objDoc.getSheets().getByIndex(0) 'seleciona primeira planilha
            
        Case PLANILHA_LO
    
        Case Else
            
            'Desabilita as mensagens do Excel
            objApp.DisplayAlerts = False
            
            'Cria uma nova pasta de trabalho
            Set objWorkBook = objApp.Workbooks.Add
            
            'Seleciona a planilha ativa na pasta de trabalho criada
            Set objSheet = objWorkBook.ActiveSheet
          
    End Select
    
    Criar_Planilha = SUCESSO
    
    Exit Function
    
Erro_Criar_Planilha:
    
    Criar_Planilha = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213515)
        
    End Select
    
End Function

Public Function Abrir_Planilha_CSV(ByVal sNomeArq As String) As Long

On Error GoTo Erro_Abrir_Planilha_CSV

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objWorkBook = objApp.createInstance("com.sun.star.frame.Desktop") 'cria o objeto desktop
    
            Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("FilterName", "Text - txt - csv (StarCalc)"), OO_MakePropertyValue("FilterOptions", "" & CStr(Asc(";")) & ",34,0,1"), OO_MakePropertyValue("Hidden", True)))
                       
            Set objFrame = objDoc.CurrentController.Frame
            
            Set objSheet = objDoc.getSheets().getByIndex(0) 'seleciona primeira planilha
            
        Case PLANILHA_LO
    
        Case Else
            
            Call objApp.Workbooks.OpenText(sNomeArq, Origin:=2, DataType:=1, TextQualifier:=1, Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:=".")
               
            Set objWorkBook = objApp.ActiveWorkBook
            
            'Seleciona a planilha ativa na pasta de trabalho criada
            Set objSheet = objWorkBook.ActiveSheet
          
    End Select
    
    Abrir_Planilha_CSV = SUCESSO
    
    Exit Function
    
Erro_Abrir_Planilha_CSV:
    
    Abrir_Planilha_CSV = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select
    
End Function

Public Function Abrir_Planilha_CSV2(ByVal sNomeArq As String) As Long

On Error GoTo Erro_Abrir_Planilha_CSV2

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objWorkBook = objApp.createInstance("com.sun.star.frame.Desktop") 'cria o objeto desktop
    
            Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("FilterName", "Text - txt - csv (StarCalc)"), OO_MakePropertyValue("FilterOptions", "" & CStr(Asc(";")) & ",34,0,1"), OO_MakePropertyValue("Hidden", True)))
                       
            Set objFrame = objDoc.CurrentController.Frame
            
            Set objSheet = objDoc.getSheets().getByIndex(0) 'seleciona primeira planilha
            
        Case PLANILHA_LO
    
        Case Else
            
            Call objApp.Workbooks.OpenText(sNomeArq, Origin:=2, DataType:=1, TextQualifier:=1, Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:=".", _
                    FieldInfo:=Array(Array(1, 1), Array(2, 1), _
                            Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
                            Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
                            16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
                            Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
                            29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
                            Array(36, 1), Array(37, 4), Array(38, 4), Array(39, 1), Array(40, 1), Array(41, 1), _
                            Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), _
                            Array(48, 1), Array(49, 1), Array(50, 1)))
               
            Set objWorkBook = objApp.ActiveWorkBook
            
            'Seleciona a planilha ativa na pasta de trabalho criada
            Set objSheet = objWorkBook.ActiveSheet
          
    End Select
    
    Abrir_Planilha_CSV2 = SUCESSO
    
    Exit Function
    
Erro_Abrir_Planilha_CSV2:
    
    Abrir_Planilha_CSV2 = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select
    
End Function

Public Function Abrir_Planilha(ByVal sNomeArq As String) As Long

On Error GoTo Erro_Abrir_Planilha

    'Obter pasta de trabalho e aba do excel
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objWorkBook = objApp.createInstance("com.sun.star.frame.Desktop") 'cria o objeto desktop
               
            If UCase(right(sNomeArq, 3)) = "XLS" Then
                Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("FilterName", "MS Excel 97"), OO_MakePropertyValue("Hidden", True)))
            Else
                Set objDoc = objWorkBook.loadComponentFromURL(OO_ConvertToUrl(sNomeArq), "_blank", 0, Array(OO_MakePropertyValue("Hidden", True)))
            End If
            
            Set objFrame = objDoc.CurrentController.Frame
            
            Set objSheet = objDoc.getSheets().getByIndex(0) 'seleciona primeira planilha
            
        Case PLANILHA_LO
    
        Case Else
            
            'Desabilita as mensagens do Excel
            objApp.DisplayAlerts = False
            
            Call objApp.Workbooks.Open(sNomeArq)
               
            Set objWorkBook = objApp.ActiveWorkBook
            
            'Seleciona a 1a planilha
            Set objSheet = objWorkBook.Sheets(1)
          
    End Select
    
    Abrir_Planilha = SUCESSO
    
    Exit Function
    
Erro_Abrir_Planilha:
    
    Abrir_Planilha = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)
        
    End Select
    
End Function

Public Function Preencher_Celula(ByVal lLinha As Long, ByVal lColuna As Long, ByVal vValor As Variant, Optional ByVal sFormato As String = "", Optional ByVal bNegrito As Boolean = False, Optional ByVal sFonte As String = "", Optional ByVal dFonteTamanho As Double = 0, Optional ByVal vAlinhamentoHorizontal As Variant = -1) As Long

Dim objCell As Object

On Error GoTo Erro_Preencher_Celula

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objCell = objSheet.getCellByPosition(lColuna - 1, lLinha - 1)
            
            If sFormato <> "" Then
                If vValor <> "" Then
                    objCell.Value = vValor
                End If
                objCell.NumberFormat = OO_FindCreateNumberFormatStyle(sFormato)
            Else
                objCell.String = vValor
            End If
            If bNegrito Then objCell.CharWeight = 150  ';Bold
            
            If sFonte <> "" Then
                objCell.CharFontName = sFonte
            End If
            If dFonteTamanho <> 0 Then
                objCell.CharHeight = dFonteTamanho
            End If
'            If Not (vAlinhamentoHorizontal Is Nothing) Then
            If vAlinhamentoHorizontal <> -1 Then
            
                Select Case vAlinhamentoHorizontal
                    Case ALINHAMENTO_CENTRALIZADO
                        objCell.HoriJustify = 2  'Center
                        
                    Case ALINHAMENTO_DIREITA
                        objCell.HoriJustify = 0
                    
                    Case ALINHAMENTO_ESQUERDA
                        objCell.HoriJustify = 1
                    
                End Select
            End If
                        
        Case PLANILHA_LO
    
        Case Else
            'Monta a Faixa de Células da coluna atual
            Set objCell = objSheet.cells(lLinha, lColuna)
            
            objCell.Value = vValor
            If sFormato <> "" And sFormato <> "dd/mm/yyyy" Then objCell.NumberFormat = sFormato
            If bNegrito Then objCell.Font.Bold = True
            
            If sFonte <> "" Then
                objCell.Font.Name = sFonte
            End If
            If dFonteTamanho <> 0 Then
                objCell.Font.Size = dFonteTamanho
            End If
            'If Not (vAlinhamentoHorizontal Is Nothing) Then
            If vAlinhamentoHorizontal <> -1 Then
                objCell.HorizontalAlignment = vAlinhamentoHorizontal
            End If
    
    End Select
    
    Preencher_Celula = SUCESSO
    
    Exit Function
    
Erro_Preencher_Celula:
    
    Preencher_Celula = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213517)
        
    End Select
    
    Exit Function
    
End Function

Public Function Obtem_Valor_Celula(ByVal lLinha As Long, ByVal lColuna As Long) As Variant
'Obtem o valor da célula pela posição (linha/coluna)
Dim vValor As Variant
Dim objCell As Object
Dim sAux As String

    vValor = ""

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objCell = objSheet.getCellByPosition(lColuna - 1, lLinha - 1)
            
            If objCell.Type = 0 Then vValor = "" 'Vazia
            If objCell.Type = 1 Then vValor = objCell.Value
            If objCell.Type = 2 Then
                sAux = objCell.String
                vValor = sAux
            End If
            If objCell.Type = 3 Then vValor = objCell.Value
                        
        Case PLANILHA_LO
    
        Case Else
            'Monta a Faixa de Células da coluna atual
            Set objCell = objSheet.cells(lLinha, lColuna)
            
            If Not IsError(objCell.Value) Then vValor = objCell.Value
    
    End Select
    
    Obtem_Valor_Celula = vValor
        
End Function

Public Function Obtem_Valor_Celula2(ByVal sRange As String) As Variant
'Obtem o valor da célula pela posição (A1, B5, etc)
Dim vValor As Variant
Dim objCell As Object
Dim sAux As String

    vValor = ""

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objCell = objSheet.getCellRangeByName(sRange)
            
            If objCell.Type = 0 Then vValor = "" 'Vazia
            If objCell.Type = 1 Then vValor = objCell.Value
            If objCell.Type = 2 Then
                sAux = objCell.String
                vValor = sAux
            End If
            If objCell.Type = 3 Then vValor = objCell.Value
                        
        Case PLANILHA_LO
    
        Case Else
            'Monta a Faixa de Células da coluna atual
            vValor = objSheet.Range(sRange)
    
    End Select
    
    Obtem_Valor_Celula2 = vValor
        
End Function

Public Function Obtem_NumberFormat_Celula(ByVal lLinha As Long, ByVal lColuna As Long) As Variant

Dim vValor As Variant
Dim objNumberFormat As Object
Dim objCell As Object, sAux As String

    vValor = ""

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objCell = objSheet.getCellByPosition(lColuna - 1, lLinha - 1)
        
            If objCell.NumberFormat <> 0 Then
                Set objNumberFormat = objDoc.NumberFormats.getByKey(objCell.NumberFormat)
                sAux = objNumberFormat.FormatString
                
                sAux = Replace(sAux, ".", "x")
                sAux = Replace(sAux, ",", ".")
                sAux = Replace(sAux, "x", ",")
                sAux = Replace(sAux, "AAAA", "YYYY")
                sAux = Replace(sAux, "AA", "YY")
                sAux = Replace(sAux, "aa", "yy")
                sAux = Replace(sAux, "aaaa", "yyyy")
                vValor = sAux
            End If
                        
        Case PLANILHA_LO
    
        Case Else
            'Monta a Faixa de Células da coluna atual
            Set objCell = objSheet.cells(lLinha, lColuna)
            
            If Not IsError(objCell.Value) Then vValor = objCell.NumberFormat
    
    End Select
    
    Obtem_NumberFormat_Celula = vValor
        
End Function

Public Function Ajustar_Tam_Colunas() As Long

On Error GoTo Erro_Ajustar_Tam_Colunas

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            objSheet.Columns().OptimalWidth = True

        Case PLANILHA_LO
    
        Case Else
            objApp.cells.Select
            objApp.cells.EntireColumn.AutoFit
            objApp.Range("A1", "A1").Select
    
    End Select
    
    Ajustar_Tam_Colunas = SUCESSO
    
    Exit Function
    
Erro_Ajustar_Tam_Colunas:
    
    Ajustar_Tam_Colunas = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213518)
        
    End Select
    
End Function

Public Function Imprimir(Optional ByVal iOrientacao As Integer = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE) As Long

Dim avArg(0)  As Object
Dim objPrinter As Object
Dim iIndice As Integer

On Error GoTo Erro_Imprimir

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            
            objDoc.getCurrentController.getFrame.getContainerWindow.setVisible True
            
            If iOrientacao = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE Then
                Set avArg(0) = OO_MakePropertyValue("PaperOrientation", 1)
                objDoc.Printer = avArg()
            End If
           
            Call objDispatcher.executeDispatch(objFrame, ".uno:PrintDefault", "", 0, Array(OO_MakePropertyValue("Wait", True)))
            
            Call Sleep(20000)
                         
            objDoc.getCurrentController.getFrame.getContainerWindow.setVisible False
                         
        Case PLANILHA_LO
    
        Case Else
            'Indica que a planilha será impressa no estilo Landscape
            objSheet.PageSetup.Orientation = iOrientacao
            'Indica que a planilha será impressa com linhas de grade
            objSheet.PageSetup.PrintGridlines = EXCEL_PLANILHA_IMPRIME_GRADE
            'Imprime o gráfico e fecha o excel
            objSheet.PrintOut
    
    End Select
    
    Imprimir = SUCESSO
    
    Exit Function
    
Erro_Imprimir:
    
    Imprimir = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213519)
        
    End Select
    
End Function

Public Function Imprimir_Grafico() As Long

Dim avArg(0)  As Object
Dim objPrinter As Object
Dim iIndice As Integer
Dim objFrameGraf As Object

On Error GoTo Erro_Imprimir_Grafico

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objDoc.CurrentController.ActiveSheet = objSheetGraf
            
            Set objFrameGraf = objDoc.CurrentController.Frame
            
            objDoc.getCurrentController.getFrame.getContainerWindow.setVisible True
            
            Set avArg(0) = OO_MakePropertyValue("PaperOrientation", 1)
            objDoc.Printer = avArg()
           
            Call objDispatcher.executeDispatch(objFrameGraf, ".uno:PrintDefault", "", 0, Array(OO_MakePropertyValue("Wait", True)))
            
            Call Sleep(20000)

            objDoc.getCurrentController.getFrame.getContainerWindow.setVisible False
                         
        Case PLANILHA_LO
    
        Case Else
            objGrafico.PrintOut
    
    End Select
    
    Imprimir_Grafico = SUCESSO
    
    Exit Function
    
Erro_Imprimir_Grafico:
    
    Imprimir_Grafico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213519)
        
    End Select
    
End Function

Public Function Mudar_Visibilidade(ByVal bVisivel As Boolean) As Long

On Error GoTo Erro_Mudar_Visibilidade

    'Obtem a célula
    Select Case iPgmPlanilha
    
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

Public Function Salvar(ByVal sNomeArq As String) As Long

Dim iPos As Integer
Dim sExtensao As String

On Error GoTo Erro_Salvar

    iPos = InStr(1, sNomeArq, ".")
    If iPos <> 0 Then
        sExtensao = Mid(sNomeArq, iPos + 1)
    End If
    
    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            'Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq & ".ods")), OO_MakePropertyValue("FilterName", "calc8")))
            'Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq))))
                        
            Select Case UCase(sExtensao)
            
                Case "XLS"
                    Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq)), OO_MakePropertyValue("FilterName", "MS Excel 97")))
                           
                Case "ODS"
                    Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq)), OO_MakePropertyValue("FilterName", "calc8")))
                
                Case "TXT"
                    Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq)), OO_MakePropertyValue("FilterName", "Text - txt - csv (StarCalc)")))
            
                Case Else
                    If iPos = 0 Then
                        Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq & ".xls")), OO_MakePropertyValue("FilterName", "MS Excel 97")))
                    Else
                        Call objDispatcher.executeDispatch(objFrame, ".uno:SaveAs", "", 0, Array(OO_MakePropertyValue("URL", OO_ConvertToUrl(sNomeArq))))
                    End If
            End Select
            
                        
        Case PLANILHA_LO
    
        Case Else
            If iPos = 0 Then
                objWorkBook.SaveAs sNomeArq & ".xls"
            Else
                objWorkBook.SaveAs sNomeArq
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
    
End Function

Public Function Muda_Nome_Aba(ByVal sNomeAba As String) As Long

On Error GoTo Erro_Muda_Nome_Aba

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Call objDispatcher.executeDispatch(objFrame, ".uno:RenameTable", "", 0, Array(OO_MakePropertyValue("Name", sNomeAba)))
                        
        Case PLANILHA_LO
    
        Case Else
            objSheet.Name = sNomeAba
    
    End Select
    
    Muda_Nome_Aba = SUCESSO
    
    Exit Function
    
Erro_Muda_Nome_Aba:
    
    Muda_Nome_Aba = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213562)
        
    End Select
    
End Function

Public Function Muda_Nome_Grafico(ByVal sNomeGrafico As String) As Long

Dim objFrameGraf As Object

On Error GoTo Erro_Muda_Nome_Grafico

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objDoc.CurrentController.ActiveSheet = objSheetGraf
            
            Set objFrameGraf = objDoc.CurrentController.Frame
            
            Call objDispatcher.executeDispatch(objFrameGraf, ".uno:RenameTable", "", 0, Array(OO_MakePropertyValue("Name", sNomeGrafico)))
                        
        Case PLANILHA_LO
    
        Case Else
            objGrafico.Name = sNomeGrafico
    
    End Select
    
    Muda_Nome_Grafico = SUCESSO
    
    Exit Function
    
Erro_Muda_Nome_Grafico:
    
    Muda_Nome_Grafico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213562)
        
    End Select
    
End Function

Public Function Fazer_Merge(ByVal lLIni As Long, ByVal lLFim As Long, ByVal lCIni As Long, ByVal lCFim As Long) As Long

Dim objRange As Object

On Error GoTo Erro_Fazer_Merge

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            Set objRange = objSheet.getCellRangeByPosition(lCIni - 1, lLIni - 1, lCFim - 1, lLFim - 1)
            Call objRange.Merge(True)
                        
        Case PLANILHA_LO
    
        Case Else
            Set objRange = objSheet.Range(objSheet.cells(lLIni, lCIni), objSheet.cells(lLFim, lCFim))
            Call objRange.Merge
    
    End Select
    
    Fazer_Merge = SUCESSO
    
    Exit Function
    
Erro_Fazer_Merge:
    
    Fazer_Merge = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213563)
        
    End Select
    
End Function

Public Function Apagar_Abas() As Long

Dim iIndice As Integer

On Error GoTo Erro_Apagar_Abas

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            For iIndice = objWorkBook.Sheets.Count To 2 Step -1
                objWorkBook.Sheets.Item(iIndice).Delete
            Next
    
    End Select
    
    Apagar_Abas = SUCESSO
    
    Exit Function
    
Erro_Apagar_Abas:
    
    Apagar_Abas = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213564)
        
    End Select
    
End Function

Public Function Criar_Aba() As Long

Dim iIndice As Integer

On Error GoTo Erro_Criar_Aba

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            If objWorkBook.Sheets.Count <= 1 Then
                objWorkBook.Sheets.Add After:=objWorkBook.Sheets.Item(1)
            End If
            Set objSheet = objWorkBook.Sheets.Item(2)
    
    End Select
    
    Criar_Aba = SUCESSO
    
    Exit Function
    
Erro_Criar_Aba:
    
    Criar_Aba = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213565)
        
    End Select
    
End Function

Public Function Criar_TabelaDinamica(ByVal lLIni As Long, ByVal lLFim As Long, ByVal lCIni As Long, ByVal lCFim As Long) As Long

Dim iIndice As Integer
Dim iVersao As Integer, sVAux As String, iPos As Integer

On Error GoTo Erro_Criar_TabelaDinamica

'sErroAux = "Criar_TabelaDinamica"

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
        
'            Versão 8 para o Excel 97
'            Versão 9 para o Excel 2000
'            Versão 10 para o Excel 2002
'            Versão 11 para o Excel 2003
'            Versão 12 para o Excel 2007
'            Versão 14 para o Excel 2010
'            Versão 15 para o Excel 2013

'sErroAux = sErroAux & ";" & "Excel Versão: " & objApp.Version
'
'sErroAux = sErroAux & ";" & "lLIni: " & CStr(lLIni)
'sErroAux = sErroAux & ";" & "lLFim: " & CStr(lLFim)
'sErroAux = sErroAux & ";" & "lCIni: " & CStr(lCIni)
'sErroAux = sErroAux & ";" & "lCFim: " & CStr(lCFim)

            iPos = InStr(1, objApp.Version, ".")
            If iPos <> 0 Then
                sVAux = left(objApp.Version, iPos - 1)
            Else
                sVAux = objApp.Version
            End If
            If IsNumeric(sVAux) Then
                iVersao = CInt(sVAux)
            Else
                iVersao = 15
            End If

            'If objApp.Version <= 11 Then
            If iVersao <= 11 Then
                
'sErroAux = sErroAux & ";" & "Comando: objWorkBook.PivotCaches.Add "
                
                Call objWorkBook.PivotCaches.Add(SourceType:=EXCEL_TABDIN_TIPO_DATABASE, SourceData:="Dados!R" & CStr(lLIni) & "C1:R" & CStr(lLFim) & "C" & CStr(lCFim)).CreatePivotTable(TableDestination:="Tabela!L1C1", TableName:="TabDin1")
            Else
                
'sErroAux = sErroAux & ";" & "Comando: objWorkBook.PivotCaches.Create "
                
                Call objWorkBook.PivotCaches.Create(SourceType:=EXCEL_TABDIN_TIPO_DATABASE, SourceData:="Dados!R" & CStr(lLIni) & "C1:R" & CStr(lLFim) & "C" & CStr(lCFim)).CreatePivotTable(TableDestination:="Tabela!R1C1", TableName:="TabDin1")
            
            End If
    
    End Select
    
    Criar_TabelaDinamica = SUCESSO
    
    Exit Function
    
Erro_Criar_TabelaDinamica:
    
    Criar_TabelaDinamica = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
'    Call Rotina_Erro(vbOKOnly, sErroAux, 990001)
    
End Function

Public Function Criar_Grafico(ByVal lLIni As Long, ByVal lLFim As Long, ByVal lCIni As Long, ByVal lCFim As Long, Optional ByVal sTituloGrafico As String = "", Optional ByVal vPlotLinhaColuna As Variant = Nothing, Optional ByVal lPosicaoLegenda As Long = 0, Optional ByVal iEixosGrafico As Integer = 0, Optional ByVal lLabelsXPosicao As Long = 0, Optional ByVal lLabelsXOrientacao As Long = 0) As Long
'Inclui o Gráfico a partir dos dados passados
Dim iIndice As Integer
Dim objRectangle As Object, objCellRangeAddress(0) As Object, objCharts As Object

On Error GoTo Erro_Criar_Grafico

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objCellRangeAddress(0) = objApp.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            objCellRangeAddress(0).Sheet = 0
            objCellRangeAddress(0).StartColumn = lCIni - 1
            objCellRangeAddress(0).StartRow = lLIni - 1
            objCellRangeAddress(0).EndColumn = lCFim - 1
            objCellRangeAddress(0).EndRow = lLFim - 1
        
            'Set objRectangle = objApp.Bridge_GetStruct("com.sun.star.awt.Rectangle")
            Set objRectangle = objDoc.getSheets().getByIndex(1).Bridge_GetStruct("com.sun.star.awt.Rectangle")
            objRectangle.X = 0
            objRectangle.Y = 0
            objRectangle.Width = 24000
            objRectangle.Height = 12000
           
            Set objSheetGraf = objDoc.getSheets().getByIndex(1)

            Set objCharts = objSheetGraf.getCharts
        
            Call objCharts.addNewByName("Grafico", objRectangle, objCellRangeAddress(), True, False)
            
            Set objGrafico = objCharts.getByName("Grafico").EmbeddedObject
            
            objGrafico.Title.String = sTituloGrafico
            objGrafico.Title.CharWeight = 150
            
            'Se o gráfico foi setado para não exibir legenda
            If lPosicaoLegenda = EXCEL_LEGENDA_NAO_EXIBE Then
    
                'Informa ao excel para ocultar a legenda
                objGrafico.HasLegend = False
    
            'Se foi setado para exibir a legenda
            Else
    
                'Informa ao excel para exibir a legenda
                objGrafico.HasLegend = True
    
            End If
                                    
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            Set objGrafico = objWorkBook.charts.Add
            
            objGrafico.ChartWizard objSheet.Range(objSheet.cells(lLIni, lCIni), objSheet.cells(lLFim, lCFim)), , , vPlotLinhaColuna, , , , sTituloGrafico
            
            'objGrafico.ChartType = iTipoGrafico
            
            'Se o gráfico foi setado para não exibir legenda
            If lPosicaoLegenda = EXCEL_LEGENDA_NAO_EXIBE Then
    
                'Informa ao excel para ocultar a legenda
                objGrafico.HasLegend = False
    
            'Se foi setado para exibir a legenda
            Else
    
                'Informa ao excel para exibir a legenda
                objGrafico.HasLegend = True
    
                'Posiciona a legenda do gráfico
                objGrafico.Legend.Position = lPosicaoLegenda
    
            End If
            
            'Se o gráfico, é um gráfico que possui eixos
            If iEixosGrafico = EXCEL_GRAFICO_COM_EIXOS Then
            
                'Determina a posição dos labels do eixo X
                objGrafico.Axes(EXCEL_XLCATEGORY).TickLabelPosition = lLabelsXPosicao
                
                'Determina a orientação dos labels do eixo X
                objGrafico.Axes(EXCEL_XLCATEGORY).TickLabels.Orientation = lLabelsXOrientacao
            End If
        
            
    End Select
    
    Criar_Grafico = SUCESSO
    
    Exit Function
    
Erro_Criar_Grafico:
    
    Criar_Grafico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_LeftHeader(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_LeftHeader

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            objGrafico.PageSetup.LeftHeader = sTexto
    
    End Select
    
    Grafico_Set_LeftHeader = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_LeftHeader:
    
    Grafico_Set_LeftHeader = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_CenterHeader(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_CenterHeader

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            objGrafico.PageSetup.CenterHeader = sTexto
    
    End Select
    
    Grafico_Set_CenterHeader = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_CenterHeader:
    
    Grafico_Set_CenterHeader = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_RightHeader(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_RightHeader

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            objGrafico.PageSetup.RightHeader = sTexto
    
    End Select
    
    Grafico_Set_RightHeader = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_RightHeader:
    
    Grafico_Set_RightHeader = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_LeftFooter(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_LeftFooter

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            objGrafico.PageSetup.LeftFooter = sTexto
    
    End Select
    
    Grafico_Set_LeftFooter = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_LeftFooter:
    
    Grafico_Set_LeftFooter = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_CenterFooter(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_CenterFooter

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            Set objGrafico.PageSetup.CenterFooter = sTexto
    
    End Select
    
    Grafico_Set_CenterFooter = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_CenterFooter:
    
    Grafico_Set_CenterFooter = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Grafico_Set_RightFooter(ByVal sTexto As String) As Long

On Error GoTo Erro_Grafico_Set_RightFooter

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            'Adiciona um gráfico à planilha
            Set objGrafico.PageSetup.RightFooter = sTexto
    
    End Select
    
    Grafico_Set_RightFooter = SUCESSO
    
    Exit Function
    
Erro_Grafico_Set_RightFooter:
    
    Grafico_Set_RightFooter = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213566)
        
    End Select
    
End Function

Public Function Incluir_Campo_TabelaDinamica(ByVal sCampo As String, ByVal iPosicao As Integer, ByVal iFormula As Integer, lLinha As Long, lColuna As Long, lFiltro As Long, lValor As Long) As Long

Dim iIndice As Integer

On Error GoTo Erro_Incluir_Campo_TabelaDinamica

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            Select Case iPosicao
                Case EXCEL_TABDIN_POS_LINHA
                    lLinha = lLinha + 1
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Orientation = EXCEL_TABDIN_POS_LINHA
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Position = lLinha
                Case EXCEL_TABDIN_POS_COLUNA
                    lColuna = lColuna + 1
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Orientation = EXCEL_TABDIN_POS_COLUNA
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Position = lColuna
                Case EXCEL_TABDIN_POS_FILTRO
                    lFiltro = lFiltro + 1
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Orientation = EXCEL_TABDIN_POS_FILTRO
                    objSheet.PivotTables("TabDin1").PivotFields(sCampo).Position = lFiltro
                Case EXCEL_TABDIN_POS_VALOR
                    lValor = lValor + 1
                    Select Case iFormula
                        Case EXCEL_FORMULA_SUM
                            objSheet.PivotTables("TabDin1").AddDataField objSheet.PivotTables("TabDin1").PivotFields(sCampo), "Soma de " & sCampo, EXCEL_FORMULA_SUM
                        Case EXCEL_FORMULA_AVG
                            objSheet.PivotTables("TabDin1").AddDataField objSheet.PivotTables("TabDin1").PivotFields(sCampo), "Média de " & sCampo, EXCEL_FORMULA_AVG
                        Case EXCEL_FORMULA_MIN
                            objSheet.PivotTables("TabDin1").AddDataField objSheet.PivotTables("TabDin1").PivotFields(sCampo), "Min. de " & sCampo, EXCEL_FORMULA_MIN
                        Case EXCEL_FORMULA_MAX
                            objSheet.PivotTables("TabDin1").AddDataField objSheet.PivotTables("TabDin1").PivotFields(sCampo), "Máx. de " & sCampo, EXCEL_FORMULA_MAX
                        Case EXCEL_FORMULA_COUNT
                            objSheet.PivotTables("TabDin1").AddDataField objSheet.PivotTables("TabDin1").PivotFields(sCampo), "Qtd de " & sCampo, EXCEL_FORMULA_COUNT
                    End Select
            End Select
    End Select
    
    Incluir_Campo_TabelaDinamica = SUCESSO
    
    Exit Function
    
Erro_Incluir_Campo_TabelaDinamica:
    
    Incluir_Campo_TabelaDinamica = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213567)
        
    End Select
    
End Function

Public Function Incluir_Campo_Grafico(ByVal lNumRegs As Long, ByVal lColunaY As Long, ByVal lColunaX As Long, ByVal iSerie As Integer, ByVal lTipoGraficoColuna As Long, Optional ByVal lDataLabels As Long = EXCEL_NAO_EXIBE_LABELS, Optional ByVal lDataLabelsOrientacao As Long = -1) As Long

Dim iIndice As Integer
Dim oDataProvider As Object
Dim oDiagram As Object
Dim oCooSys, oCoods
Dim oDataSeriesList
Dim oSeries As Object
Dim sRangeX As String, sRangeY As String, sTitulo As String
Dim oNewDataSeriesList() As Object
Dim oChartTypes, oChartType
'Dim oYAxis As Object, oXAxis As Object
Dim oAxis As Object, oScaleData As Object
Dim oCategorySequence As Object

On Error GoTo Erro_Incluir_Campo_Grafico

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            ' adding new series
            Set oDataProvider = objGrafico.getDataProvider()
            
'            com.sun.star.chart.BarDiagram
'            com.sun.star.chart.AreaDiagram
'            com.sun.star.chart.LineDiagram
'            com.sun.star.chart.PieDiagram
'            com.sun.star.chart.DonutDiagram
'            com.sun.star.chart.NetDiagram
'            com.sun.star.chart.XYDiagram
'            com.sun.star.chart.StockDiagram

''Constantes para setagem de configurações do gráfico no excel
''tipos de gráfico
'Public Const EXCEL_GRAFICO_COLUMN_CLUSTERED = 51 'Coluna
'Public Const EXCEL_GRAFICO_3D_COLUMN_CLUSTERED = 54 'Coluna 3D
'Public Const EXCEL_GRAFICO_3D_COLUMN_STACKED = 55
'Public Const EXCEL_GRAFICO_3DPIE = -4102 'Torta 3D
'Public Const EXCEL_GRAFICO_LINE_MARKERS = 65 'Linha com marcas
'Public Const EXCEL_GRAFICO_3D_LINE = -4101 'Linha 3D

            If iSerie = 1 Then
                Select Case lTipoGraficoColuna
                    Case EXCEL_GRAFICO_3D_LINE, EXCEL_GRAFICO_LINE_MARKERS
                        Set objGrafico.Diagram = objGrafico.createInstance("com.sun.star.chart.LineDiagram")
                        If lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE Then
                            objGrafico.Diagram.Dim3D = True
                            objGrafico.Diagram.Deep = False
                            'objGrafico.Diagram.RightAngledAxes = True 'needs OpenOffice.org 2.3 or newer
                            'objGrafico.Diagram.D3DScenePerspective = com.sun.star.drawing.ProjectionMode.Perspective
                            'objGrafico.Diagram.Perspective = 100 'needs OpenOffice.org 2.4.1 or newer
                            'objGrafico.Diagram.RotationHorizontal = 60 'needs OpenOffice.org 2.4.1 or newer
                            'objGrafico.Diagram.RotationVertical = 30 'needs OpenOffice.org 2.4.1 or newer
                        End If
                    Case EXCEL_GRAFICO_3DPIE
                        Set objGrafico.Diagram = objGrafico.createInstance("com.sun.star.chart.PieDiagram")
                        If lTipoGraficoColuna = EXCEL_GRAFICO_3DPIE Then
                            objGrafico.Diagram.Dim3D = True
                        End If
                    Case EXCEL_GRAFICO_COLUMN_CLUSTERED, EXCEL_GRAFICO_3D_COLUMN_CLUSTERED, EXCEL_GRAFICO_3D_COLUMN_STACKED
                        Set objGrafico.Diagram = objGrafico.createInstance("com.sun.star.chart.BarDiagram")
                        If lTipoGraficoColuna = EXCEL_GRAFICO_3D_COLUMN_CLUSTERED Or lTipoGraficoColuna = EXCEL_GRAFICO_3D_COLUMN_STACKED Then
                            objGrafico.Diagram.Dim3D = True
                        End If
                End Select
            End If
            
            Set oDiagram = objGrafico.getFirstDiagram() ' chart2
            
            oCooSys = oDiagram.getCoordinateSystems()
            Set oCoods = oCooSys(0) ' this chart has only a coordinate system
                        
            oChartTypes = oCoods.getChartTypes() ' chart type one by one
            Set oChartType = oChartTypes(0)
        
            If iSerie = 1 Then
                oChartType.setDataSeries (Array())
                
                With objGrafico
                    .Diagram.SplineType = 0
                    .Diagram.HasXAxisTitle = True
                    .Diagram.HasYAxisTitle = True
                    .Diagram.XAxisTiTle.CharHeight = 8
                    .Diagram.YAxisTiTle.CharHeight = 8
                    .Diagram.YAxis.CharHeight = 8
                    .Diagram.XAxis.CharHeight = 8
                    .Diagram.SymbolType = -3
                    .Diagram.HasYAxisDescription = True
                End With
                
            End If
             
            oDataSeriesList = oChartType.getDataSeries()
            
            sRangeX = objSheet.Name & "." & Chr$(64 + lColunaX) & "2:" & Chr$(64 + lColunaX) & CStr(lNumRegs)
            sRangeY = objSheet.Name & "." & Chr$(64 + lColunaY) & "2:" & Chr$(64 + lColunaY) & CStr(lNumRegs)
            sTitulo = objSheet.Name & "." & Chr$(64 + lColunaY) & "1" 'objSheet.getCellByPosition(lColunaY - 1, 0).String
            
            Set oSeries = OO_CreateDataSeries_XYDiagram(oDataProvider, sRangeX, sRangeY, sTitulo)
            
            Select Case iSerie
            
                Case 2
                    ' set series properties see com.sun.star.chart2.DataSeries
                    'oSeries1.Color = 17798
                    oSeries.Color = 16728590
                    
                Case 3
                    oSeries.Color = 16765728
            
            End Select

            ReDim oNewDataSeriesList(LBound(oDataSeriesList) To UBound(oDataSeriesList) + 1)
             
            For iIndice = LBound(oNewDataSeriesList) To UBound(oNewDataSeriesList)
                If iIndice = iSerie - 1 Then
                    Set oNewDataSeriesList(iIndice) = oSeries
                Else
                    Set oNewDataSeriesList(iIndice) = oDataSeriesList(iIndice)
                End If
            Next
            
            ' update chart (only the charttype is updated)
            oChartType.setDataSeries (oNewDataSeriesList)
            
'            ' modify axises
'            Set oYAxis = oCoods.getAxisByDimension(1, 0)
'            Set oXAxis = oCoods.getAxisByDimension(0, 0)
'
'            Call OO_SetSimpleTitle(oYAxis, "Y", True)
'            Call OO_SetSimpleTitle(oXAxis, "X", True)
            
            Set oAxis = oCoods.getAxisByDimension(0, 0)
            
            Set oCategorySequence = OO_CreateUnoService("com.sun.star.chart2.data.LabeledDataSequence")

            Call oCategorySequence.setValues(OO_CreateDataSequence(oDataProvider, sRangeX, "categories"))
            
            Set oScaleData = oAxis.getScaleData()
            Set oScaleData.Categories = oCategorySequence
            Call oAxis.setScaleData(oScaleData)
            
            'If lDataLabels <> EXCEL_NAO_EXIBE_LABELS Then objGrafico.Diagram.XAxis.DisplayLabels = True
            
        Case PLANILHA_LO
    
        Case Else
        
            If objGrafico.SeriesCollection.Count < iSerie Then
                Call objGrafico.SeriesCollection.NewSeries
            End If
            
            objGrafico.SeriesCollection(iSerie).XValues = objSheet.Range(objSheet.cells(2, lColunaX), objSheet.cells(lNumRegs, lColunaX))
            objGrafico.SeriesCollection(iSerie).Values = objSheet.Range(objSheet.cells(2, lColunaY), objSheet.cells(lNumRegs, lColunaY))
        
            'Define o nome da série de dados
            objGrafico.SeriesCollection(iSerie).Name = "=" & Chr$(34) & objSheet.Range(objSheet.cells(1, lColunaY), objSheet.cells(1, lColunaY)) & Chr$(34)

            'Define o tipo do gráfico da série de dados
            objGrafico.SeriesCollection(iSerie).ChartType = lTipoGraficoColuna

            'Define o que será exibido como label da série (Valor, Legenda, Percentual, Legenda e Percentual, Não exibe)
            objGrafico.SeriesCollection(iSerie).ApplyDataLabels (lDataLabels)

            'Se os labels de valores do eixo X forem exibidos => define a orientação dos labels
            If lDataLabels <> EXCEL_NAO_EXIBE_LABELS Then objGrafico.SeriesCollection(iSerie).DataLabels.Orientation = lDataLabelsOrientacao

    End Select
    
    Incluir_Campo_Grafico = SUCESSO
    
    Exit Function
    
Erro_Incluir_Campo_Grafico:
    
    Incluir_Campo_Grafico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213567)
        
    End Select
    
    Exit Function
    
End Function

Public Function Incluir_Grafico(ByVal iTipoGrafico As Integer) As Long
'Inclui o Gráfico a partir de uma tabela dinâmica padrão
Dim iIndice As Integer

On Error GoTo Erro_Incluir_Grafico

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
            objWorkBook.charts.Add
            Set objGrafico = objWorkBook.charts.Item(1)
            objGrafico.SetSourceData Source:=objWorkBook.Sheets("Tabela").PivotTables("TabDin1").rowrange.currentregion
            objGrafico.ChartType = iTipoGrafico
            objGrafico.Location 1, "Gráfico"
            objWorkBook.Sheets("Gráfico").Move After:=objWorkBook.Sheets("Tabela")

            objWorkBook.ShowPivotTableFieldList = True
            
    End Select
    
    Incluir_Grafico = SUCESSO
    
    Exit Function
    
Erro_Incluir_Grafico:
    
    Incluir_Grafico = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213568)
        
    End Select
    
    Exit Function
    
End Function

Public Function Preencher_Celula_Formula(ByVal lLIni As Long, ByVal lLFim As Long, ByVal lColuna As Long, ByVal iFormula As Integer, Optional ByVal sFormato As String = "") As Long

Dim objCell As Object

On Error GoTo Erro_Preencher_Celula_Formula

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            'FALTOU FAZER
                        
        Case PLANILHA_LO
    
        Case Else
        
            Set objCell = objSheet.cells(lLFim + 1, lColuna)
            Select Case iFormula
                Case EXCEL_FORMULA_SUM
                    objCell.FormulaR1C1 = "=SUM(R" & CStr(lLIni + 1) & "C" & CStr(lColuna) & ":R" & CStr(lLFim) & "C" & CStr(lColuna) & ")"
                Case EXCEL_FORMULA_AVG
                    objCell.FormulaR1C1 = "=AVERAGE(R" & CStr(lLIni + 1) & "C" & CStr(lColuna) & ":R" & CStr(lLFim) & "C" & CStr(lColuna) & ")"
                Case EXCEL_FORMULA_MAX
                    objCell.FormulaR1C1 = "=MAX(R" & CStr(lLIni + 1) & "C" & CStr(lColuna) & ":R" & CStr(lLFim) & "C" & CStr(lColuna) & ")"
                Case EXCEL_FORMULA_MIN
                    objCell.FormulaR1C1 = "=MIN(R" & CStr(lLIni + 1) & "C" & CStr(lColuna) & ":R" & CStr(lLFim) & "C" & CStr(lColuna) & ")"
                Case EXCEL_FORMULA_COUNT
                    objCell.FormulaR1C1 = "=COUNTA(R" & CStr(lLIni + 1) & "C" & CStr(lColuna) & ":R" & CStr(lLFim) & "C" & CStr(lColuna) & ")"
            End Select
            objCell.Font.Bold = True
            If sFormato <> "" Then objCell.NumberFormat = sFormato
    
    End Select
    
    Preencher_Celula_Formula = SUCESSO
    
    Exit Function
    
Erro_Preencher_Celula_Formula:
    
    Preencher_Celula_Formula = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213569)
        
    End Select
    
End Function

Public Function Qtde_Abas() As Integer

Dim iNumAbas As Integer

    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
            iNumAbas = objDoc.getSheets.Count
                        
        Case PLANILHA_LO
    
        Case Else
            iNumAbas = objWorkBook.Sheets.Count
            
    End Select
    
    Qtde_Abas = iNumAbas
    
    Exit Function

End Function

Public Function Planilha_Muda_Aba(ByVal iAba As Integer) As Long

On Error GoTo Erro_Planilha_Muda_Aba

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
        
            Set objSheet = objDoc.getSheets().getByIndex(iAba - 1)
                                 
        Case PLANILHA_LO
    
        Case Else
        
            Set objSheet = objWorkBook.Sheets(iAba)
    
    End Select
    
    Planilha_Muda_Aba = SUCESSO
    
    Exit Function
    
Erro_Planilha_Muda_Aba:
    
    Planilha_Muda_Aba = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213519)
        
    End Select
    
End Function

Public Function Atualiza_TabsDim() As Long

Dim iTabelas As Integer
Dim objTabelaDinamica As Object

On Error GoTo Erro_Atualiza_TabsDim

    'Obtem a célula
    Select Case iPgmPlanilha
    
        Case PLANILHA_OO
                                 
        Case PLANILHA_LO
    
        Case Else
    
            'Para cada Tabela dinâmica
            For iTabelas = 1 To objSheet.PivotTables.Count
            
                'Obtém a Tabela
                Set objTabelaDinamica = objSheet.PivotTables.Item(iTabelas)
                
                'Atualiza a Tabela
                objTabelaDinamica.PivotCache.Refresh
            
            Next
            
    End Select
    
    Atualiza_TabsDim = SUCESSO
    
    Exit Function
    
Erro_Atualiza_TabsDim:
    
    Atualiza_TabsDim = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213519)
        
    End Select
    
End Function

Private Sub OpenOffice_Trata_Config()

Dim lRet As Long, vbResult As VbMsgBoxResult
Dim sNomeArqTeste As String
Dim sOOWriterExec As String, sDummy As String

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

Private Function OO_FindCreateNumberFormatStyle(ByVal sFormat As String) As Long

Dim lFormatNum As Long
Dim sFormatAux As String
Dim iIndice As Integer

    sFormatAux = Replace(sFormat, ".", "x")
    sFormatAux = Replace(sFormatAux, ",", ".")
    sFormatAux = Replace(sFormatAux, "x", ",")
    sFormatAux = Replace(sFormatAux, "yyyy", "AAAA")
    sFormatAux = Replace(sFormatAux, "YYYY", "AAAA")
    sFormatAux = Replace(sFormatAux, "yy", "AA")
    sFormatAux = Replace(sFormatAux, "YY", "AA")
    sFormat = UCase(sFormatAux)
    
    lFormatNum = -1
  
    lFormatNum = objDoc.NumberFormats.queryKey(sFormat, objLocalSettings, True)

    If (lFormatNum = -1) Then
        For iIndice = LBound(asColFormat) To UBound(asColFormat)
            If asColFormat(iIndice) = "" Then Exit For
            If sFormat = asColFormat(iIndice) Then
                lFormatNum = alColFormatNumber(iIndice)
                Exit For
            End If
        Next
    End If

    If (lFormatNum = -1) Then
        lFormatNum = objDoc.NumberFormats.addNew(sFormat, objLocalSettings)
        If (lFormatNum = -1) Then
            lFormatNum = 0
        Else
            For iIndice = LBound(asColFormat) To UBound(asColFormat)
                If asColFormat(iIndice) = "" Then
                    asColFormat(iIndice) = sFormat
                    alColFormatNumber(iIndice) = lFormatNum
                    Exit For
                End If
            Next
        End If
    End If
    
    OO_FindCreateNumberFormatStyle = lFormatNum
  
End Function

' create new series for XYDiagram
' a series has two or more LabeledDataSequence that is named by Label DataSequence.
' oDataProvider: com.sun.star.chart2.data.XDataProvider
' sXRangeRepresentation: X
' sYRangeRepresentation: Y
' sLabelRangeRepresentation: range representation for name of the series
Private Function OO_CreateDataSeries_XYDiagram(oDataProvider As Object, sXRangeRepresentation As String, sYRangeRepresentation As String, Optional sLabelRangeRepresentation As String) As Object
  
  Dim oNewDataSeries As Object
  ' create new DataSeries
  Dim oDataY, oDataX, aData, oSequenceY, oSequenceX, oSequenceLabel
  
  Set oNewDataSeries = OO_CreateUnoService("com.sun.star.chart2.DataSeries")
  
  Dim oData(1) As Object ' x and y: .chart2.data.XLabeledDataSequence
  
  ' Y
  Set oDataY = OO_CreateUnoService("com.sun.star.chart2.data.LabeledDataSequence")
  Set oSequenceY = OO_CreateDataSequence(oDataProvider, sYRangeRepresentation, "values-y")
  If Not IsNull(oSequenceY) Then
    oDataY.setValues (oSequenceY)
  
    If Not ((IsMissing(sLabelRangeRepresentation)) And (sLabelRangeRepresentation <> "")) Then
      Set oSequenceLabel = OO_CreateDataSequence(oDataProvider, sLabelRangeRepresentation, "")
      oDataY.setLabel (oSequenceLabel) ' label is used as name
    End If
  End If
  
  ' X
  Set oDataX = OO_CreateUnoService("com.sun.star.chart2.data.LabeledDataSequence")
  Set oSequenceX = OO_CreateDataSequence(oDataProvider, sXRangeRepresentation, "values-x")
  If Not IsNull(oSequenceX) Then
    oDataX.setValues (oSequenceX)
  End If
  
  ' set x and y data to series
  aData = Array(oDataY, oDataX)
  oNewDataSeries.SetData (aData)
  
  Set OO_CreateDataSeries_XYDiagram = oNewDataSeries
  
End Function

' creat new DataSequence from range representaion
' that provides real data and its role in the series
' oDataProvider: com.sun.star.chart2.data.XDataProvider
' sRangeRepresentation: range address e.g. Sheet1.A1:B2
' sRole: role is defined in com.sun.star.chart2.data.DataSequenceRole
Private Function OO_CreateDataSequence(oDataProvider As Object, sRangeRepresentation As String, sRole As String) As Object
  
Dim oDataSequence As Object
  
  On Error GoTo Erro_OO_CreateDataSequence
  
  ' create .chart2.data.DataSequence from range representation
  Set oDataSequence = oDataProvider.createDataSequenceByRangeRepresentation(sRangeRepresentation)
  If Not IsNull(oDataSequence) Then
    oDataSequence.Role = sRole
  End If
  
Erro_OO_CreateDataSequence:
    Set OO_CreateDataSequence = oDataSequence
End Function

' create and set a simple title object for axis
' oAxis: axis of the coordinate
' sTitle: title string
' bDisplayLabels: set to show/hide its title
Private Sub OO_SetSimpleTitle(oAxis As Object, sTitle As String, Optional bDisplayLabels As Boolean)
Dim oNewTitle, oNewPart
  Set oNewTitle = OO_CreateUnoService("com.sun.star.chart2.Title")
  Set oNewPart = OO_CreateUnoService("com.sun.star.chart2.FormattedString")
  oNewPart.setString (sTitle)
  oNewTitle.SetText (Array(oNewPart))
  'If Not IsMissing(DisplayLabels) Then
  oAxis.DisplayLabels = bDisplayLabels
  oAxis.setTitleObject (oNewTitle)
End Sub

Private Sub Class_Terminate()
    'MsgBox ("excelapp terminate " & CStr(objApp Is Nothing))
    Call Finalizar
End Sub
