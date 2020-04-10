VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const EXCEL_STRING_NOME_PLANILHA = 31

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Public Function Excel_Gera_Planilha(ByVal objPlanilha As ClassPlanilhaExcel) As Long

Dim lErro As Long
Dim iLinha As Long
Dim iColuna As Integer
Dim objColuna As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
'Dim avValoresCelulas() As Variant

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
'Dim objGrafico As Object
Dim iLin As Long
Dim iCol As Integer
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Gera_Planilha

'    'Abre o excel
'    lErro = Excel_Abrir()
'    If lErro <> SUCESSO Then gError 102106

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    'Cria uma nova pasta de trabalho
'    Set objPastaTrabalho = GL_objExcelSist.Workbooks.Add
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
'
'    'Redimensiona o array de acordo a quantidade de linhas e colunas recebidas
'    ReDim avValoresCelulas(1 To objPlanilha.colColunas(1).colCelulas.Count, 1 To objPlanilha.colColunas.Count)
'
    lErro = objExcelApp.Criar_Planilha()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Para cada coluna que deverá receber dados
    For iColuna = 1 To objPlanilha.colColunas.Count
        
        'Instancia objColuna como um obj da Coleção de colunas
        Set objColuna = objPlanilha.colColunas(iColuna)
        
        'E para cada linha que receberá dados
        For iLinha = 1 To objColuna.colCelulas.Count
            
            'Instancia objCelulas como um obj da Coleção de células
            Set objCelulas = objColuna.colCelulas(iLinha)
            
            'Monta o array que será usado para preencher a planilha
            'avValoresCelulas(iLinha, iColuna) = objCelulas.vValor
            
            lErro = objExcelApp.Preencher_Celula(iLinha, iColuna, objCelulas.vValor, objCelulas.sNumberFormat, objCelulas.bFonteNegrito)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
        
'        'Monta a Faixa de Células da coluna atual
'        Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(1, iColuna), objPlanilhaExcel.Cells(objColuna.colCelulas.Count, iColuna))
'
'        'Altera a largura da coluna conforme a largura passada
'        objFaixaCelulasPreencher.ColumnWidth = objColuna.dLarguraColuna
    
    Next

'    'Monta a Faixa das Células que receberão os valores passado
'    Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(1, 1), objPlanilhaExcel.Cells(objPlanilha.colColunas(1).colCelulas.Count, objPlanilha.colColunas.Count))
'
'    'Passa os valores armazenados no array para a faixa determinada acima
'    objFaixaCelulasPreencher.Value = avValoresCelulas
'
'    '########################################################
'    'Inserido por Wagner 01/03/2007
'    For iCol = 1 To objPlanilha.colColunas.Count
'
'        For iLin = 1 To objPlanilha.colColunas(iCol).colCelulas.Count
'
'            If objPlanilha.colColunas(iCol).colCelulas.Item(iLin).bFonteNegrito Then
'                objFaixaCelulasPreencher(iLin, iCol).Font.Bold = True
'            End If
'
'            If Len(Trim(objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat)) > 0 Then
'                objFaixaCelulasPreencher(iLin, iCol).NumberFormat = objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat
'            End If
'        Next
'
'    Next
'    '########################################################

    'Se foi passado um nome para a planilha
    If Len(Trim(objPlanilha.sNomePlanilha)) > 0 Then
    
'        'Exibe o nome da planilha no excel
'        objPlanilhaExcel.Name = ExcelSubstitui_NomeInvalido(Mid(objPlanilha.sNomePlanilha, 1, EXCEL_STRING_NOME_PLANILHA))
    
        lErro = objExcelApp.Muda_Nome_Aba(ExcelSubstitui_NomeInvalido(Mid(objPlanilha.sNomePlanilha, 1, EXCEL_STRING_NOME_PLANILHA)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    End If
    
'    'Esse trecho serve apenas para garantir que o Excel será redesenhado,
'    'e só é utilizado pq mnão foi encontrada uma solução melhor
'    Set objGrafico = objPastaTrabalho.Charts.Add
'    objGrafico.Name = "Lixo"
'    objGrafico.ChartWizard
'    objGrafico.Delete
'    '********************************************************************

    lErro = objExcelApp.Ajustar_Tam_Colunas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se é para apenas imprimir a planilha
    If objPlanilha.iModoImpressao = EXCEL_MODO_IMPRESSÃO_GRAFICO Then

'        'Indica que a planilha será impressa no estilo Landscape
'        objPlanilhaExcel.PageSetup.Orientation = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE
'
'        'Indica que a planilha será impressa com linhas de grade
'        objPlanilhaExcel.PageSetup.PrintGridlines = EXCEL_PLANILHA_IMPRIME_GRADE
'
'        'Imprime o gráfico e fecha o excel
'        objPlanilhaExcel.PrintOut

        lErro = objExcelApp.Imprimir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'        'Fecha o Excel
'        Call Excel_Fechar

        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
        
        '#################################################
        'Inserido por Wagner 01/03/2007
        If objPlanilha.iModoGravacao = DESMARCADO Then
'            'Torna o Excel Visível
'            GL_objExcelSist.Visible = True

            lErro = objExcelApp.Mudar_Visibilidade(True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            'CRIAR CóDIGO PARA GRAVAÇÂO
        
'            objPastaTrabalho.SaveAS objPlanilha.sNomeDiretorioArq & ".xls"
'            objPastaTrabalho.Close

            lErro = objExcelApp.Salvar(objPlanilha.sNomeDiretorioArq)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
'            'Fecha o Excel
'            Call Excel_Fechar

            lErro = objExcelApp.Fechar()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If
        '#################################################
        
    End If
    
    Excel_Gera_Planilha = SUCESSO
    
    Exit Function
    
Erro_Excel_Gera_Planilha:

    Excel_Gera_Planilha = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 102106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142313)

    End Select

'    'Fecha o Excel
'    Call Excel_Fechar
    Call objExcelApp.Fechar

    Exit Function

End Function

Public Function Excel_Abrir() As Long
'Cria uma instÂncia do Excel para uso pelo Corporator

On Error GoTo Erro_Excel_Abrir

    'Se não existe uma instância do excel em aberto
    If GL_objExcelSist Is Nothing Then
        
        'Cria uma nova instância do Excel
        Set GL_objExcelSist = CreateObject("Excel.Application")
    
    End If
    
    Excel_Abrir = SUCESSO
    
    Exit Function
    
Erro_Excel_Abrir:
    
    Excel_Abrir = gErr
    
    Select Case gErr

        Case 429
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_NAO_INSTALADO", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142314)
        
    End Select
    
    Call Excel_Fechar

End Function

Public Function Excel_Fechar() As Long
'Fecha o Excel e libera o obj global que matném a referência ao Excel ativa para o sistema

    'Se existe uma instância do excel em aberto
    If Not GL_objExcelSist Is Nothing Then
        
        'Fecha o Excel
        GL_objExcelSist.Quit
        
        'Libera a memória alocada para referenciar o Excel
        Set GL_objExcelSist = Nothing
    
    End If

End Function

Public Function Excel_Obtem_Largura_Coluna(ByVal dLarguraTwips As Double, dLarguraExcel As Double) As Long
'Converte a largura de um controle em twips para a largura no padrão do excel

Dim lErro As Long
Dim lPoints As Long
Dim lTotalPixels As Long
Dim dLarguraPixels As Double
Dim hdc As Long
Dim tTextMetric As TEXTMETRIC

On Error GoTo Erro_Excel_Obtem_Largura_Coluna

    hdc = GetDC(0)
    
    'Obtém o total de pixels exibidos de acordo com a resolução utilizada para o monitor
    lTotalPixels = GetDeviceCaps(hdc, LOGPIXELSX)
    
    'Converte a largura passada para pixels
    dLarguraPixels = dLarguraTwips / (1440 / lTotalPixels)
    
    'Obtém a largura média em pixels
    Call GetTextMetrics(hdc, tTextMetric)
    
    Call ReleaseDC(0, hdc)
    
    'Calcula a largura no padrão do Excel
    dLarguraExcel = dLarguraPixels / tTextMetric.tmAveCharWidth
    
    Excel_Obtem_Largura_Coluna = SUCESSO
    
    Exit Function
    
Erro_Excel_Obtem_Largura_Coluna:

    Excel_Obtem_Largura_Coluna = gErr
    
    Select Case gErr

        Case 102107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142315)

        End Select

    Exit Function

End Function

Private Function ExcelSubstitui_NomeInvalido(sNome As String) As String
'o caracter "/" estava dando erro, por isso essa funcao foi criada por tulio
'snome eh parametro de input e traz o nome q pode ter a barra
'retorna uma string sem a barra

Dim sStr As String

    'se encontrou o caracter / na string
    If InStr(1, sNome, "/") > 0 Then
        
        'troca o "/" por "_"
        sStr = Replace(sNome, "/", "_")
            
    'senao
    Else
        
        'recebe o proprio snome, q em teoria, eh valido
        sStr = sNome
        
    End If

    ExcelSubstitui_NomeInvalido = sStr
        
End Function

Public Function Excel_Gera_Planilha_Fluxo(ByVal objPlanilha As ClassPlanilhaExcel) As Long

Dim lErro As Long
Dim iLinha As Long
Dim iColuna As Integer
Dim objColuna As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
'Dim avValoresCelulas() As Variant

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
'Dim objGrafico As Object
Dim iLin As Long
Dim iCol As Integer
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Gera_Planilha_Fluxo

'    'Abre o excel
'    lErro = Excel_Abrir()
'    If lErro <> SUCESSO Then gError 199214

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    'Cria uma nova pasta de trabalho
'    Set objPastaTrabalho = GL_objExcelSist.Workbooks.Add
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
'
'    'Redimensiona o array de acordo a quantidade de linhas e colunas recebidas
'    ReDim avValoresCelulas(1 To objPlanilha.colColunas(1).colCelulas.Count, 1 To objPlanilha.colColunas.Count)
'
    lErro = objExcelApp.Criar_Planilha()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Para cada coluna que deverá receber dados
    For iColuna = 1 To objPlanilha.colColunas.Count
        
        'Instancia objColuna como um obj da Coleção de colunas
        Set objColuna = objPlanilha.colColunas(iColuna)
        
        'E para cada linha que receberá dados
        For iLinha = 1 To objColuna.colCelulas.Count
            
            'Instancia objCelulas como um obj da Coleção de células
            Set objCelulas = objColuna.colCelulas(iLinha)
            
'            'Monta o array que será usado para preencher a planilha
'            avValoresCelulas(iLinha, iColuna) = objCelulas.vValor
            
            lErro = objExcelApp.Preencher_Celula(iLinha, iColuna, objCelulas.vValor, objCelulas.sNumberFormat, objCelulas.bFonteNegrito, objCelulas.sFonte, objCelulas.dFonteTamanho, objCelulas.vAlinhamentoHorizontal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
'
'        'Monta a Faixa de Células da coluna atual
'        Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(1, iColuna), objPlanilhaExcel.Cells(objColuna.colCelulas.Count, iColuna))
'
'        'Altera a largura da coluna conforme a largura passada
'        objFaixaCelulasPreencher.ColumnWidth = objColuna.dLarguraColuna
    
    Next

'    'Monta a Faixa das Células que receberão os valores passado
'    Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(1, 1), objPlanilhaExcel.Cells(objPlanilha.colColunas(1).colCelulas.Count, objPlanilha.colColunas.Count))
'
'    'Passa os valores armazenados no array para a faixa determinada acima
'    objFaixaCelulasPreencher.Value = avValoresCelulas
'
'    '########################################################
'    'Inserido por Wagner 01/03/2007
'    For iCol = 1 To objPlanilha.colColunas.Count
'
'        For iLin = 1 To objPlanilha.colColunas(iCol).colCelulas.Count
'
'            Set objCelulas = objPlanilha.colColunas(iCol).colCelulas(iLin)
'
'            If objCelulas.iCelulaCor = 1 Then
'                objFaixaCelulasPreencher(iLin, iCol).Interior.Color = RGB(192, 192, 192)
'            End If
'
'            If objCelulas.sFonte <> "" Then
'                objFaixaCelulasPreencher(iLin, iCol).Font.Name = objCelulas.sFonte
'                objFaixaCelulasPreencher(iLin, iCol).Font.Size = objCelulas.dFonteTamanho
'            End If
'            If Not (objCelulas.vAlinhamentoHorizontal = Empty) Then
'                objFaixaCelulasPreencher(iLin, iCol).HorizontalAlignment = objCelulas.vAlinhamentoHorizontal
'            End If
'
'            If objPlanilha.colColunas(iCol).colCelulas.Item(iLin).bFonteNegrito Then
'                objFaixaCelulasPreencher(iLin, iCol).Font.Bold = True
'            End If
'
'            If Len(Trim(objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat)) > 0 Then
'                objFaixaCelulasPreencher(iLin, iCol).NumberFormat = objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat
'            End If
'        Next
'
'    Next
    '########################################################

    'Se foi passado um nome para a planilha
    If Len(Trim(objPlanilha.sNomePlanilha)) > 0 Then
    
'        'Exibe o nome da planilha no excel
'        objPlanilhaExcel.Name = ExcelSubstitui_NomeInvalido(Mid(objPlanilha.sNomePlanilha, 1, EXCEL_STRING_NOME_PLANILHA))
    
        lErro = objExcelApp.Muda_Nome_Aba(ExcelSubstitui_NomeInvalido(Mid(objPlanilha.sNomePlanilha, 1, EXCEL_STRING_NOME_PLANILHA)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
'    'Esse trecho serve apenas para garantir que o Excel será redesenhado,
'    'e só é utilizado pq mnão foi encontrada uma solução melhor
'    Set objGrafico = objPastaTrabalho.Charts.Add
'    objGrafico.Name = "Lixo"
'    objGrafico.ChartWizard
'    objGrafico.Delete
'    '********************************************************************
    
    lErro = objExcelApp.Ajustar_Tam_Colunas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se é para apenas imprimir a planilha
    If objPlanilha.iModoImpressao = EXCEL_MODO_IMPRESSÃO_GRAFICO Then

'        'Indica que a planilha será impressa no estilo Landscape
'        objPlanilhaExcel.PageSetup.Orientation = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE
'
'        'Indica que a planilha será impressa com linhas de grade
'        objPlanilhaExcel.PageSetup.PrintGridlines = EXCEL_PLANILHA_IMPRIME_GRADE
'
'        'Imprime o gráfico e fecha o excel
'        objPlanilhaExcel.PrintOut
        
        lErro = objExcelApp.Imprimir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        'Fecha o Excel
'        Call Excel_Fechar

        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
        
        '#################################################
        'Inserido por Wagner 01/03/2007
        If objPlanilha.iModoGravacao = DESMARCADO Then
'            'Torna o Excel Visível
'            GL_objExcelSist.Visible = True

            lErro = objExcelApp.Mudar_Visibilidade(True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Else
            'CRIAR CóDIGO PARA GRAVAÇÂO
        
'            objPastaTrabalho.SaveAS objPlanilha.sNomeDiretorioArq & ".xls"
'            objPastaTrabalho.Close
            
            lErro = objExcelApp.Salvar(objPlanilha.sNomeDiretorioArq)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
'            'Fecha o Excel
'            Call Excel_Fechar

            lErro = objExcelApp.Fechar()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If
        '#################################################
        
    End If
    
    Excel_Gera_Planilha_Fluxo = SUCESSO
    
    Exit Function
    
Erro_Excel_Gera_Planilha_Fluxo:

    Excel_Gera_Planilha_Fluxo = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 199214
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199215)

    End Select

'    'Fecha o Excel
'    Call Excel_Fechar
    Call objExcelApp.Fechar

    Exit Function

End Function

Public Function Browse_Executa_SQL_Excel(ByVal objBrowse As AdmBrowse, ByVal lNumRegistros As Long, ByVal iModoImpressao As Integer) As Long
'Executa o comando SQL guarda em objBrowse.sSQL
'objBrowse RECEBE(Input) as informações do browse que serão necessárias para execução do select
'objBrowse RETORNA(Output) os registros retornados pela leitura dentro da coleção colRegistros

Dim lComando As Long
Dim lErro As Long
Dim avCampo(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Dim avCampoOrd(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Dim iLinha As Long
Dim iColuna As Long
Dim iFiltro As Long
Dim iValor As Long
Dim dLarguraColuna As Double
Dim iLinhaInicial As Long
Dim iIndice As Long

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objPlanilhaGrafico As Object
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim objBrowseExcelAux As AdmBrowseExcelAux
Dim objBrowseUsuCampo As AdmBrowseUsuarioCampo
Dim colBrowseUsuarioCampo As New Collection
Dim iIndiceColdeFato As Long
Dim colIndice As New Collection
Dim colFormato As New Collection
Dim vValor As Variant
Dim sFormato As String
Dim iIndiceColAux As Long
Dim iIndiceLin As Long
Dim iExercicio As Integer

Dim sConta As String
Dim sContaMascarada As String
Dim sCcl As String
Dim sCclMascarado As String
Dim sProduto As String
Dim sProdutoMascarado As String
Dim sCgc As String
Dim sDocOrigem As String
Dim iExercicioAux As Integer
Dim iPeriodo As Integer
Dim sNomeExterno As String
Dim objFrmAguarde As ClassFrmAguarde
Dim objFrmAguardeTela As FrmAguarde
Dim bArqAberto As Boolean
Dim sNomeArq As String
Dim sRegistro As String
Dim bAchou As Boolean
Dim sValorAux As String
Dim colCampos As New Collection
Dim objCampo As AdmCampos, sAspas As String
Dim bModal As Boolean
Dim objExcelApp As New ClassExcelApp
Dim sErroAux As String, sPontoErro As String

On Error GoTo Erro_Browse_Executa_SQL_Excel

    sAspas = """"

'sPontoErro = "01"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202179
    
'sPontoErro = sPontoErro & "," & "02"
    
    bModal = False
    If Not (objBrowse.objPrincMDIChild Is Nothing) Then
        If UCase(TypeName(objBrowse.objPrincMDIChild)) = "PRINCPOPUP" Then
            bModal = True
        End If
    End If
    If Not bModal Then
        Set objFrmAguarde = New ClassFrmAguarde
        Set objFrmAguardeTela = New FrmAguarde
    End If
    
'sPontoErro = sPontoErro & "," & "03"
    
    'Prepara as variáveis que serão usadas no select para receber os valores lidos
    lErro = objBrowse.Browse_Reinicializa_avCampo2(objBrowse, avCampo)
    If lErro <> SUCESSO Then gError 202180

'sPontoErro = sPontoErro & "," & "04"

    If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
    
'sPontoErro = sPontoErro & "," & "05"

        If Len(Trim(objBrowse.objBrowseExcel.sArquivo)) = 0 Then
            sNomeArq = objBrowse.objForm.Name & "_" & gsUsuario & "_" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & ".txt"
        Else
            sNomeArq = objBrowse.objBrowseExcel.sArquivo & ".txt"
        End If
    
'sPontoErro = sPontoErro & "," & "06"
        
        Open objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq For Output As #1
        bArqAberto = True

'sPontoErro = sPontoErro & "," & "07"

        If Len(Trim(objBrowse.objBrowseExcel.sTitulo)) > 0 Then
            Print #1, objBrowse.objBrowseExcel.sTitulo
            iLinhaInicial = 2
        Else
            iLinhaInicial = 1
        End If
    
'sPontoErro = sPontoErro & "," & "08"

    Else
    
'        'Abre o excel
'        lErro = Excel_Abrir()
'        If lErro <> SUCESSO Then gError 202181

'sPontoErro = sPontoErro & "," & "09"

        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'sPontoErro = sPontoErro & "," & "10"

'        'Desabilita as mensagens do Excel
'        GL_objExcelSist.DisplayAlerts = False
'        'GL_objExcelSist.ScreenUpdating = False
'
'        'Cria uma nova pasta de trabalho
'        Set objPastaTrabalho = GL_objExcelSist.Workbooks.Add
'
'        For iIndice = objPastaTrabalho.Sheets.Count To 2 Step -1
'            objPastaTrabalho.Sheets.Item(iIndice).Delete
'        Next
'
'        'Seleciona a planilha ativa na pasta de trabalho criada
'        Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
'        objPlanilhaExcel.Name = "Dados"

        lErro = objExcelApp.Criar_Planilha()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'sPontoErro = sPontoErro & "," & "11"

        lErro = objExcelApp.Muda_Nome_Aba("Dados")
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'sPontoErro = sPontoErro & "," & "12"

        lErro = objExcelApp.Apagar_Abas()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'sPontoErro = sPontoErro & "," & "13"

        If Len(Trim(objBrowse.objBrowseExcel.sTitulo)) > 0 Then
            
'sPontoErro = sPontoErro & "," & "14"

            iLinhaInicial = 2
'            Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(1, 1), objPlanilhaExcel.Cells(1, objBrowse.objGrid.Cols))
'            Call objFaixaCelulasPreencher.Merge
'            objFaixaCelulasPreencher.Value = objBrowse.objBrowseExcel.sTitulo
'            objFaixaCelulasPreencher.Font.Bold = True
'            objFaixaCelulasPreencher.Font.Size = 24

            lErro = objExcelApp.Fazer_Merge(1, 1, 1, objBrowse.objGrid.Cols)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'sPontoErro = sPontoErro & "," & "15"
            
            lErro = objExcelApp.Preencher_Celula(1, 1, objBrowse.objBrowseExcel.sTitulo, "", True, "", 24)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'sPontoErro = sPontoErro & "," & "16"

        Else
            
'sPontoErro = sPontoErro & "," & "17"
            
            iLinhaInicial = 1
        End If
        
    End If
    
'sPontoErro = sPontoErro & "," & "18"
    
    sRegistro = ""
    
    'Para cada coluna que deverá receber dados
    For iColuna = 1 To objBrowse.objGrid.Cols
    
'sPontoErro = sPontoErro & "," & "19"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela & " Coluna: " & CStr(iColuna) & "-" & objBrowse.objGrid.TextMatrix(0, iColuna - 1)

        iIndice = 0
        For Each objBrowseUsuCampo In objBrowse.colBrowseUsuarioCampo
            iIndice = iIndice + 1
            If objBrowseUsuCampo.iPosicaoTela = iColuna Then
                colBrowseUsuarioCampo.Add objBrowseUsuCampo
                iIndiceColdeFato = iIndice
                colIndice.Add iIndiceColdeFato
                Exit For
            End If
        Next
        
'sPontoErro = sPontoErro & "," & "20"

        If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
            
'sPontoErro = sPontoErro & "," & "21"
            
            sRegistro = sRegistro & IIf(sRegistro <> "", ";", "") & Replace(objBrowse.objGrid.TextMatrix(0, iColuna - 1), sAspas, sAspas & sAspas)
        
'sPontoErro = sPontoErro & "," & "22"
        
        Else
                
'            'Obtém a largura da coluna no padrão do Excel
'            lErro = CF("Excel_Obtem_Largura_Coluna", objBrowse.objGrid.ColWidth(iColuna - 1), dLarguraColuna)
'            If lErro <> SUCESSO Then gError 202182
'
'            'Monta a Faixa de Células da coluna atual
'            Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(iLinhaInicial, iColuna), objPlanilhaExcel.Cells(iLinhaInicial, iColuna))
'            objFaixaCelulasPreencher.Value = objBrowse.objGrid.TextMatrix(0, iColuna - 1)
'            objFaixaCelulasPreencher.Font.Bold = True
'
'            'Altera a largura da coluna conforme a largura passada
'            objFaixaCelulasPreencher.ColumnWidth = dLarguraColuna

'sPontoErro = sPontoErro & "," & "23"

            lErro = objExcelApp.Preencher_Celula(iLinhaInicial, iColuna, objBrowse.objGrid.TextMatrix(0, iColuna - 1), "", True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
'sPontoErro = sPontoErro & "," & "24"
     
        End If
        
    Next
    
'sPontoErro = sPontoErro & "," & "25"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela
    
    If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
        Print #1, Replace(Replace(sRegistro, Chr(10), ""), Chr(13), "")
    End If
    
'sPontoErro = sPontoErro & "," & "26"

    Call Browse_Executa_SQL_Excel1(objBrowse, colIndice, objBrowse.colBrowseUsuarioCampo, colCampos)
    
'sPontoErro = sPontoErro & "," & "27"

    'Executa o select
    lErro = CF("Browse_Prepara_Comando_SQL", objBrowse.sSQL, objBrowse.colValorCampo.Count, lComando, avCampo, objBrowse.colSelecao, BROWSE_EXECUTA_BIND)
    If lErro <> SUCESSO Then gError 202183
    
'sPontoErro = sPontoErro & "," & "28"

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202184
    
'sPontoErro = sPontoErro & "," & "29"
    
    'Se não encontrou=> erro
    'If lErro = AD_SQL_SEM_DADOS Then gError 102092
    
    'cria um array com o tamanho de um registro visível no browser
    ReDim avRegistro(1 To objBrowse.colBrowseUsuarioCampo.Count)
    
'sPontoErro = sPontoErro & "," & "30"

    iLinha = iLinhaInicial
    
    If Not bModal Then
        objFrmAguarde.iTotalItens = lNumRegistros
        Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    End If
    
'sPontoErro = sPontoErro & "," & "31"

    'Enquanto houver registros
    Do While lErro = AD_SQL_SUCESSO
    
'sPontoErro = sPontoErro & "," & "32"

        Call Browse_Executa_SQL_Excel2(objBrowse, avCampo, avCampoOrd, colIndice, objBrowse.colBrowseUsuarioCampo, iExercicio)
    
'sPontoErro = sPontoErro & "," & "33"

        iLinha = iLinha + 1
        sRegistro = ""
    
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela & " Linha: " & CStr(iLinha)
    
        'Para cada campo acessível ao usuário no browse, ou seja, cada campo que está sendo lido no BD
        For iColuna = 1 To colBrowseUsuarioCampo.Count
        
'sPontoErro = sPontoErro & "," & "34"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela & " Linha: " & CStr(iLinha) & " Coluna: " & CStr(iColuna)

            Set objCampo = colCampos.Item(iColuna)
            vValor = avCampoOrd(iColuna)
            sFormato = ""
        
'sPontoErro = sPontoErro & "," & "35"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela & " Linha: " & CStr(iLinha) & " Coluna: " & CStr(iColuna) & " Campo: " & colBrowseUsuarioCampo(iColuna).sNome & " Valor: " & CStr(vValor)
            
            Select Case objCampo.iTipo

                Case ADM_TIPO_DATE

'sPontoErro = sPontoErro & "," & "36"

                    If vValor < #1/1/1900# Then
                        vValor = ""
                    Else
                        vValor = StrParaDate(CStr(vValor))
                    End If
                    If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
                        sFormato = "yyyy-mm-dd"
                    Else
                        sFormato = "dd/mm/yyyy"
                    End If

'sPontoErro = sPontoErro & "," & "37"

                Case ADM_TIPO_DOUBLE

'sPontoErro = sPontoErro & "," & "38"

                    vValor = StrParaDbl(CStr(vValor))

                    Select Case objCampo.iSubTipo

                        Case ADM_SUBTIPO_PERCENTUAL
                            sFormato = "0.00%"

                        Case ADM_SUBTIPO_HORA
                            vValor = CDate(vValor)
                            sFormato = "hh:mm:ss"

                        Case Else
                            sFormato = "#,##0.00##"

                    End Select

'sPontoErro = sPontoErro & "," & "39"

                Case ADM_TIPO_SMALLINT

'sPontoErro = sPontoErro & "," & "40"
                    Select Case objCampo.iSubTipo

                        Case ADM_SUBTIPO_EXERCICIO

                            iExercicioAux = vValor
                            sNomeExterno = ""

'sPontoErro = sPontoErro & "," & "41"

                            'le o exercicio do BD e exibe o nome externo
                            lErro = CF("Exercicio_Le_NomeExterno", objBrowse.lComando2, iExercicioAux, sNomeExterno)
                            If lErro <> SUCESSO Then gError 202173

                            vValor = sNomeExterno

'sPontoErro = sPontoErro & "," & "42"
                        Case ADM_SUBTIPO_PERIODO

                            If iExercicio <> 0 Then

                                iPeriodo = vValor
                                sNomeExterno = ""
'sPontoErro = sPontoErro & "," & "43"
                                'le o periodo do BD e exibe o nome externo
                                lErro = CF("Periodo_Le_NomeExterno", objBrowse.lComando2, iExercicio, iPeriodo, sNomeExterno)
                                If lErro <> SUCESSO Then
                                    gError 202174
                                End If

                                vValor = sNomeExterno
'sPontoErro = sPontoErro & "," & "44"
                            End If

                        Case ADM_SUBTIPO_STATUS
'sPontoErro = sPontoErro & "," & "45"
                            sValorAux = ""
                            Call CF("Converte_Status", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "46"
                        Case ADM_SUBTIPO_TIPOCONTACCL
'sPontoErro = sPontoErro & "," & "47a"
                            sValorAux = ""
                            Call CF("Converte_TipoContaCcl", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "47b"
                        Case ADM_SUBTIPO_NATUREZACONTA
'sPontoErro = sPontoErro & "," & "48a"
                            sValorAux = ""
                            Call CF("Converte_NaturezaConta", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "48b"
                        Case ADM_SUBTIPO_TIPORATEIO
'sPontoErro = sPontoErro & "," & "49a"
                            sValorAux = ""
                            Call CF("Converte_TipoRateio", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "49b"
                        Case ADM_SUBTIPO_AGLUTINA
'sPontoErro = sPontoErro & "," & "50a"
                            sValorAux = ""
                            Call CF("Converte_Aglutina", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "50b"
                        Case ADM_SUBTIPO_NATUREZAPRODUTO
'sPontoErro = sPontoErro & "," & "51a"
                            sValorAux = ""
                            Call CF("Converte_NaturezaProduto", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "51b"
                        Case ADM_SUBTIPO_KIT
'sPontoErro = sPontoErro & "," & "52a"
                            sValorAux = ""
                            Call CF("Converte_Situacao_Kit", vValor, sValorAux)
                            vValor = sValorAux
'sPontoErro = sPontoErro & "," & "52b"

                        Case Else
'sPontoErro = sPontoErro & "," & "53"
                            sFormato = "#,##0"

                    End Select

            Case ADM_TIPO_INTEGER

'sPontoErro = sPontoErro & "," & "54"

                Select Case objCampo.iSubTipo

                    Case ADM_SUBTIPO_TIPO_DOC_PROJETO

                        sDocOrigem = ""

'sPontoErro = sPontoErro & "," & "55"
                        lErro = CF("Projeto_Le_DocInfo", vValor, sDocOrigem)
                        If lErro <> SUCESSO Then gError 202175

                        vValor = sDocOrigem
'sPontoErro = sPontoErro & "," & "56"
                    Case Else
                            sFormato = "##0"

                End Select

            Case Else

'sPontoErro = sPontoErro & "," & "57"

                Select Case objCampo.iSubTipo

                    Case ADM_SUBTIPO_CONTA

'sPontoErro = sPontoErro & "," & "58"

                        sContaMascarada = ""
                        If Len(vValor) > 0 Then
                            sConta = vValor

                            lErro = CF("Mascara_MascararConta", sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 202176
                        End If


                        vValor = sContaMascarada

'sPontoErro = sPontoErro & "," & "59"

                    Case ADM_SUBTIPO_CCL

'sPontoErro = sPontoErro & "," & "60"

                        sCclMascarado = ""
                        If Len(vValor) > 0 Then

                            sCcl = vValor

                            lErro = CF("Mascara_MascararCcl", sCcl, sCclMascarado)
                            If lErro <> SUCESSO Then gError 202177
                        End If

                        vValor = sCclMascarado

'sPontoErro = sPontoErro & "," & "61"

                    Case ADM_SUBTIPO_PRODUTO

'sPontoErro = sPontoErro & "," & "62"
                        sProdutoMascarado = ""
                        If Len(vValor) > 0 Then

                            sProduto = vValor

                            lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
                            If lErro <> SUCESSO Then gError 202178

                        End If

                        vValor = sProdutoMascarado

'sPontoErro = sPontoErro & "," & "63"

                    Case ADM_SUBTIPO_CGC

'sPontoErro = sPontoErro & "," & "64"
                        sCgc = ""
                        Select Case Len(Trim(vValor))

                            Case STRING_CPF 'CPF
                                sCgc = Format(vValor, "000\.000\.000-00; ; ; ")

                            Case STRING_CGC 'CGC
                                sCgc = Format(vValor, "00\.000\.000\/0000-00; ; ; ")

                        End Select

                        vValor = sCgc

'sPontoErro = sPontoErro & "," & "65"

                End Select

'sPontoErro = sPontoErro & "," & "66"

                If IsNumeric(vValor) Then
                    If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
                        vValor = vValor & " "
                    Else
                        vValor = "'" & vValor
                    End If
                End If

'sPontoErro = sPontoErro & "," & "67"

            End Select
            
'sPontoErro = sPontoErro & "," & "68"
            
            'Monta a Faixa de Células da linha atual
            If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
                
'sPontoErro = sPontoErro & "," & "69"
                If sFormato <> "" Then vValor = Format(vValor, sFormato)
                sRegistro = sRegistro & IIf(iColuna <> 1, ";", "") & """" & Replace(vValor, sAspas, sAspas & sAspas) & """"
'sPontoErro = sPontoErro & "," & "70"
            Else
'                Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(iLinha, iColuna), objPlanilhaExcel.Cells(iLinha, iColuna))
'                objFaixaCelulasPreencher.Value = vValor
'                If sFormato <> "" Then objFaixaCelulasPreencher.NumberFormat = sFormato
'sPontoErro = sPontoErro & "," & "71"
                lErro = objExcelApp.Preencher_Celula(iLinha, iColuna, vValor, sFormato)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'sPontoErro = sPontoErro & "," & "72"
            End If
'sPontoErro = sPontoErro & "," & "73"
'            If iLinha = iLinhaInicial + 1 Then
'                colFormato.Add sFormato
'            End If
'sPontoErro = sPontoErro & "," & "74"
        Next

'sPontoErro = sPontoErro & "," & "75"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela & " Linha: " & CStr(iLinha)
        If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
            Print #1, Replace(Replace(sRegistro, Chr(10), ""), Chr(13), "")
        End If
'sPontoErro = sPontoErro & "," & "76"
        If Not bModal Then
            Call objFrmAguardeTela.ProcessouItem
            
            If objFrmAguarde.iCancelar = MARCADO Then Exit Do
        End If
'sPontoErro = sPontoErro & "," & "77"
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202185
'sPontoErro = sPontoErro & "," & "78"
    Loop
'sPontoErro = sPontoErro & "," & "79"
'sErroAux = "Browse: " & objBrowse.objBrowseArquivo.sNomeTela

    If Not bModal Then
        Set objFrmAguardeTela = Nothing
        If objFrmAguarde.iCancelar = MARCADO Then gError 202186
    End If
'sPontoErro = sPontoErro & "," & "80"
    
    Call Comando_Fechar(lComando)
        
'sPontoErro = sPontoErro & "," & "81"
    If objBrowse.objBrowseExcel.iFormato = EXCEL_FORMATO_CSV Or objBrowse.objBrowseExcel.sLocalizacaoCSV <> "" Then
        
'sPontoErro = sPontoErro & "," & "82"
        Close #1
        bArqAberto = False
        
'        'Abre o excel
'        lErro = Excel_Abrir()
'        If lErro <> SUCESSO Then gError 202181
       
'sPontoErro = sPontoErro & "," & "83"
        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
        'Desabilita as mensagens do Excel
        'GL_objExcelSist.DisplayAlerts = False
        'GL_objExcelSist.ScreenUpdating = False
        
        'Cria uma nova pasta de trabalho
        'Set objPastaTrabalho = GL_objExcelSist.WorkBooks.Add
        
'        Call GL_objExcelSist.WorkBooks.Open(objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq, Format:=6, Delimiter:=";")
'        Call GL_objExcelSist.WorkBooks.OpenText(objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq, DataType:=1, TextQualifier:=-4142, Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:=".")
        
        'GL_objExcelSist.ScreenUpdating = False
        
        'Call GL_objExcelSist.Workbooks.OpenText(objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq, Origin:=2, DataType:=1, TextQualifier:=1, Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:=".")
               
        'Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
                
        'Seleciona a planilha ativa na pasta de trabalho criada
        'Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
        'objPlanilhaExcel.Name = "Dados"
        
'sPontoErro = sPontoErro & "," & "84"
        
        'Abre o excel
        lErro = objExcelApp.Abrir_Planilha_CSV(objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'sPontoErro = sPontoErro & "," & "85"

        lErro = objExcelApp.Muda_Nome_Aba("Dados")
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'sPontoErro = sPontoErro & "," & "86"
'        'GL_objExcelSist.Visible = True
'
'        Call GL_objExcelSist.Columns("A:A").Select
'        Call GL_objExcelSist.Selection.TextToColumns(Destination:=GL_objExcelSist.Range("A1"), DataType:=1, TextQualifier:=1, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=True, Comma:=False, Space:=False, Other:=False)
'
'        'GL_objExcelSist.ScreenUpdating = True
'
'        GL_objExcelSist.Cells.Select
'        GL_objExcelSist.Cells.EntireColumn.AutoFit

        lErro = objExcelApp.Ajustar_Tam_Colunas()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'sPontoErro = sPontoErro & "," & "87"
'        GL_objExcelSist.Rows("1:1").Select
'        GL_objExcelSist.Selection.Font.Bold = True
'        If Len(Trim(objBrowse.objBrowseExcel.sTitulo)) > 0 Then
'            GL_objExcelSist.Selection.Font.Size = 24
'            GL_objExcelSist.Rows("2:2").Select
'            GL_objExcelSist.Selection.Font.Bold = True
'        End If
'
'        For iIndice = objPastaTrabalho.Sheets.Count To 2 Step -1
'            objPastaTrabalho.Sheets.Item(iIndice).Delete
'        Next

'sPontoErro = sPontoErro & "," & "88"
        lErro = objExcelApp.Apagar_Abas()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'sPontoErro = sPontoErro & "," & "89"
        If Len(Trim(objBrowse.objBrowseExcel.sTitulo)) > 0 Then
            
'sPontoErro = sPontoErro & "," & "90"

            lErro = objExcelApp.Fazer_Merge(1, 1, 1, objBrowse.objGrid.Cols)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'sPontoErro = sPontoErro & "," & "91"
            lErro = objExcelApp.Preencher_Celula(1, 1, objBrowse.objBrowseExcel.sTitulo, "", True, "", 24)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
'sPontoErro = sPontoErro & "," & "92"
        End If
        
        If objBrowse.objBrowseExcel.iFormato <> EXCEL_FORMATO_CSV Then
        
'sPontoErro = sPontoErro & "," & "93"
            lErro = objExcelApp.Salvar(objBrowse.objBrowseExcel.sLocalizacaoCSV & Replace(sNomeArq, ".txt", ""))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'sPontoErro = sPontoErro & "," & "94"
            'objPastaTrabalho.SaveAS objBrowse.objBrowseExcel.sLocalizacaoCSV & Replace(sNomeArq, ".txt", ".xls"), FileFormat:=-4143, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            
            Kill objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq
'sPontoErro = sPontoErro & "," & "95"
        End If
        
    Else
'sPontoErro = sPontoErro & "," & "96"
        lErro = objExcelApp.Ajustar_Tam_Colunas()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'sPontoErro = sPontoErro & "," & "97"
    End If
    
    
    'FAZER POSTERIORMENTE COMPATÍVEL COM OPENOFFICE
    If objBrowse.objBrowseExcel.iFormato <> EXCEL_FORMATO_CSV Then

'sPontoErro = sPontoErro & "," & "98"
        
        If objBrowse.objBrowseExcel.colFormulas.Count > 0 Then
            'Para cada coluna que deverá receber dados
            For iColuna = 1 To objBrowse.objGrid.Cols
                Set objBrowseUsuCampo = colBrowseUsuarioCampo.Item(iColuna)
                For Each objBrowseExcelAux In objBrowse.objBrowseExcel.colFormulas
                    If objBrowseExcelAux.sCampo = objBrowseUsuCampo.sNome Then
                        'Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(lNumRegistros + iLinhaInicial + 1, iColuna), objPlanilhaExcel.Cells(lNumRegistros + iLinhaInicial + 1, iColuna))
'                        Select Case objBrowseExcelAux.iFormula
'                            Case EXCEL_FORMULA_SUM
'                                objFaixaCelulasPreencher.FormulaR1C1 = "=SUM(R" & CStr(iLinhaInicial + 1) & "C" & CStr(iColuna) & ":R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(iColuna) & ")"
'                            Case EXCEL_FORMULA_AVG
'                                objFaixaCelulasPreencher.FormulaR1C1 = "=AVERAGE(R" & CStr(iLinhaInicial + 1) & "C" & CStr(iColuna) & ":R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(iColuna) & ")"
'                            Case EXCEL_FORMULA_MAX
'                                objFaixaCelulasPreencher.FormulaR1C1 = "=MAX(R" & CStr(iLinhaInicial + 1) & "C" & CStr(iColuna) & ":R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(iColuna) & ")"
'                            Case EXCEL_FORMULA_MIN
'                                objFaixaCelulasPreencher.FormulaR1C1 = "=MIN(R" & CStr(iLinhaInicial + 1) & "C" & CStr(iColuna) & ":R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(iColuna) & ")"
'                            Case EXCEL_FORMULA_COUNT
'                                objFaixaCelulasPreencher.FormulaR1C1 = "=COUNTA(R" & CStr(iLinhaInicial + 1) & "C" & CStr(iColuna) & ":R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(iColuna) & ")"
'                        End Select
'                        objFaixaCelulasPreencher.Font.Bold = True
                        sFormato = colFormato.Item(iColuna)
                        'If sFormato <> "" Then objFaixaCelulasPreencher.NumberFormat = sFormato
                        
'sPontoErro = sPontoErro & "," & "99"
                        lErro = objExcelApp.Preencher_Celula_Formula(iLinhaInicial, iLinhaInicial + lNumRegistros, iColuna, objBrowseExcelAux.iFormula, sFormato)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
'sPontoErro = sPontoErro & "," & "100"
                        Exit For
                    End If
                Next
            Next
        End If

        If objBrowse.objBrowseExcel.iTabelaDinamica = MARCADO Then

'sPontoErro = sPontoErro & "," & "101"
            bAchou = False
'            For Each objBrowseExcelAux In objBrowse.objBrowseExcel.colCampos
'                For iColuna = 1 To objBrowse.objGrid.Cols
'                    Set objBrowseUsuCampo = colBrowseUsuarioCampo.Item(iColuna)
'                    If objBrowseExcelAux.sCampo = objBrowseUsuCampo.sNome Then
'                        Set objBrowseExcelAux.objBrowseUsuarioCampo = objBrowseUsuCampo
'                        'If objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_LINHA Then bAchou = True
'                        bAchou = True
'                        Exit For
'                    End If
'                Next
'            Next

'            If objPastaTrabalho.Sheets.Count <= 1 Then
'                objPastaTrabalho.Sheets.Add After:=objPastaTrabalho.Sheets.Item(1)
'            End If
'            Set objPlanilhaExcel = objPastaTrabalho.Sheets.Item(2)
'            objPlanilhaExcel.Name = "Tabela"
'sPontoErro = sPontoErro & "," & "102"
            lErro = objExcelApp.Criar_Aba()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
'sPontoErro = sPontoErro & "," & "103"
            lErro = objExcelApp.Muda_Nome_Aba("Tabela")
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            '2007
            'Call objPastaTrabalho.PivotCaches.Create(SourceType:=EXCEL_TABDIN_TIPO_DATABASE, SourceData:="Dados!R" & CStr(iLinhaInicial) & "C1:R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(objBrowse.objGrid.Cols)).CreatePivotTable(TableDestination:="Tabela!L1C1", TableName:="TabDin1")

            '2003
            'Call objPastaTrabalho.PivotCaches.Add(SourceType:=EXCEL_TABDIN_TIPO_DATABASE, SourceData:="Dados!R" & CStr(iLinhaInicial) & "C1:R" & CStr(lNumRegistros + iLinhaInicial) & "C" & CStr(objBrowse.objGrid.Cols)).CreatePivotTable(TableDestination:="Tabela!L1C1", TableName:="TabDin1")

'sPontoErro = sPontoErro & "," & "104"
            lErro = objExcelApp.Criar_TabelaDinamica(iLinhaInicial, lNumRegistros + iLinhaInicial, 1, objBrowse.objGrid.Cols)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            iLinha = 0
            iColuna = 0
            iValor = 0
            iFiltro = 0
            
'sPontoErro = sPontoErro & "," & "105"
            For Each objBrowseExcelAux In objBrowse.objBrowseExcel.colCampos

                For iColuna = 1 To objBrowse.objGrid.Cols
                    Set objBrowseUsuCampo = colBrowseUsuarioCampo.Item(iColuna)
                    If objBrowseExcelAux.sCampo = objBrowseUsuCampo.sNome Then
                        Set objBrowseExcelAux.objBrowseUsuarioCampo = objBrowseUsuCampo
                        'If objBrowseExcelAux.iPosicao = EXCEL_TABDIN_POS_LINHA Then bAchou = True
                        bAchou = True
                        Exit For
                    End If
                Next

'sPontoErro = sPontoErro & "," & "106"
                Set objBrowseUsuCampo = objBrowseExcelAux.objBrowseUsuarioCampo

                If bAchou Then
                
'sPontoErro = sPontoErro & "," & "107"
                    lErro = objExcelApp.Incluir_Campo_TabelaDinamica(objBrowseUsuCampo.sTitulo, objBrowseExcelAux.iPosicao, objBrowseExcelAux.iFormula, iLinha, iColuna, iFiltro, iValor)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'sPontoErro = sPontoErro & "," & "108"
'                    Select Case objBrowseExcelAux.iPosicao
'                        Case EXCEL_TABDIN_POS_LINHA
'                            iLinha = iLinha + 1
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Orientation = EXCEL_TABDIN_POS_LINHA
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Position = iLinha
'                        Case EXCEL_TABDIN_POS_COLUNA
'                            iColuna = iColuna + 1
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Orientation = EXCEL_TABDIN_POS_COLUNA
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Position = iColuna
'                        Case EXCEL_TABDIN_POS_FILTRO
'                            iFiltro = iFiltro + 1
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Orientation = EXCEL_TABDIN_POS_FILTRO
'                            objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo).Position = iFiltro
'                        Case EXCEL_TABDIN_POS_VALOR
'                            iValor = iValor + 1
'                            Select Case objBrowseExcelAux.iFormula
'                                Case EXCEL_FORMULA_SUM
'                                    objPlanilhaExcel.PivotTables("TabDin1").AddDataField objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo), "Soma de " & objBrowseUsuCampo.sTitulo, EXCEL_FORMULA_SUM
'                                Case EXCEL_FORMULA_AVG
'                                    objPlanilhaExcel.PivotTables("TabDin1").AddDataField objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo), "Média de " & objBrowseUsuCampo.sTitulo, EXCEL_FORMULA_AVG
'                                Case EXCEL_FORMULA_MIN
'                                    objPlanilhaExcel.PivotTables("TabDin1").AddDataField objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo), "Min. de " & objBrowseUsuCampo.sTitulo, EXCEL_FORMULA_MIN
'                                Case EXCEL_FORMULA_MAX
'                                    objPlanilhaExcel.PivotTables("TabDin1").AddDataField objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo), "Máx. de " & objBrowseUsuCampo.sTitulo, EXCEL_FORMULA_MAX
'                                Case EXCEL_FORMULA_COUNT
'                                    objPlanilhaExcel.PivotTables("TabDin1").AddDataField objPlanilhaExcel.PivotTables("TabDin1").PivotFields(objBrowseUsuCampo.sTitulo), "Qtd de " & objBrowseUsuCampo.sTitulo, EXCEL_FORMULA_COUNT
'                            End Select
'                    End Select

                End If
            Next

            If objBrowse.objBrowseExcel.iIncluirGrafico = MARCADO And (iLinha) <> 0 Then
'sPontoErro = sPontoErro & "," & "109"
'                objPastaTrabalho.Charts.Add
'                Set objPlanilhaGrafico = objPastaTrabalho.Charts.Item(1)
'                objPlanilhaGrafico.SetSourceData Source:=objPastaTrabalho.Sheets("Tabela").PivotTables("TabDin1").rowrange.currentregion
'                objPlanilhaGrafico.ChartType = objBrowse.objBrowseExcel.iTipoGrafico
'                objPlanilhaGrafico.Location 1, "Gráfico"
'                objPastaTrabalho.Sheets("Gráfico").Move After:=objPastaTrabalho.Sheets("Tabela")
'
'                objPastaTrabalho.ShowPivotTableFieldList = True
                'objPastaTrabalho.ShowPivotChartActiveFields = False
                
                lErro = objExcelApp.Incluir_Grafico(objBrowse.objBrowseExcel.iTipoGrafico)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'sPontoErro = sPontoErro & "," & "110"
            End If

        End If

    End If
    'FAZER POSTERIORMENTE COMPATÍVEL COM OPENOFFICE
       
    'GL_objExcelSist.ScreenUpdating = True

    'Se é para apenas imprimir a planilha
    If iModoImpressao = EXCEL_MODO_IMPRESSAO Then

'sPontoErro = sPontoErro & "," & "111"
'        'Indica que a planilha será impressa no estilo Landscape
'        objPlanilhaExcel.PageSetup.Orientation = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE
'
'        'Indica que a planilha será impressa com linhas de grade
'        objPlanilhaExcel.PageSetup.PrintGridlines = EXCEL_PLANILHA_IMPRIME_GRADE
'
'        'Imprime o gráfico e fecha o excel
'        objPlanilhaExcel.PrintOut

'sPontoErro = sPontoErro & "," & "112"
        lErro = objExcelApp.Imprimir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'        'Fecha o Excel
'        Call Excel_Fechar
'sPontoErro = sPontoErro & "," & "113"
        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
'        GL_objExcelSist.Visible = True
'sPontoErro = sPontoErro & "," & "114"
        lErro = objExcelApp.Mudar_Visibilidade(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'sPontoErro = sPontoErro & "," & "115"
    End If
    
'sPontoErro = sPontoErro & "," & "116"

    Browse_Executa_SQL_Excel = SUCESSO
    
    Exit Function
    
Erro_Browse_Executa_SQL_Excel:

    Browse_Executa_SQL_Excel = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 202173 To 202178, 202180, 202181, 202182, 202183, 202186
            
        Case 202179
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 202184, 202185
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, objBrowse.objBrowseArquivo.sNomeArq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202187)

    End Select

    'Se não foi cancelado
    'If Browse_Executa_SQL_Excel <> 202186 Then
    '    Call Rotina_Erro(vbOKOnly, "Rastreamento por posição: ..." & right(sPontoErro, 50), gErr)
    '    Call Rotina_Erro(vbOKOnly, "Complemento: " & sErroAux, gErr)
    'End If

    Call Comando_Fechar(lComando)
    
    If bArqAberto Then
        Close #1
        Kill objBrowse.objBrowseExcel.sLocalizacaoCSV & sNomeArq
    End If
    
    'Fecha o Excel
    'Call Excel_Fechar
    
    Call objExcelApp.Fechar

    If Not bModal Then
        If Not (objFrmAguardeTela Is Nothing) Then
            Call objFrmAguardeTela.Trata_Erro
        End If
    End If
    
    Exit Function

End Function

Private Sub Browse_Executa_SQL_Excel1(ByVal objBrowse As AdmBrowse, colIndice As Collection, colBrowseUsuarioCampo As Collection, colCampos As Collection)
'Acerta a ordem dos valores em avCampos
'1 - Respeita as alterações de posicionamento
'2 - Retira os campos de exercício acrescentados dinamicamente e retorna o exercicio se existir

Dim objBrowseUsuCampo As AdmBrowseUsuarioCampo
Dim objCampo As AdmCampos
Dim iColuna As Integer
Dim iPulaCampo As Integer
Dim iCampo As Integer
Dim colCamposAux As New Collection

On Error GoTo Erro_Browse_Executa_SQL_Excel1

    'Remove Exercício que não estão em BrowseUsuarioCampos
    iCampo = 0
    iPulaCampo = 0
    For Each objBrowseUsuCampo In colBrowseUsuarioCampo
        iCampo = iCampo + 1
        Set objCampo = New AdmCampos
        objCampo.iTipo = objBrowse.colValorCampo.Item(iCampo + iPulaCampo).iTipo
        objCampo.iSubTipo = objBrowse.colValorCampo.Item(iCampo + iPulaCampo).iSubTipo
        colCamposAux.Add objCampo
        If left(objBrowseUsuCampo.sNome, 7) = "Periodo" And left(objBrowseUsuCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuCampo.sNome, 10) <> "PeriodoAte" Then
            iPulaCampo = iPulaCampo + 1
        End If
    Next

    'Acerta a Ordenação
    iCampo = 0
    For Each objBrowseUsuCampo In colBrowseUsuarioCampo
        iCampo = iCampo + 1
        iColuna = colIndice(iCampo)
        colCampos.Add colCamposAux.Item(iColuna)
    Next
    
    Exit Sub
    
Erro_Browse_Executa_SQL_Excel1:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202187)

    End Select
    
    Exit Sub

End Sub

Private Sub Browse_Executa_SQL_Excel2(ByVal objBrowse As AdmBrowse, avCampo() As Variant, avCampoOrd() As Variant, colIndice As Collection, colBrowseUsuarioCampo As Collection, iExercicio As Integer)
'Acerta a ordem dos valores em avCampos
'1 - Respeita as alterações de posicionamento
'2 - Retira os campos de exercício acrescentados dinamicamente e retorna o exercicio se existir

Dim objBrowseUsuCampo As AdmBrowseUsuarioCampo
Dim iColuna As Integer
Dim iPulaCampo As Integer
Dim iCampo As Integer
Dim avCampoAux(1 To BROWSER_NUM_MAX_CAMPOS) As Variant

On Error GoTo Erro_Browse_Executa_SQL_Excel2

    'Remove Exercício que não estão em BrowseUsuarioCampos
    iCampo = 0
    iPulaCampo = 0
    For Each objBrowseUsuCampo In colBrowseUsuarioCampo
        iCampo = iCampo + 1
        avCampoAux(iCampo) = avCampo(iCampo + iPulaCampo)
        If left(objBrowseUsuCampo.sNome, 7) = "Periodo" And left(objBrowseUsuCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuCampo.sNome, 10) <> "PeriodoAte" Then
            iPulaCampo = iPulaCampo + 1
            If iExercicio = 0 Then iExercicio = avCampo(iCampo + 1)
        End If
    Next

    'Acerta a Ordenação
    iCampo = 0
    For Each objBrowseUsuCampo In colBrowseUsuarioCampo
        iCampo = iCampo + 1
        iColuna = colIndice(iCampo)
        avCampoOrd(iCampo) = avCampoAux(iColuna)
    Next
    
    Exit Sub
    
Erro_Browse_Executa_SQL_Excel2:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202187)

    End Select
    
    Exit Sub

End Sub

Public Function Excel_Gera_Planilha_Grid(ByVal objGrid As AdmGrid) As Long

Dim lErro As Long
Dim iLinha As Long
Dim iColuna As Integer
Dim dLarguraColuna As Double, vValor As Variant, sFormato As String
Dim vbResp As VbMsgBoxResult, iModoImpressao As Integer

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
'Dim objGrafico As Object
Dim iLin As Long
Dim iCol As Integer
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Gera_Planilha_Grid

    vbResp = Rotina_Aviso(vbYesNo, "AVISO_IMPRIMIR_PLANILHA")
    If vbResp = vbYes Then iModoImpressao = EXCEL_MODO_IMPRESSAO

'    lErro = Excel_Abrir()
'    If lErro <> SUCESSO Then gError 102106

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    'Cria uma nova pasta de trabalho
'    Set objPastaTrabalho = GL_objExcelSist.Workbooks.Add
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    lErro = objExcelApp.Criar_Planilha()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'Para cada coluna que deverá receber dados
    For iColuna = 1 To objGrid.colColuna.Count
        
        'E para cada linha que receberá dados
        For iLinha = 0 To objGrid.iLinhasExistentes
                                
'            'Monta a Faixa de Células da coluna atual
'            Set objFaixaCelulasPreencher = objPlanilhaExcel.Range(objPlanilhaExcel.Cells(iLinha + 1, iColuna), objPlanilhaExcel.Cells(iLinha + 1, iColuna))
            
            vValor = objGrid.objGrid.TextMatrix(iLinha, iColuna - 1)
            sFormato = ""
            If IsDate(vValor) And (InStr(1, vValor, ":") <> 0 Or InStr(1, vValor, "/") <> 0) Then
                If InStr(1, vValor, ":") <> 0 Then
                    vValor = StrParaDbl(CStr(vValor))
                    vValor = CDate(vValor)
                    sFormato = "hh:mm:ss"
                Else
                    vValor = StrParaDate(CStr(vValor))
                End If
            ElseIf IsNumeric(vValor) Then
                    
                If Mid(vValor, 1, 1) = "0" And Mid(vValor, 2, 1) <> "," And Len(Trim(vValor)) > 1 Then
                    vValor = "'" & vValor
                ElseIf InStr(1, vValor, ":") <> 0 Then
                    vValor = StrParaDbl(CStr(vValor))
                    vValor = CDate(vValor)
                    sFormato = "hh:mm:ss"
                ElseIf InStr(1, vValor, "%") <> 0 Then
                    vValor = StrParaDbl(CStr(vValor))
                    sFormato = "0.00%"
                ElseIf InStr(1, vValor, ",") <> 0 Then
                    vValor = StrParaDbl(CStr(vValor))
                    sFormato = "#,##0.00##"
                End If
       
            End If
            
'            objFaixaCelulasPreencher.Value = vValor
'
'            If sFormato <> "" Then objFaixaCelulasPreencher.NumberFormat = sFormato
'
'            If iLinha = 0 Then
'
'                'Obtém a largura da coluna no padrão do Excel
'                lErro = CF("Excel_Obtem_Largura_Coluna", objGrid.objGrid.ColWidth(iColuna - 1), dLarguraColuna)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                objFaixaCelulasPreencher.Font.Bold = True
'                'Altera a largura da coluna conforme a largura passada
'                objFaixaCelulasPreencher.ColumnWidth = dLarguraColuna
'
'            End If

            lErro = objExcelApp.Preencher_Celula(iLinha + 1, iColuna, vValor, sFormato, IIf(iLinha = 0, True, False))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
    
    Next
    
'    GL_objExcelSist.Cells.Select
'    GL_objExcelSist.Cells.EntireColumn.AutoFit
'    GL_objExcelSist.Range("A1", "A1").Select

    lErro = objExcelApp.Ajustar_Tam_Colunas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
  
    'Se é para apenas imprimir a planilha
    If iModoImpressao = EXCEL_MODO_IMPRESSAO Then

'        'Indica que a planilha será impressa no estilo Landscape
'        objPlanilhaExcel.PageSetup.Orientation = EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE
'
'        'Indica que a planilha será impressa com linhas de grade
'        objPlanilhaExcel.PageSetup.PrintGridlines = EXCEL_PLANILHA_IMPRIME_GRADE
'
'        'Imprime o gráfico e fecha o excel
'        objPlanilhaExcel.PrintOut

        lErro = objExcelApp.Imprimir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'        'Fecha o Excel
'        Call Excel_Fechar

        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
        
'        'Torna o Excel Visível
'        GL_objExcelSist.Visible = True
                
        lErro = objExcelApp.Mudar_Visibilidade(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Excel_Gera_Planilha_Grid = SUCESSO
    
    Exit Function
    
Erro_Excel_Gera_Planilha_Grid:

    Excel_Gera_Planilha_Grid = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 102106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142313)

    End Select

'    'Fecha o Excel
'    Call Excel_Fechar

    Call objExcelApp.Fechar

    Exit Function

End Function
