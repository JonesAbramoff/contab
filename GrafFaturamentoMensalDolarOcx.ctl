VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl GrafFaturamentoMensalDolarOcx 
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LockControls    =   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   3240
   Begin VB.Frame Frame1 
      Caption         =   "Meses"
      Height          =   1005
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin MSMask.MaskEdBox Ano 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   855
         TabIndex        =   4
         Top             =   450
         Width           =   405
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   360
      Left            =   1725
      Picture         =   "GrafFaturamentoMensalDolarOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Fechar"
      Top             =   1395
      Width           =   1305
   End
   Begin VB.CommandButton BotaoGrafico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      Picture         =   "GrafFaturamentoMensalDolarOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Gerar Gr�fico"
      Top             =   1395
      Width           =   1305
   End
End
Attribute VB_Name = "GrafFaturamentoMensalDolarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iAlterado As Integer

'??? Jones:transferir p/outro lugar
Private Const MOEDA_DOLAR = 1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Comparativo Mensal em Dolar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GrafFaturamentoMensalDolar"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub Ano_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ano_GotFocus()
     
     Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******


Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGrafico_Click()

Dim lErro As Long
Dim iAno As Integer

On Error GoTo Erro_BotaoGrafico_Click

    'Se o Ano n�o foi digitado => ERRO
    If Len(Trim(Ano.ClipText)) = 0 Then gError 90056
    
    'Guarda o Ano que servir� como base para obten��o dos resultados
    iAno = StrParaInt(Ano.Text)
        
    'Obt�m os dados que ser�o utilizados para gerar a planilha que servir� de base ao gr�fico
    'Chama a fun��o que montar� o gr�fico no excel
    lErro = Gera_Grafico(giFilialEmpresa, iAno)
    If lErro <> SUCESSO Then gError 90057
   
    Unload Me
    
    Exit Sub
    
Erro_BotaoGrafico_Click:
    
    Select Case gErr

        Case 90056
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOFATURAMENTO_VAZIO", gErr)
            
        Case 90057
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161693)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Private Function Gera_Grafico(iFilialEmpresa As Integer, iAno As Integer) As Long
'Fun��o revisada em 03/06/2001 por Luiz Gustavo de Freitas Nogueira
'Essa fun��o gera os dados que montar�o a planilha-base do gr�fico
'E seta todas as configura��es do gr�fico, como t�tulo, cabe�alho, legenda, etc.
'Ap�s gerados os dados e setadas as configura��es � chamada a fun��o que faz a interface com o Excel

Dim lErro As Long
Dim dMediaFaturamento As Double
Dim dTotalAno As Double
Dim iMesesSemFat As Integer
Dim iMesAtual As Integer
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim objColunaMeses As New ClassColunasExcel
Dim objColunaFaturamento As New ClassColunasExcel
Dim objColunaMediaMensal As New ClassColunasExcel
Dim colLinhasMeses As New Collection
Dim colLinhasFaturamento As New Collection
Dim colLinhasMediaMensal As New Collection
Dim objCelMeses As ClassCelulasExcel
Dim objCelFaturamento As ClassCelulasExcel
Dim objCelMediaMensal As ClassCelulasExcel
Dim objPlanilha As New ClassPlanilhaExcel
Dim adValoresFaturados(1 To 12) As Double
Dim asMes(1 To 12) As String

On Error GoTo Erro_Gera_Grafico

    'Exibe o ponteiro Ampulheta
    MousePointer = vbHourglass

    'Informa ao excel o nome da planilha que exibe o gr�fico
    objPlanilha.sNomeGrafico = "Gr�f.-Comparat.Mensal em D�lar"
    
    'Informa ao excel o nome da planilha que exibe os dados do gr�fico
    objPlanilha.sNomePlanilha = "Comparativo Mensal em D�lar"

    '*** GERA OS DADOS DO GR�FICO ***
    
    'Preenche o array com os meses do ano
    'Esse array ser� utilizado para preencher a coluna de meses
    asMes(1) = "Janeiro"
    asMes(2) = "Fevereiro"
    asMes(3) = "Mar�o"
    asMes(4) = "Abril"
    asMes(5) = "Maio"
    asMes(6) = "Junho"
    asMes(7) = "Julho"
    asMes(8) = "Agosto"
    asMes(9) = "Setembro"
    asMes(10) = "Outubro"
    asMes(11) = "Novembro"
    asMes(12) = "Dezembro"
    
    ' *** COLUNA MESES ***
    'Configura essa coluna como integrante do eixo X
    objColunaMeses.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
    
    'Configura o gr�fico Linha 3D como gr�fico a ser utilizado para essa coluna
    objColunaMeses.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Configura o gr�fico para n�o exibir os DataLabels referentes a essa coluna
    objColunaMeses.lDataLabels = EXCEL_NAO_EXIBE_LABELS
    
    '*********************
    
    ' *** COLUNA FATURAMENTO ***
    'Configura essa coluna como integrante do eixo Y
    objColunaFaturamento.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Configura o gr�fico Linha 3D como gr�fico a ser utilizado para essa coluna
    objColunaFaturamento.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Configura o gr�fico para exibir valores como DataLabels referentes a essa coluna
    objColunaFaturamento.lDataLabels = EXCEL_EXIBE_LABELS_VALOR
    
    'Configura o gr�fico para exibir os DataLabels no sentido horizontal
    objColunaFaturamento.lDataLabelsOrientacao = EXCEL_LABEL_ORIENTACAO_HORIZONTAL
    '***********************
    
    ' *** COLUNA M�DIA MENSAL ***
    'Informa ao excel em qual eixo essa coluna far� parte do gr�fico
    objColunaMediaMensal.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Informa ao excel o tipo de gr�fico que ser� usado para representar essa coluna
    objColunaMediaMensal.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Informa ao excel como ser�o exibidos os Datalabels para essa coluna
    objColunaMediaMensal.lDataLabels = EXCEL_NAO_EXIBE_LABELS
    '***********************
    
    'Define as datas de in�cio e fim para leitura do BD
    'Essa data ser� sempre o primeiro dia do ano (data in�cio) e o �ltimo dia do ano (data fim)
    dtDataIni = StrParaDate("01/01/" & iAno)
    dtDataFim = StrParaDate("31/12/" & iAno)
    
    'Monta as cole��es de linhas para cada coluna
    
    'Instancia os obj�s que ir�o armazenar as c�lulas com os t�tulos de cada coluna
    Set objCelMeses = New ClassCelulasExcel
    Set objCelFaturamento = New ClassCelulasExcel
    Set objCelMediaMensal = New ClassCelulasExcel
    
    'Guarda nos obj�s os t�tulos de cada coluna
    objCelMeses.vValor = "Meses"
    objCelFaturamento.vValor = "Faturamento"
    objCelMediaMensal.vValor = "M�dia Mensal"
    
    'Guarda os t�tulos das colunas nas cole��es de linhas de cada coluna
    colLinhasMeses.Add objCelMeses
    colLinhasFaturamento.Add objCelFaturamento
    colLinhasMediaMensal.Add objCelMediaMensal
    
    lErro = FatMensal_MoedaExt_Le(MOEDA_DOLAR, iFilialEmpresa, dtDataIni, dtDataFim, adValoresFaturados)
    If lErro <> SUCESSO Then gError 90563
    
    'Para cada m�s
    'Calcula o valor do faturamento do m�s, convertido ao valor em d�lar referente � data do faturamento
    For iMesAtual = 1 To 12
    
        'Instacia o obj que receber� o conte�do das c�lulas da coluna Meses
        Set objCelMeses = New ClassCelulasExcel
        
        'Guarda o valor da c�lula (nome do m�s iMesAtual) no obj
        objCelMeses.vValor = asMes(iMesAtual)
        
        'Guarda a c�lula na cole��o
        colLinhasMeses.Add objCelMeses

        'Se n�o houve faturamento no m�s => incrementa a vari�vel que conta os meses em que n�o houve faturamento
        If adValoresFaturados(iMesAtual) = 0 Then iMesesSemFat = iMesesSemFat + 1
        
        'Soma o valor faturado no m�s ao total faturado no ano
        dTotalAno = dTotalAno + adValoresFaturados(iMesAtual)
        
        'Instacia o obj que receber� o conte�do das c�lulas da coluna Faturamento
        Set objCelFaturamento = New ClassCelulasExcel
        
        'Guarda o valor da c�lula (nome do m�s iMesAtual) no obj
        objCelFaturamento.vValor = Format(adValoresFaturados(iMesAtual), "Standard")
        
        'Preenche a linha com o valor do Faturamento de iMesAtual
        colLinhasFaturamento.Add objCelFaturamento
        
    Next
    
    'Se em nenhum m�s houve faturamento
    If iMesesSemFat = 12 Then
        
        'A m�dia do faturamento � zero
        dMediaFaturamento = 0
        
    'Sen�o, ou seja, se houve faturamento em pelo menos um m�s
    Else
    
        'Calcula a m�dia de faturamento mensal
        'Ou seja, divide o total anual pelo n�mero de meses em que houve faturamento
        dMediaFaturamento = dTotalAno / (12 - iMesesSemFat)
    
    End If
    
    'Preenche as linhas da coluna M�dia Mensal
    For iMesAtual = 1 To 12
    
        'Instacia o obj que receber� o conte�do das c�lulas da coluna M�dia Mensal
        Set objCelMediaMensal = New ClassCelulasExcel
        
        'Guarda o valor da c�lula (nome do m�s iMesAtual) no obj
        objCelMediaMensal.vValor = dMediaFaturamento
        
        '*** COLUNA M�DIA MENSAL
        colLinhasMediaMensal.Add objCelMediaMensal
        '***********************
    
    Next
    
    'Transfere as cole��es de c�lulas de cada coluna para os objetos que guardam os dados de cada coluna
    Set objColunaMeses.colCelulas = colLinhasMeses
    Set objColunaFaturamento.colCelulas = colLinhasFaturamento
    Set objColunaMediaMensal.colCelulas = colLinhasMediaMensal
    
    'Guarda na cole��o de colunas as cole��es de c�lulas de cada coluna
    objPlanilha.colColunas.Add objColunaMeses
    objPlanilha.colColunas.Add objColunaFaturamento
    objPlanilha.colColunas.Add objColunaMediaMensal
    
    '*** SETAGEM DE OUTRAS CONFIGURA��ES DO GR�FICO (CABE�ALHO, T�TULO, ETC.)
    
    'Informa ao excel o t�tulo do gr�fico
    objPlanilha.sTituloGrafico = "Ano: " & iAno & vbCrLf & "M�dia Mensal em D�lares: US$ " & CStr(Format(dMediaFaturamento, "Standard"))

    'Instancia a cole��o que guardar� as se��es de cabe�alho / rodap�
    Set objPlanilha.colCabecalhoRodape = New Collection

    'Monta o cabe�alho e o rodap� do Gr�fico
    lErro = Grafico_Monta_Cabecalho_Rodape(objPlanilha.colCabecalhoRodape)
    If lErro <> SUCESSO Then gError 90539

    'Informa ao excel a posi��o da legenda
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_DIREITA
    
    'Informa ao excel a posi��o dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW

    'Informa ao excel a orienta��o dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD

    'Informa ao excel que a plotagem do dados ser� por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS

    'Monta a planilha e o gr�fico com os dados passados em objPlanilha
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 79972

    MousePointer = vbDefault
    
    'Exibe o ponteiro padr�o
    MousePointer = vbDefault
    
    Gera_Grafico = SUCESSO

    Exit Function

Erro_Gera_Grafico:

    Gera_Grafico = gErr

    Select Case gErr

        Case 79972, 90539, 90563
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161694)

    End Select
    
    'Exibe o ponteiro padr�o
    MousePointer = vbDefault
    
    Exit Function

End Function

Public Function Grafico_Monta_Cabecalho_Rodape(colLinhas As Collection) As Long
'Fun��o criada em 29/05/01 por Luiz Gustavo de Freitas Nogueira
'Essa fun��o preenche os objetos com os dados de cada linha que ser� exibida no cabe�alho

Dim objLinha As ClassLinhaCabecalhoExcel

On Error GoTo Erro_Grafico_Monta_Cabecalho_Rodape

        '*** PREENCHIMENTO DO CABE�ALHO ESQUERDO ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
        
            ' *** LINHA 01 - CABE�ALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 1 'Indica quantas linhas devem existir entre essa linha e a pr�xima
            objLinha.iLinha = 1 'Indica a posi��o da linha no cabe�alho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha � cole��o de linhas de cabe�alho / rodap�
            colLinhas.Add objLinha
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 02 - CABE�ALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeFilialEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 2 'Indica quantas linhas devem existir entre essa linha e a pr�xima
            objLinha.iLinha = 2 'Indica a posi��o da linha no cabe�alho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha � cole��o de linhas de cabe�alho / rodap�
            colLinhas.Add objLinha
            
        ' *** FIM DO CABE�ALHO ESQUERDO ***
        
        ' *** PREENCHIMENTO DO CABE�ALHO CENTRAL ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABE�ALHO CENTRAL ***
            objLinha.iSecao = EXCEL_CABECALHO_CENTRAL
            objLinha.sTexto = "Comparativo Mensal em Dolar"
            objLinha.sFonte = EXCEL_FONTE_BOOKMAN
            objLinha.iTamanhoFonte = 20
            objLinha.sNegrito = EXCEL_CABECALHO_RODAPE_NEGRITO
            objLinha.iEspacoLinha = 0 'Indica quantas linhas devem existir entre essa linha e a pr�xima
            objLinha.iLinha = 1 'Indica a posi��o da linha no cabe�alho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_CENTRAL
            
            'Adiciona a linha � cole��o de linhas de cabe�alho / rodap�
            colLinhas.Add objLinha

        ' *** FIM DO CABE�ALHO CENTRAL ***
        
        ' *** PREENCHIMENTO DO CABE�ALHO DIREITO ***
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABE�ALHO DIREITO ***
            objLinha.iSecao = EXCEL_CABECALHO_DIREITO
            objLinha.sTexto = CStr(Date)
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = EXCEL_CABECALHO_RODAPE_NAO_QUEBRA_LINHA 'Indica quantas linhas devem existir entre essa linha e a pr�xima
            objLinha.iLinha = 1 'Indica a posi��o da linha no cabe�alho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_DIREITA
            
            'Adiciona a linha � cole��o de linhas de cabe�alho / rodap�
            colLinhas.Add objLinha

        ' *** FIM DO CABE�ALHO DIREITO ***
        
    Grafico_Monta_Cabecalho_Rodape = SUCESSO
    
    Exit Function
    
Erro_Grafico_Monta_Cabecalho_Rodape:

    Grafico_Monta_Cabecalho_Rodape = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161695)
            
    End Select
    
    Exit Function

End Function

Public Function FatMensal_MoedaExt_Le(iMoeda As Integer, iFilialEmpresa As Integer, dtDataIni As Date, dtDataFim As Date, adValoresFaturados() As Double) As Long
'Fun��o criada em 01/07/2001 por Luiz Gustavo de Freitas Nogueira
'Essa fun��o l� no BD o valor faturado para cada dia de cada m�s
'Converte esse valor para iMoeda e acumula os valores convertidos de cada m�s
'E devolve esses valores dentro de um array

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iMesAtual As Integer
Dim dValorCotacao As Double
Dim dValorFaturado As Double
Dim dtData As Date
Dim dTotalMes As Double
Dim sFiltro As String

On Error GoTo Erro_FatMensal_MoedaExt_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90030
     
    'Abre o Comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 90031
        
    If iFilialEmpresa = EMPRESA_TODA Then
    
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
    
    End If
        
    'Para cada m�s
    'Calcula o valor do faturamento do m�s, convertido ao valor em d�lar referente � data do faturamento
    For iMesAtual = 1 To 12
    
        If iFilialEmpresa = 0 Then
        
            'Empresa toda
            'Calcula o valor que ser� jogado na linha da coluna Faturamento
            'Ou seja, calcula o faturamento total para iMesAtual
            'L� em SldDiaFat, o valor faturado em cada dia de iMesAtual
            lErro = Comando_Executar(lComando, "SELECT Data, ValorFaturado FROM SldDiaFat WHERE Data >= ? AND Data <= ?  AND {fn MONTH(Data)} = ? " & sFiltro & " ORDER BY Data", dtData, dValorFaturado, dtDataIni, dtDataFim, iMesAtual)
            If lErro <> AD_SQL_SUCESSO Then gError 90032
        
        Else
        
            'Calcula o valor que ser� jogado na linha da coluna Faturamento
            'Ou seja, calcula o faturamento total para iMesAtual
            'L� em SldDiaFat, o valor faturado em cada dia de iMesAtual
            lErro = Comando_Executar(lComando, "SELECT Data, ValorFaturado FROM SldDiaFat WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ?  AND {fn MONTH(Data)} = ? ORDER BY Data", dtData, dValorFaturado, iFilialEmpresa, dtDataIni, dtDataFim, iMesAtual)
            If lErro <> AD_SQL_SUCESSO Then gError 90032
        
        End If
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90033
        
        'Enquanto houver Dados
        'Ou seja, para cada dia do m�s, no qual houve faturamento
        Do While lErro <> AD_SQL_SEM_DADOS
                    
            'L� a Cotacao do dia em CotacoesMoeda
            lErro = Comando_Executar(lComando1, "SELECT Valor FROM CotacoesMoeda WHERE Data=? AND Moeda=? ", dValorCotacao, dtData, iMoeda)
            If lErro <> AD_SQL_SUCESSO Then gError 90035
    
            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90036
            
            'Se n�o foi encontrada cota��o para o dia dtData
            'dtData foi encontrada na leitura feita em SldDiaFat
            If lErro = AD_SQL_SEM_DADOS Then gError 90037
            
            'Se o valor da cota��o for diferente de zero
            If dValorCotacao <> 0 Then
                
                'Soma ao valor do faturamento desse m�s, o valor do faturamento do dia convertido para d�lar
                dTotalMes = dTotalMes + dValorFaturado / dValorCotacao
                
            End If
            
            'Busca o pr�ximo dia que teve faturamento para o iMesAtual
            'Obs.: esse comando_buscarproximo refere-se ao comando_executar utilizado no in�cio da fun��o (ver lcomando)
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90038
    
        Loop
        
        'Guarda no array o valor calculado para iMesAtual
        adValoresFaturados(iMesAtual) = dTotalMes
    
        'Limpa a vari�vel que totaliza o valor do faturamento do m�s
        dTotalMes = 0
    
    Next

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Comando
    Call Comando_Fechar(lComando1)

    FatMensal_MoedaExt_Le = SUCESSO
    
    Exit Function
    
Erro_FatMensal_MoedaExt_Le:

    FatMensal_MoedaExt_Le = gErr
    
    Select Case gErr
    
        Case 90030, 90031
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90032, 90033, 90038
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT", gErr)
            
        Case 90034
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDDIAFAT_NAO_ENCONTRADO", gErr)
        
        Case 90035, 90036
            
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr)
        
        Case 90037
            Call Rotina_Erro(vbOKOnly, "ERRO_COTACOESMOEDA_INEXISTENTE1", gErr, dtData)
        
        Case 177597
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161696)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function
