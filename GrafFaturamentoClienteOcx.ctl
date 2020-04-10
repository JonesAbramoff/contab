VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl GrafFaturamentoCliOcx 
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LockControls    =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   5805
   Begin VB.CommandButton BotaoFechar 
      Height          =   360
      Left            =   2970
      Picture         =   "GrafFaturamentoClienteOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fechar"
      Top             =   1245
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
      Left            =   1290
      Picture         =   "GrafFaturamentoClienteOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Gerar Gr�fico"
      Top             =   1245
      Width           =   1305
   End
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo Per�odo"
      Height          =   900
      Left            =   135
      TabIndex        =   4
      Top             =   165
      Width           =   5520
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   780
         TabIndex        =   0
         Top             =   375
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDe 
         Height          =   300
         Left            =   1935
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   390
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3420
         TabIndex        =   1
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownAte 
         Height          =   300
         Left            =   4590
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "At�:"
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
         Left            =   2985
         TabIndex        =   8
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   345
         TabIndex        =   7
         Top             =   420
         Width           =   315
      End
   End
End
Attribute VB_Name = "GrafFaturamentoCliOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Const NUMERO_CLIENTES_GRAFICO_FATCLI = 10

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Faturamento por Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GrafFaturamentoCli"
    
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

Private Sub DataAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataAte_GotFocus()
     
     Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataAte_Validate

    'Verifica se a Data Final foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 90044
    
     'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 90072

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de cr�tica, segura o foco
        Case 90044
        
        Case 90072
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161684)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_DataDe_Validate

    'Verifica se a Data Inicial foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 90045

    'Compara com a data Fianal
    If Len(Trim(DataAte.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 90071

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de cr�tica, segura o foco
        Case 90045
        
        Case 90071
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161685)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro Then gError 90046

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 90046

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161686)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro Then gError 90047

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 90047

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161687)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro Then gError 90048

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 90048

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161688)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro Then gError 90049

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 90049

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161689)

    End Select

    Exit Sub

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
Dim dtDataIni As Date
Dim dtDataFim As Date

On Error GoTo Erro_BotaoGrafico_Click

    'Exibe o ponteiro ampulheta
    MousePointer = vbHourglass
    
    'Se a data inicial n�o foi preenchida => erro
    If Len(Trim(DataDe.ClipText)) = 0 Then gError 90050
        
    'Se a data final n�o foi preenchida => erro
    If Len(Trim(DataAte.ClipText)) = 0 Then gError 90051
      
    'Guarda as datas de in�cio e fim do per�odo que servir� de base para o gr�fico
    dtDataIni = StrParaDate(DataDe.Text)
    dtDataFim = StrParaDate(DataAte.Text)
    
    'Obt�m os dados que ser�o utilizados para gerar a planilha que servir� de base ao gr�fico
    'Chama a fun��o que montar� o gr�fico no excel
    lErro = Gera_Grafico(giFilialEmpresa, dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 90043
    
    'Exibe o ponteiro padr�o
    MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGrafico_Click:
    
    Select Case gErr

        Case 90043
        
        Case 90050
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", gErr)
           
        Case 90051
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161690)
            
    End Select
    
    'Exibe o ponteiro padr�o
    MousePointer = vbDefault
    
    Exit Sub

End Sub

Public Sub Form_Load()
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Private Function Gera_Grafico(iFilialEmpresa As Integer, dtDataIni As Date, dtDataFim As Date) As Long
'Fun��o revisada em 04/06/2001 por Luiz Gustavo de Freitas Nogueira
'Essa fun��o gera os dados que montar�o a planilha-base do gr�fico
'E seta todas as configura��es do gr�fico, como t�tulo, cabe�alho, legenda, etc.
'Ap�s gerados os dados e setadas as configura��es � chamada a fun��o que faz a interface com o Excel

Dim lErro As Long
Dim lComando As Long
Dim sNomeRed As String
Dim dPercent As Double
Dim dTotal As Double
Dim dValorSoma As Double
Dim dValorOutros As Double
Dim iCliente As Integer
Dim iIndice As Integer
Dim colLinhasClientes As New Collection
Dim colLinhasFaturamentoCliente As New Collection
Dim objColunaClientes As New ClassColunasExcel
Dim objColunaFaturamentoCliente As New ClassColunasExcel
Dim objCelClientes As ClassCelulasExcel
Dim objCelFaturamentoCliente As ClassCelulasExcel
Dim objPlanilha As New ClassPlanilhaExcel
Dim sFiltro As String

On Error GoTo Erro_Gera_Grafico

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90022
    
    'Informa ao excel o nome da planilha que exibe o gr�fico
    objPlanilha.sNomeGrafico = "Gr�f. - Faturamento por Cliente"
    
    'Informa ao excel o nome da planilha que exibe os dados do gr�fico
    objPlanilha.sNomePlanilha = "Faturamento por Cliente"
    
    '*** COLUNA CLIENTES ****
    'T�tulo da Coluna
    'Instancia um obj que armazenar� o t�tulo da coluna Clientes
    Set objCelClientes = New ClassCelulasExcel
    
    'Guarda no obj o t�tulo da coluna
    objCelClientes.vValor = "Clientes"
    
    'Guarda o obj na cole��o de c�lulas
    colLinhasClientes.Add objCelClientes
    '**************************
    
    '*** COLUNA FATURAMENTO CLIENTE ****
    'T�tulo da Coluna
    'Instancia um obj que armazenar� o t�tulo da coluna Faturamento Cliente
    Set objCelFaturamentoCliente = New ClassCelulasExcel
    
    'Guarda no obj o t�tulo da coluna
    objCelFaturamentoCliente.vValor = "Faturamento Cliente"
    
    'Guarda o obj na cole��o de c�lulas
    colLinhasFaturamentoCliente.Add objCelFaturamentoCliente
    '**************************

    'Inicializa a vari�vel que receber� o nome reduzido do cliente
    sNomeRed = String(STRING_BUFFER_MAX_TEXTO, 0)

    'Se a filial passada indica que est� sendo acessada a empresa como um todo
    If iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        'L� na View FaturamentoCliente, o valor faturado para esse cliente por todas as filiais da empresa
        lErro = Comando_Executar(lComando, "SELECT NomeReduzido, SUM(Soma) AS Soma From FaturamentoCliente WHERE DataEmissao>= ? AND DataEmissao<= ? " & sFiltro & " GROUP BY NomeReduzido ORDER BY Soma DESC", _
        sNomeRed, dValorSoma, dtDataIni, dtDataFim)
    
    'Sen�o
    'Ou seja, se est� sendo acessada apenas uma filial
    Else
        
        'L� na View FaturamentoCliente, o valor faturado para esse cliente pela filial que est� sendo acessada
        lErro = Comando_Executar(lComando, "SELECT NomeReduzido, SUM(Soma) AS Soma From FaturamentoCliente WHERE FilialEmpresa=? AND DataEmissao>= ? AND DataEmissao<= ? GROUP BY NomeReduzido ORDER BY Soma DESC", _
        sNomeRed, dValorSoma, iFilialEmpresa, dtDataIni, dtDataFim)
        
    End If
        
    If lErro <> AD_SQL_SUCESSO Then gError 90023

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90024
    
    'Essa vari�vel contar� o n�mero de clientes para os quais existem valor faturado
    iCliente = 1
    
    'Enquanto houver Dados, adiciona na Cole��o
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Se esse cliente n�o est� entre os 10 maiores faturamento
        If iCliente > NUMERO_CLIENTES_GRAFICO_FATCLI Then
            
            'Soma o valor do faturamento dele ao total de "Outros"
            dValorOutros = dValorOutros + dValorSoma
         
         'Se ele est� entre os 10 maiores
         Else
            
            ' *** COLUNA CLIENTES ***
            'Nome do Cliente
            'Instancia um novo obj que armazenar� o nome do cliente
            Set objCelClientes = New ClassCelulasExcel
            
            'Guarda no obj o nome do cliente iCliente
            objCelClientes.vValor = sNomeRed
            
            'Guarda o obj na cole��o de c�lulas
            colLinhasClientes.Add objCelClientes
            '*************************
            
            ' *** COLUNA FATURAMENTO CLIENTE ***
            'Valor faturado para o cliente
            'Instancia um obj que armazenar� o valor do faturamento para iCliente
            Set objCelFaturamentoCliente = New ClassCelulasExcel
            
            'Guarda no obj o valor do faturamento para iCliente
            objCelFaturamentoCliente.vValor = Format(dValorSoma, "Standard")
            
            'Guarda o obj na cole��o de c�lulas
            colLinhasFaturamentoCliente.Add objCelFaturamentoCliente
            '*************************************
            
        End If
        
        'Incrementa a vari�vel que conta o n�mero de clientes
        iCliente = iCliente + 1
           
        'Acumula o valor total faturado
        dTotal = dTotal + dValorSoma
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90025

    Loop
    
    'Se existem mais de 10 clientes com valor faturado
    'Cria a linha "Outros"
    If iCliente > NUMERO_CLIENTES_GRAFICO_FATCLI Then
    
        ' *** COLUNA CLIENTES ***
        'Cliente "Outros"
        'Instancia um obj que armazenar� o t�tulo da linha Outros
        'Essa linha armazenar� o somat�rio faturado para os clientes que n�o fazem parte
        'do grupo dos 10 maiores faturamento
        Set objCelClientes = New ClassCelulasExcel
        
        'Guarda no obj o nome do cliente "Outros"
        objCelClientes.vValor = "Outros"
        
        'Guarda o obj na cole��o de c�lulas
        colLinhasClientes.Add objCelClientes
        '*************************
        
        '*** COLUNA FATURAMENTO CLIENTE ***
        'Valor faturado
        'Instancia um obj que armazenar� o nome do cliente "Outros"
        Set objCelFaturamentoCliente = New ClassCelulasExcel
        
        'Guarda no obj o valor do faturamento para "Outros"
        objCelFaturamentoCliente.vValor = dValorOutros
        
        'Guarda o obj na cole��o de c�lulas
        colLinhasFaturamentoCliente.Add objCelFaturamentoCliente
    
    End If
    
    'Concatena ao nome do cliente o percentual que o valor faturado para o mesmo, representa dentro do faturamento total da empresa
    'Para cada Cliente (iIndice come�a em 2 porque a primeira linha refere-se ao t�tulo da coluna
    For iIndice = 2 To iCliente
    
        'Se existem mais de NUMERO_CLIENTES_GRAFICO_FATCLI clientes com valor faturado e iIndice � maior igual a NUMERO_CLIENTES_GRAFICO_FATCLI + 2
        'Ou seja, se existem mais de NUMERO_CLIENTES_GRAFICO_FATCLI, os clientes com iIndice maior igual a NUMERO_CLIENTES_GRAFICO_FATCLI + 2, fazem parte da linha "Outros"
        'Isso significa que o percentual ser� acrescentado apenas � linha "Outros" pois n�o existe uma linha para cada cliente
        If iCliente > NUMERO_CLIENTES_GRAFICO_FATCLI And iIndice >= NUMERO_CLIENTES_GRAFICO_FATCLI + 2 Then
        
            'Seta o obj que cont�m o nome do Cliente como sendo o obj referente � linha "Outros"
            Set objCelClientes = colLinhasClientes(NUMERO_CLIENTES_GRAFICO_FATCLI + 2)
            
            'Seta o obj que cont�m o valor de faturamento referente � linha "Outros"
            Set objCelFaturamentoCliente = colLinhasFaturamentoCliente(NUMERO_CLIENTES_GRAFICO_FATCLI + 2)
            
            'Calcula o percentual que o faturamento para esse cliente representa dentro do
            'faturamento total da empresa
            dPercent = CDbl(objCelFaturamentoCliente.vValor) / dTotal
            
            'Inclui o percentual ao nome do cliente "Outros"
            objCelClientes.vValor = Format(dPercent, "Percent") & " " & objCelClientes.vValor
            
            'Sai do Loop, pois todos as linhas j� foram alteradas
            Exit For
        
        'Sen�o
        Else
        
            'Seta o obj que cont�m o nome do Cliente iIndice
            Set objCelClientes = colLinhasClientes(iIndice)
            
            'Seta o obj que cont�m o valor faturado para o Cliente iIndice
            Set objCelFaturamentoCliente = colLinhasFaturamentoCliente(iIndice)
            
            'Calcula o percentual que o faturamento para esse cliente representa dentro do
            'faturamento total da empresa
             dPercent = CDbl(objCelFaturamentoCliente.vValor) / dTotal
            
            'Inclui o percentual ao nome do cliente iIndice
            objCelClientes.vValor = Format(dPercent, "Percent") & " " & objCelClientes.vValor
        
        End If
    
    Next
    
    ' *** COLUNA CLIENTES ***
    'Eixo do gr�fico, tipo do gr�fico, exibi��o de labels e cole�ao de c�lulas
    'Informa ao excel em qual eixo essa coluna far� parte do gr�fico
    objColunaClientes.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
    
    'Informa ao excel o tipo de gr�fico que ser� usado para representar essa coluna
    objColunaClientes.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
    
    'Informa ao excel como ser�o exibidos os Datalabels para essa coluna
    objColunaClientes.lDataLabels = EXCEL_EXIBE_LABELS_VALOR
    
    'Transfere para a cole��o de c�lulas da planilha a cole��o de linhas da coluna
    Set objColunaClientes.colCelulas = colLinhasClientes
    '***********************
    
    ' *** COLUNA FATURAMENTO CLIENTE ***
    'Eixo do gr�fico, tipo do gr�fico, exibi��o de labels e cole�ao de c�lulas
    'Informa ao excel em qual eixo essa coluna far� parte do gr�fico
    objColunaFaturamentoCliente.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Informa ao excel o tipo de gr�fico que ser� usado para representar essa coluna
    objColunaFaturamentoCliente.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
    
    'Informa ao excel como ser�o exibidos os Datalabels para essa coluna
    objColunaFaturamentoCliente.lDataLabels = EXCEL_EXIBE_LABELS_VALOR
    
    'Transfere para a cole��o de c�lulas da planilha a cole��o de linhas da coluna
    Set objColunaFaturamentoCliente.colCelulas = colLinhasFaturamentoCliente
    '***********************

    'Transfere para o obj com os dados do gr�fico os obj�s que cont�m os dados de cada coluna
    objPlanilha.colColunas.Add objColunaClientes
    objPlanilha.colColunas.Add objColunaFaturamentoCliente
    
    'Informa ao excel o t�tulo do gr�fico
    objPlanilha.sTituloGrafico = "Per�odo: " & CStr(dtDataIni) & " at� " & CStr(dtDataFim) & vbCrLf & "Faturamento Total: " & CStr(Format(dTotal, "Standard"))

    'Informa ao excel que o gr�fico n�o dever� exibir legenda
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_NAO_EXIBE
    
    'Informa ao excel a posi��o dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW

    'Informa ao excel a orienta��o dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD

    'Informa ao excel que a plotagem do dados ser� por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS

    'Instancia a cole��o que guardar� as se��es de cabe�alho / rodap�
    Set objPlanilha.colCabecalhoRodape = New Collection
    
    'Monta o cabe�alho e o rodap� do Gr�fico
    lErro = Excel_Monta_Cabecalho_Rodape(objPlanilha.colCabecalhoRodape, dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 90537
    
    'Monta a planilha e o gr�fico com os dados passados em objPlanilha
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 79972

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Gera_Grafico = SUCESSO

    Exit Function

Erro_Gera_Grafico:

    Gera_Grafico = gErr

    Select Case gErr

        Case 79972, 177597
        
        Case 90022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90023, 90024, 90025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FATURAMENTOCLIENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161691)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Excel_Monta_Cabecalho_Rodape(colLinhas As Collection, dtDataInicial As Date, dtDataFinal As Date) As Long
'Fun��o criada em 29/05/01 por Luiz Gustavo de Freitas Nogueira
'Essa fun��o preenche os objetos com os dados de cada linha que ser� exibida no cabe�alho
'dtDataInicial, dtDataFinal e sCategoria s�o par�metros que ser�o utilizados para preencher as linhas

Dim objLinha As ClassLinhaCabecalhoExcel

On Error GoTo Erro_Excel_Monta_Cabecalho_Rodape

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
            objLinha.sTexto = "Faturamento por Cliente"
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

    Excel_Monta_Cabecalho_Rodape = SUCESSO
    
    Exit Function
    
Erro_Excel_Monta_Cabecalho_Rodape:

    Excel_Monta_Cabecalho_Rodape = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161692)
            
    End Select
    
    Exit Function

End Function

