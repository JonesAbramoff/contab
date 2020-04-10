VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl GrafFaturamentoCategProdOcx 
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LockControls    =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   5805
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
      Left            =   1320
      Picture         =   "GrafFaturamentoCategProdOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Gerar Gráfico"
      Top             =   2445
      Width           =   1305
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   360
      Left            =   2985
      Picture         =   "GrafFaturamentoCategProdOcx.ctx":0432
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fechar"
      Top             =   2445
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Seleção"
      Height          =   2220
      Left            =   165
      TabIndex        =   5
      Top             =   105
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Intervalo Período"
         Height          =   900
         Left            =   210
         TabIndex        =   8
         Top             =   270
         Width           =   5160
         Begin MSMask.MaskEdBox DataDe 
            Height          =   300
            Left            =   855
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
            Left            =   2010
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   375
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   300
            Left            =   3360
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
            Left            =   4500
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2895
            TabIndex        =   12
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   495
            TabIndex        =   11
            Top             =   420
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Produto"
         Height          =   795
         Left            =   195
         TabIndex        =   6
         Top             =   1245
         Width           =   5160
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   2820
         End
         Begin VB.Label Label5 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   750
            TabIndex        =   7
            Top             =   345
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "GrafFaturamentoCategProdOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    Caption = "Faturamento por Área"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GrafFaturamentoCategProd"
    
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
    If lErro <> SUCESSO Then gError 90058
    
    'Compara com a data Final
    If Len(Trim(DataDe.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 90070

    End If


    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 90058
        
        Case 90070
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161670)

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
    If lErro <> SUCESSO Then gError 90059
    
    'Compara com a data Fianal
    If Len(Trim(DataAte.ClipText)) > 0 Then
        
        dtDataDe = CDate(DataDe.Text)
        dtDataAte = CDate(DataAte.Text)
        
        If dtDataDe > dtDataAte Then gError 90069

    End If


    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 90059
        
        Case 90069
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161671)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro Then gError 90060

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 90060

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161672)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro Then gError 90061

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 90061

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161673)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro Then gError 90062

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 90062

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161674)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro Then gError 90063

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 90063

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161675)

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
Dim sCategoria As String
Dim objGrafico As New ClassGrafico
Dim objItemGrafico As ClassItemGrafico
Dim dTotal As Double
Dim dPercent As Double

On Error GoTo Erro_BotaoGrafico_Click

    'Verifica se a Data Inicial foi digitada
    If Len(Trim(DataDe.ClipText)) = 0 Then gError 90064
    
    'Verifica se a Data Final foi digitada
    If Len(Trim(DataAte.ClipText)) = 0 Then gError 90065
    
    'Verifica se a Categoria foi digitada
    If Len(Trim(CategoriaProduto.Text)) = 0 Then gError 90067
      
    dtDataIni = StrParaDate(DataDe.Text)
    dtDataFim = StrParaDate(DataAte.Text)
    sCategoria = CategoriaProduto.Text
    
    lErro = Grafico_Le_FaturamentoCategoria(objGrafico, giFilialEmpresa, dtDataIni, dtDataFim, sCategoria)
    If lErro <> SUCESSO Then gError 90066
    
    objGrafico.ChartType = 14
    objGrafico.FootNote = ""
    objGrafico.TitleText = ""
    
    For Each objItemGrafico In objGrafico.colcolItensGrafico(1)
        dTotal = dTotal + objItemGrafico.dValorColuna
    Next
    
    For Each objItemGrafico In objGrafico.colcolItensGrafico(1)
        dPercent = objItemGrafico.dValorColuna / dTotal
        objItemGrafico.LegendText = Format(dPercent, "Percent") & "  " & objItemGrafico.sNomeColuna
    Next
        
    'Call Chama_Tela_Nova_Instancia("Grafico", objGrafico)
    Call Gera_Grafico(objGrafico, dtDataIni, dtDataFim, sCategoria)
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGrafico_Click:
    
    Select Case gErr

        Case 90066
        
        Case 90064
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA1", gErr)
           
        Case 90065
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 90067
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161676)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    'Lê as Categorias de Produtos
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 90068

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next
    
    'Seleciona o Default da Combo CategoriaProduto
    lErro = ComboCategoria_Padrao()
    If lErro <> SUCESSO Then gError 91886
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 90068, 91886

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161677)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Function Grafico_Le_FaturamentoCategoria(objGrafico As ClassGrafico, iFilialEmpresa As Integer, dtDataIni As Date, dtDataFim As Date, sCategoria As String) As Long
'Lê o ValorFaturado de cada Item da Categoria

Dim lErro As Long
Dim lComando As Long
Dim objItemGrafico As ClassItemGrafico
Dim dValor As Double
Dim sItem As String
Dim colItensGrafico As New Collection
Dim sFiltro As String

On Error GoTo Erro_Grafico_Le_FaturamentoCategoria

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90039
     
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    If iFilialEmpresa = 0 Then
    
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
    
        'Lê o ValorFaturado de cada Item da Categoria em todas as filiais
        lErro = Comando_Executar(lComando, "SELECT SUM(SldDiaFat.ValorFaturado) AS Valor, CategoriaProdutoItem.Item FROM SldDiaFat,ProdutoCategoria,CategoriaProdutoItem WHERE SldDiaFat.Produto = ProdutoCategoria.Produto AND ProdutoCategoria.Categoria = CategoriaProdutoItem.Categoria AND ProdutoCategoria.Item = CategoriaProdutoItem.Item AND SldDiaFat.Data>=? AND SldDiaFat.Data<=? AND ProdutoCategoria.Categoria = ? " & sFiltro & " GROUP BY CategoriaProdutoItem.Item, CategoriaProdutoItem.Descricao", _
        dValor, sItem, iFilialEmpresa, dtDataIni, dtDataFim, sCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 90040
    
    Else
    
        'Lê o ValorFaturado de cada Item da Categoria
        lErro = Comando_Executar(lComando, "SELECT SUM(SldDiaFat.ValorFaturado) AS Valor, CategoriaProdutoItem.Item FROM SldDiaFat,ProdutoCategoria,CategoriaProdutoItem WHERE SldDiaFat.Produto = ProdutoCategoria.Produto AND ProdutoCategoria.Categoria = CategoriaProdutoItem.Categoria AND ProdutoCategoria.Item = CategoriaProdutoItem.Item AND SldDiaFat.FilialEmpresa=? AND SldDiaFat.Data>=? AND SldDiaFat.Data<=? AND ProdutoCategoria.Categoria = ? GROUP BY CategoriaProdutoItem.Item, CategoriaProdutoItem.Descricao", _
        dValor, sItem, iFilialEmpresa, dtDataIni, dtDataFim, sCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 90040

    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90041

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemGrafico = New ClassItemGrafico
        objItemGrafico.sNomeColuna = sItem
        objItemGrafico.dValorColuna = dValor
        colItensGrafico.Add objItemGrafico
                      
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90042

    Loop
    
    objGrafico.colcolItensGrafico.Add colItensGrafico

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Grafico_Le_FaturamentoCategoria = SUCESSO

    Exit Function

Erro_Grafico_Le_FaturamentoCategoria:

    Grafico_Le_FaturamentoCategoria = gErr

    Select Case gErr

        Case 90039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90040, 90041, 90042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FATURAMENTOCATEGORIA", gErr)

        Case 177597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161678)

    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Gera_Grafico(objGrafico As ClassGrafico, dtDataInicial As Date, dtDataFinal As Date, sCategoria As String) As Long
'Obtém os dados necessários para gerar o gráfico
'Seta configurações do gráfico

Dim iColuna As Integer
Dim iLinha As Integer
Dim lErro As Long
Dim sPercentual As String
Dim sNomeFonte As String
Dim sTextoAux As String
Dim objFilialEmpresa As New AdmFiliais
Dim objPlanilha As New ClassPlanilhaExcel
Dim objLinha As ClassLinhaCabecalhoExcel
Dim objColunasAreas As New ClassColunasExcel
Dim objColunasValores As New ClassColunasExcel
Dim objCelAreas As ClassCelulasExcel
Dim objCelValores As ClassCelulasExcel
Dim objItemGrafico As ClassItemGrafico
Dim colLinhas As New Collection
Dim dValorTotal As Double

On Error GoTo Erro_Gera_Grafico

    MousePointer = vbHourglass
    
    'Informa ao excel o nome do gráfico e o nome da planilha como ajustado
    objPlanilha.sNomeGrafico = "Gráfico - Faturamento por Área"
    objPlanilha.sNomePlanilha = "Faturamento por Área"
    
    'Instancia a coleção que guardará as seções de cabeçalho / rodapé
    Set objPlanilha.colCabecalhoRodape = New Collection

    'Monta o cabeçalho e o rodapé do Gráfico
    '???
    lErro = Excel_Monta_Cabecalho_Rodape(objPlanilha.colCabecalhoRodape, dtDataInicial)
    If lErro <> SUCESSO Then gError 90532

    'Informa se o gráfico possui ou não eixos
    objPlanilha.iEixosGrafico = EXCEL_GRAFICO_SEM_EIXOS
    
    'Seta a posição da legenda do gráfico
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_ABAIXO
    
    ' *** COLUNA ÁREAS ***
    'Informa ao excel em qual eixo essa coluna fará parte do gráfico
    objColunasAreas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
    
    'Informa ao excel o tipo de gráfico que será usado para representar essa coluna
    objColunasAreas.lTipoGraficoColuna = EXCEL_GRAFICO_3DPIE
    
    'Informa ao excel como serão exibidos os Datalabels para essa coluna
    objColunasAreas.lDataLabels = EXCEL_NAO_EXIBE_LABELS
    '***************************

    ' *** COLUNA VALORES ***
    'Informa ao excel em qual eixo essa coluna fará parte do gráfico
    objColunasValores.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Informa ao excel o tipo de gráfico que será usado para representar essa coluna
    objColunasValores.lTipoGraficoColuna = EXCEL_GRAFICO_3DPIE
    
    'Informa ao excel como serão exibidos os Datalabels para essa coluna
    objColunasValores.lDataLabels = EXCEL_EXIBE_LABELS_VALOR
    
    Set objCelAreas = New ClassCelulasExcel
    Set objCelValores = New ClassCelulasExcel
    '************************
    
    objCelAreas.vValor = "Áreas"
    objCelValores.vValor = "Faturamento"
        
    objColunasAreas.colCelulas.Add objCelAreas
    objColunasValores.colCelulas.Add objCelValores
    
    For Each objItemGrafico In objGrafico.colcolItensGrafico(1)
    
        Set objCelAreas = New ClassCelulasExcel
        Set objCelValores = New ClassCelulasExcel
    
        objCelAreas.vValor = objItemGrafico.LegendText
        objCelValores.vValor = Format(objItemGrafico.dValorColuna, "Standard")
        
        dValorTotal = dValorTotal + objItemGrafico.dValorColuna
        
        objColunasAreas.colCelulas.Add objCelAreas
        objColunasValores.colCelulas.Add objCelValores
    
    Next

    'Adiciona as colunas à coleção de colunas
    objPlanilha.colColunas.Add objColunasAreas
    objPlanilha.colColunas.Add objColunasValores
    
    'Informa ao excel o título do gráfico
    objPlanilha.sTituloGrafico = "Período: " & CStr(dtDataInicial) & " até " & CStr(dtDataFinal) & vbCrLf & "Categoria: " & sCategoria & vbCrLf & "Valor Total R$: " & Format(dValorTotal, "STANDARD")
    
    'Informa ao excel a posição dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW
        
    'Informa ao excel a orientação dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD
    
    'Informa ao excel que a plotagem do dados será por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS
    
    'Monta a planilha e o gráfico com os dados passados em objPlanilha
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 79972

    MousePointer = vbDefault
    
    Gera_Grafico = SUCESSO
    
    Exit Function
    
Erro_Gera_Grafico:

    Gera_Grafico = gErr
    
    Select Case gErr
        
        Case 79972, 90531
        
        Case 79971
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORES_COLUNAS_NAO_TRATADOS_GRAFICO", gErr, iColuna)
            
        Case 79970
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAFICO_VALORES_A_EXIBIR_NAO_DEFINIDOS2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161679)
            
    End Select
    
    MousePointer = vbDefault
    
    Exit Function
    
End Function


' *** Fernando, favor subir as funções abaixo. Acho que elas devem ir para o AdmFormata ***
Public Function Excel_CabecalhoRodape_Obtem_Fonte(sTipoFonte As String, sNomeFonte As String) As Long
'Criada em 28/05/2001 por Luiz Gustavo de Freitas Nogueira
'Essa função monta uma string que será passada para o Excel para formatar um texto de cabeçalho / rodapé com a fonte passada como parâmetro

On Error GoTo Erro_Excel_CabecalhoRodape_Obtem_Fonte

    'Seleciona a fonte que será utilizada para formatar o texto
    Select Case sTipoFonte
        
        'Se for Times New Roman
        Case EXCEL_FONTE_TIMES_N_ROMAN
            
            'Monta uma string para formatar o cabeçalho / rodapé com Times New Roman
            sNomeFonte = " &" & Chr$(34) & sTipoFonte & Chr$(34)
        
        'Se for Bookman
        Case EXCEL_FONTE_BOOKMAN
        
            'Monta uma string para formatar o cabeçalho / rodapé com Bookman
            sNomeFonte = " &" & Chr$(34) & sTipoFonte & Chr$(34)
        
        'Se for Courier New
        Case EXCEL_FONTE_COURIER_NEW
            
            'Monta uma string para formatar o cabeçalho / rodapé com Courier New
            sNomeFonte = " &" & Chr$(34) & sTipoFonte & Chr$(34)
        
        'Se for outro tipo de fonte
        Case Else
            
            'Erro, pois não foi implementado o tratamento para a fonte
            gError 90524
    
    End Select
    
    Excel_CabecalhoRodape_Obtem_Fonte = SUCESSO
    
    Exit Function
    
Erro_Excel_CabecalhoRodape_Obtem_Fonte:
    
    Excel_CabecalhoRodape_Obtem_Fonte = gErr
    
    Select Case gErr
    
        Case 90524
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_CABECALHO_RODAPE_FONTE_NAO_IMPLEMENTADA", gErr, sTipoFonte)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161680)
    
    End Select
    
    Exit Function
    
End Function

' *** Retirar as funções abaixo do código ***
Public Function Excel_CabecalhoRodape_MontaTexto(colLinhaTexto As Collection, sTamanhoFonte As String, sTexto As String) As Long
'Criada em 29/05/2001 por Luiz Gustavo de Freitas Nogueira
'Essa função monta um texto que será exibido como cabeçalho / rodapé no Excel
'Ela recebe uma coleção de linhas a serem exibidas e insere as quebras de linha
'Ela também prepara o texto para ser formatado quanto ao tamanho da fonte

Dim iIndiceTexto As Integer

On Error GoTo Erro_Excel_CabecalhoRodape_MontaTexto
    
    'Se o conteúdo passado na variável sTamanhoFonte é um número
    If IsNumeric(sTamanhoFonte) = True Then
        
        'Verifica se ele tem mais do que 2 algarismos. Se tiver => erro
        If Len(Trim(sTamanhoFonte)) >= 3 Then gError 90530
        
        'Verifica se ele tem apenas 1 algarismo
        If Len(Trim(sTamanhoFonte)) = 1 Then
        
            'Se tiver, acrescenta um zero à esquerda (para atender aos padrões do Excel)
            sTamanhoFonte = "0" & sTamanhoFonte
        
        End If
    
    'Se não for numérico
    Else
    
        'Erro
        gError 90529
    
    End If
    
    'Para cada linha de texto dentro da coleção
    For iIndiceTexto = 1 To colLinhaTexto.Count
        
        'Aplica o tamanho da fonte e insere a quebra de linha
        sTexto = sTexto & "&" & sTamanhoFonte & "  " & colLinhaTexto(iIndiceTexto) & vbCrLf & " &01" & vbCrLf
    
    Next
        
    Excel_CabecalhoRodape_MontaTexto = SUCESSO
    
    Exit Function
    
Erro_Excel_CabecalhoRodape_MontaTexto:

    Excel_CabecalhoRodape_MontaTexto = gErr
    
    Select Case gErr
    
        Case 90529
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_CABECALHO_RODAPE_TAMANHOFONTE_NUMERICO", gErr)
        
        Case 90530
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_CABECALHO_RODAPE_TAMANHOFONTE_ALGARISMO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161681)
        
    End Select
    
    Exit Function
    
End Function
' *************

Public Function Excel_Monta_Cabecalho_Rodape(colLinhas As Collection, dtDataInicial As Date) As Long
'Função criada em 29/05/01 por Luiz Gustavo de Freitas Nogueira
'Essa função preenche os objetos com os dados de cada linha que será exibida no cabeçalho
'dtDataInicial, dtDataFinal e sCategoria são parâmetros que serão utilizados para preencher as linhas

Dim objLinha As ClassLinhaCabecalhoExcel

On Error GoTo Erro_Excel_Monta_Cabecalho_Rodape

        '*** PREENCHIMENTO DO CABEÇALHO ESQUERDO ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
        
            ' *** LINHA 01 - CABEÇALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 1 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 1 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 02 - CABEÇALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeFilialEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 2 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 2 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha
            
        ' *** FIM DO CABEÇALHO ESQUERDO ***
        
        ' *** PREENCHIMENTO DO CABEÇALHO CENTRAL ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABEÇALHO CENTRAL ***
            objLinha.iSecao = EXCEL_CABECALHO_CENTRAL
            objLinha.sTexto = "Faturamento por Área"
            objLinha.sFonte = EXCEL_FONTE_BOOKMAN
            objLinha.iTamanhoFonte = 20
            objLinha.sNegrito = EXCEL_CABECALHO_RODAPE_NEGRITO
            objLinha.iEspacoLinha = 0 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 3 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_CENTRAL
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha

        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
        
        ' *** FIM DO CABEÇALHO CENTRAL ***
        
        ' *** PREENCHIMENTO DO CABEÇALHO DIREITO ***
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABEÇALHO DIREITO ***
            objLinha.iSecao = EXCEL_CABECALHO_DIREITO
            objLinha.sTexto = CStr(Date)
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = EXCEL_CABECALHO_RODAPE_NAO_QUEBRA_LINHA 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 1 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_DIREITA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha

        ' *** FIM DO CABEÇALHO DIREITO ***

    Excel_Monta_Cabecalho_Rodape = SUCESSO
    
    Exit Function
    
Erro_Excel_Monta_Cabecalho_Rodape:

    Excel_Monta_Cabecalho_Rodape = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161682)
            
    End Select
    
    Exit Function

End Function

Function ComboCategoria_Padrao() As Long
'Seleciona na Combo CategoriaProduto, o ítem "serviço" como padrão

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_ComboCategoria_Padrao

    'Verifica se a combo está preenchida
    If CategoriaProduto.ListCount = 0 Then Exit Function
    
    'Seleciona o ítem serviço na combo
    For iIndice = 0 To CategoriaProduto.ListCount
        If CategoriaProduto.List(iIndice) = "serviços" Then
            CategoriaProduto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    ComboCategoria_Padrao = SUCESSO
    
    Exit Function
    
Erro_ComboCategoria_Padrao:

    ComboCategoria_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161683)
            
    End Select
    
    Exit Function

End Function
