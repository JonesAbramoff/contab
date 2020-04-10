VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FluxoAplicOcx 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ScaleHeight     =   4695
   ScaleWidth      =   6015
   Begin VB.PictureBox Picture6 
      Height          =   555
      Left            =   4710
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoAplicOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoAplicOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton Botao_ImprimeFluxo 
      Caption         =   "Imprime Fluxo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3300
      Picture         =   "FluxoAplicOcx.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton Botao_ExibeFluxo 
      Caption         =   "Exibe Fluxo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1920
      Picture         =   "FluxoAplicOcx.ctx":07B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   1590
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GridFCaixa 
      Height          =   3105
      Left            =   120
      TabIndex        =   6
      Top             =   825
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   5477
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin MSMask.MaskEdBox DataResgatePrevista 
      Height          =   300
      Left            =   570
      TabIndex        =   0
      Top             =   195
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CodAplicacao 
      Height          =   210
      Left            =   705
      TabIndex        =   3
      Top             =   4080
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorResgatePrevisto 
      Height          =   240
      Left            =   3630
      TabIndex        =   5
      Top             =   4065
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SaldoAplicado 
      Height          =   225
      Left            =   2115
      TabIndex        =   4
      Top             =   4065
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   75
      TabIndex        =   11
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LabelTotais 
      AutoSize        =   -1  'True
      Caption         =   "Totais:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1845
      TabIndex        =   12
      Top             =   4410
      Width           =   600
   End
   Begin VB.Label TotalValor 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   4260
      TabIndex        =   13
      Top             =   4245
      Width           =   1290
   End
   Begin VB.Label TotalSaldo 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   2925
      TabIndex        =   14
      Top             =   4245
      Width           =   1290
   End
End
Attribute VB_Name = "FluxoAplicOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGrid1 As AdmGrid
Dim lFluxoId As Long
Dim objFluxo1 As ClassFluxo

'Colunas do Grid
Const GRID_CODIGO_APLICACAO_COL = 1
Const GRID_SALDO_APLICADO_COL = 2
Const GRID_VALOR_RESGATE_COL = 3

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long

On Error GoTo Erro_Botao_ExibeFluxo_Click

    'se a data da tela não estiver preenchido ==> erro
    If Len(DataResgatePrevista.ClipText) = 0 Then Error 21161

    lErro = Ordenados()
    If lErro <> SUCESSO Then Error 55920

    Exit Sub

Erro_Botao_ExibeFluxo_Click:

    Select Case Err

        Case 21161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 55920

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160354)

    End Select

    Exit Sub

End Sub

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoAplic As New ClassFluxoAplic
Dim colFluxoAplic As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOAPLIC_CPR")
    If lErro <> SUCESSO Then Error 47898
    
    'obter dados comuns a todas as linhas do grid
    dtData = StrParaDate(DataResgatePrevista.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoAplic)
    If lErro <> SUCESSO Then Error 47899
    
    For iIndice1 = 1 To colFluxoAplic.Count
    
        Set objFluxoAplic = colFluxoAplic.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoAplic.lCodigo)
        Call colTemp.Add(objFluxoAplic.dSaldoAplicado)
        Call colTemp.Add(objFluxoAplic.dValorResgatePrevisto)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47900
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47901
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47898, 47899, 47900, 47901
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160355)
     
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_FluxoAplic

End Sub

Sub Limpa_Tela_FluxoAplic()

    Call Grid_Limpa(objGrid1)
    DataResgatePrevista.PromptInclude = False
    DataResgatePrevista.Text = ""
    DataResgatePrevista.PromptInclude = True
    TotalSaldo.Caption = ""
    TotalValor.Caption = ""

End Sub

Private Sub DataResgatePrevista_Change()

    If objGrid1.iLinhasExistentes > 0 Then
        Call Grid_Limpa(objGrid1)
    End If
    
    TotalSaldo.Caption = ""
    TotalValor.Caption = ""

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21162
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21162

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160356)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then Error 21163

    Set objFluxo1 = objFluxo

    DataResgatePrevista.Text = Format(objFluxo.dtDataBase, "dd/mm/yy")

    lErro = Ordenados()
    If lErro <> SUCESSO Then Error 55921

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Resgates"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 21163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 55921

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160357)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Inicializa_GridFCaixa() As Long

Dim iIndice As Integer

    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Código")
    objGrid1.colColuna.Add ("Saldo Aplicado")
    objGrid1.colColuna.Add ("Valor Previsto de Resgate")

   'campos de edição do grid
    objGrid1.colCampo.Add (CodAplicacao.Name)
    objGrid1.colCampo.Add (SaldoAplicado.Name)
    objGrid1.colCampo.Add (ValorResgatePrevisto.Name)

    objGrid1.objGrid = GridFCaixa

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 10

    objGrid1.objGrid.ColWidth(0) = 300

    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid1)

    'Posiciona o totalizador
    TotalSaldo.Top = GridFCaixa.Top + GridFCaixa.Height
    TotalSaldo.Left = GridFCaixa.Left

    For iIndice = 0 To GRID_SALDO_APLICADO_COL - 1
        TotalSaldo.Left = TotalSaldo.Left + GridFCaixa.ColWidth(iIndice) + GridFCaixa.GridLineWidth + 10
    Next

    TotalSaldo.Width = GridFCaixa.ColWidth(GRID_SALDO_APLICADO_COL)

    TotalValor.Top = TotalSaldo.Top
    TotalValor.Left = TotalSaldo.Left + TotalSaldo.Width + GridFCaixa.GridLineWidth
    TotalValor.Width = GridFCaixa.ColWidth(GRID_VALOR_RESGATE_COL)

    LabelTotais.Top = TotalSaldo.Top + (TotalSaldo.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalSaldo.Left - LabelTotais.Width - 50

    Inicializa_GridFCaixa = SUCESSO

End Function

Function Preenche_GridFCaixa(colFluxoAplic As Collection) As Long
'preenche o grid com as aplicações contidas na coleção colFluxoAplic

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoAplic As ClassFluxoAplic
Dim dColunaSomaValor As Double
Dim dColunaSomaSaldo As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoAplic.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colFluxoAplic.Count + 1
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoAplic.Count

    dColunaSomaSaldo = 0
    dColunaSomaValor = 0

    'preenche o grid com os dados retornados na coleção colFluxoAplic
    For iIndice = 1 To colFluxoAplic.Count

        Set objFluxoAplic = colFluxoAplic.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_APLICACAO_COL) = objFluxoAplic.lCodigo
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_APLICADO_COL) = Format(objFluxoAplic.dSaldoAplicado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_RESGATE_COL) = Format(objFluxoAplic.dValorResgatePrevisto, "Standard")
        dColunaSomaValor = dColunaSomaValor + objFluxoAplic.dValorResgatePrevisto
        dColunaSomaSaldo = dColunaSomaSaldo + objFluxoAplic.dSaldoAplicado

    Next

    TotalSaldo.Caption = Format(dColunaSomaSaldo, "Standard")
    TotalValor.Caption = Format(dColunaSomaValor, "Standard")

    Preenche_GridFCaixa = SUCESSO

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160358)

    End Select

    Exit Function

End Function

Private Sub DataResgatePrevista_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataResgatePrevista, iAlterado)

End Sub

Private Sub DataResgatePrevista_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_DataResgatePrevista_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataResgatePrevista.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataResgatePrevista.Text)
        If lErro <> SUCESSO Then Error 21165

        dtData = CDate(DataResgatePrevista.Text)

        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 21166

    End If

    Exit Sub

Erro_DataResgatePrevista_Validate:

    Cancel = True


    Select Case Err

        Case 21165

        Case 21166
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160359)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDown1_DownClick

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 21167

        dtData = CDate(sData)

        If dtData < objFluxo1.dtDataBase Then
            DataResgatePrevista.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
        ElseIf dtData > objFluxo1.dtDataFinal Then
            DataResgatePrevista.Text = Format(objFluxo1.dtDataFinal, "dd/mm/yy")
        Else
            DataResgatePrevista.Text = sData
        End If

    Else
        DataResgatePrevista.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
    End If

    Call Botao_ExibeFluxo_Click
    
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 21167

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160360)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDown1_UpClick

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 21168

        dtData = CDate(sData)

        If dtData < objFluxo1.dtDataBase Then
            DataResgatePrevista.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
        ElseIf dtData > objFluxo1.dtDataFinal Then
            DataResgatePrevista.Text = Format(objFluxo1.dtDataFinal, "dd/mm/yy")
        Else
            DataResgatePrevista.Text = sData
        End If

    Else
        DataResgatePrevista.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
    End If

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 21168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160361)

    End Select

    Exit Sub

End Sub

Private Sub GridFCaixa_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridFCaixa_GotFocus()
    Call Grid_Recebe_Foco(objGrid1)
End Sub

Private Sub GridFCaixa_EnterCell()
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
End Sub

Private Sub GridFCaixa_LeaveCell()
    Call Saida_Celula(objGrid1)
End Sub

Private Sub GridFCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
End Sub

Private Sub GridFCaixa_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridFCaixa_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid1)
End Sub

Private Sub GridFCaixa_RowColChange()
    Call Grid_RowColChange(objGrid1)
End Sub

Private Sub GridFCaixa_Scroll()
    Call Grid_Scroll(objGrid1)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 21169

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21169
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160362)

    End Select

    Exit Function

End Function

Private Function Ordenados() As Long

Dim sOrdenacao As String
Dim lErro As Long
Dim colFluxoAplic As New Collection

On Error GoTo Erro_Ordenados

    'se a data da tela não estiver preenchido ==> não exibe os dados no grid
    If Len(DataResgatePrevista.ClipText) <> 0 Then

        'monta a expressão SQL de Ordenação
        sOrdenacao = " ORDER BY Codigo"
        
        'le as aplicações selecionadas
        lErro = CF("FluxoAplic_Le",colFluxoAplic, sOrdenacao, objFluxo1.lFluxoId, CDate(DataResgatePrevista.Text))
        If lErro <> SUCESSO And lErro <> 21173 Then Error 21175
        
        If lErro = 21173 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOAPLIC_ULTRAPASSOU_LIMITE", Format(DataResgatePrevista.Text, "dd/mm/yyyy"), MAX_FLUXO)
        
        'preenche o grid com as aplicações lidas
        lErro = Preenche_GridFCaixa(colFluxoAplic)
        If lErro <> SUCESSO Then Error 21176

    End If

    Ordenados = SUCESSO

    Exit Function

Erro_Ordenados:

    Ordenados = Err

    Select Case Err

        Case 21175, 21176

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160363)

    End Select

    Exit Function

End Function

Function Grid_FCaixa_Obter(colFluxoAplic As Collection) As Long

Dim objFluxoAplic As ClassFluxoAplic
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoAplic = New ClassFluxoAplic
        
        objFluxoAplic.lCodigo = StrParaLong(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_APLICACAO_COL))
        objFluxoAplic.dSaldoAplicado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_APLICADO_COL))
        objFluxoAplic.dValorResgatePrevisto = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_RESGATE_COL))
        
        colFluxoAplic.Add objFluxoAplic

    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
        Select Case Err
            
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160364)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_RESGATES
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Resgates"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoAplic"
    
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
End Sub

Private Sub Unload(objme As Object)
    
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


Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub TotalValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalValor, Source, X, Y)
End Sub

Private Sub TotalValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalValor, Button, Shift, X, Y)
End Sub

Private Sub TotalSaldo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalSaldo, Source, X, Y)
End Sub

Private Sub TotalSaldo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalSaldo, Button, Shift, X, Y)
End Sub

