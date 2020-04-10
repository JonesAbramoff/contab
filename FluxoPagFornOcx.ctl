VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FluxoPagFornOcx 
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   4350
   ScaleWidth      =   9270
   Begin VB.CommandButton BotaoDataDown 
      Height          =   150
      Left            =   2055
      Picture         =   "FluxoPagFornOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   405
      Width           =   240
   End
   Begin VB.CommandButton BotaoDataUp 
      Height          =   150
      Left            =   2055
      Picture         =   "FluxoPagFornOcx.ctx":005A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   255
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7455
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1140
         Picture         =   "FluxoPagFornOcx.ctx":00B4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "FluxoPagFornOcx.ctx":0232
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoPagFornOcx.ctx":0764
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
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
      Height          =   615
      Left            =   4665
      Picture         =   "FluxoPagFornOcx.ctx":08BE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   1290
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
      Height          =   615
      Left            =   2850
      Picture         =   "FluxoPagFornOcx.ctx":09C0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   1290
   End
   Begin VB.ListBox ListaFornecedores 
      Height          =   3180
      Left            =   6960
      TabIndex        =   9
      Top             =   975
      Width           =   2220
   End
   Begin VB.CheckBox Usuario 
      Enabled         =   0   'False
      Height          =   210
      Left            =   810
      TabIndex        =   3
      Top             =   3900
      Width           =   870
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   210
      Left            =   1605
      TabIndex        =   4
      Top             =   3870
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ValorAjustado 
      Height          =   225
      Left            =   4380
      TabIndex        =   6
      Top             =   3870
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridFCaixa 
      Height          =   2805
      Left            =   120
      TabIndex        =   8
      Top             =   990
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Top             =   255
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorSistema 
      Height          =   225
      Left            =   3210
      TabIndex        =   5
      Top             =   3840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorReal 
      Height          =   225
      Left            =   5580
      TabIndex        =   7
      Top             =   3855
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
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
      Left            =   300
      TabIndex        =   16
      Top             =   285
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
      Left            =   2430
      TabIndex        =   17
      Top             =   4155
      Width           =   600
   End
   Begin VB.Label TotalSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label TotalReal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5505
      TabIndex        =   19
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label TotalAjustado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4365
      TabIndex        =   20
      Top             =   4125
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   21
      Top             =   750
      Width           =   1170
   End
End
Attribute VB_Name = "FluxoPagFornOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Dim objFluxo1 As ClassFluxo
Dim iVAAlterado As Integer

'Colunas do Grid
Const GRID_USUARIO_COL = 1
Const GRID_FORNECEDOR_COL = 2
Const GRID_VALORSISTEMA_COL = 3
Const GRID_VALORAJUSTADO_COL = 4
Const GRID_VALORREAL_COL = 5

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long
Dim colFluxoForn As New Collection

On Error GoTo Erro_ExibeFluxo_Click

    'se a data da tela não estiver preenchida ==> não exibe os dados no grid
    If Len(Data.ClipText) = 0 Then Error 20199

    'le os pagamentos consolidados por fornecedor
    lErro = CF("FluxoForn_Le", colFluxoForn, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_PAGTO)
    If lErro <> SUCESSO And lErro <> 20205 Then Error 20210
    
    If lErro = 20205 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOFORN_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 55922

    'preenche o grid com os pagamentos lidos
    lErro = Preenche_GridFCaixa(colFluxoForn)
    If lErro <> SUCESSO Then Error 20212
    
    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub

Erro_ExibeFluxo_Click:

    Select Case Err

        Case 20199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 20210, 20212, 55922
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160389)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 20237
    
    Call Limpa_Tela_FluxoPagForn
    
    iAlterado = 0
    
    iVAAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 20237
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160390)

    End Select

    Exit Sub
    
End Sub

Function Gravar_Registro() As Long
'Grava os registros da tela

Dim lErro As Long
Dim colFluxoForn As New Collection
Dim objFluxoForn As ClassFluxoForn
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se a data da tela não estiver preenchida ==> erro
    If Len(Data.ClipText) = 0 Then Error 20235
    
    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoForn = New ClassFluxoForn
        
        objFluxoForn.iUsuario = CInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoForn.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_FORNECEDOR_COL)
        objFluxoForn.dTotalAjustado = -CDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        
        colFluxoForn.Add objFluxoForn

    Next
    
    lErro = CF("FluxoForn_Grava", colFluxoForn, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_PAGTO)
    If lErro <> SUCESSO Then Error 20241
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 20235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 20241
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160391)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 20238

    Call Limpa_Tela_FluxoPagForn
    
    iAlterado = 0
    
    iVAAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 20238
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160392)

    End Select

    Exit Sub
        
End Sub

Sub Limpa_Tela_FluxoPagForn()

    Call Limpa_GridFCaixa
    Data.Text = "  /  /  "
    TotalSistema.Caption = ""
    TotalAjustado.Caption = ""
    TotalReal.Caption = ""
    
End Sub

Sub Limpa_GridFCaixa()

Dim iIndice As Integer

    Call Grid_Limpa(objGrid1)
    For iIndice = 1 To GridFCaixa.Rows - 1
        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO
    Next
    Call Grid_Refresh_Checkbox(objGrid1)
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim colFornecedores As New AdmCollCodigoNome
Dim objFornecedor As AdmlCodigoNome

On Error GoTo Erro_Form_Load
    
    'Le cada Codigo e NomeReduzido da tabela de Fornecedores e poe na colecao
    lErro = CF("LCod_Nomes_Le", "Fornecedores", "Codigo", "NomeReduzido", STRING_FORNECEDOR_NOME_REDUZIDO, colFornecedores)
    If lErro <> SUCESSO Then Error 20234

    'preenche a ListBox FornecedoresList com os objetos da colecao
    For Each objFornecedor In colFornecedores
        ListaFornecedores.AddItem objFornecedor.sNome
    Next
    
    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 20213
    
    'Posiciona o totalizador
    TotalSistema.Top = GridFCaixa.Top + GridFCaixa.Height
    TotalSistema.Left = GridFCaixa.Left
    
    For iIndice = 0 To GRID_VALORSISTEMA_COL - 1
        TotalSistema.Left = TotalSistema.Left + GridFCaixa.ColWidth(iIndice) + GridFCaixa.GridLineWidth + 10
    Next
    
    TotalSistema.Width = GridFCaixa.ColWidth(GRID_VALORSISTEMA_COL)
    
    TotalAjustado.Top = TotalSistema.Top
    TotalAjustado.Left = TotalSistema.Left + TotalSistema.Width + GridFCaixa.GridLineWidth
    TotalAjustado.Width = GridFCaixa.ColWidth(GRID_VALORAJUSTADO_COL)
    
    TotalReal.Top = TotalAjustado.Top
    TotalReal.Left = TotalAjustado.Left + TotalAjustado.Width + GridFCaixa.GridLineWidth
    TotalReal.Width = GridFCaixa.ColWidth(GRID_VALORREAL_COL)
    
    LabelTotais.Top = TotalSistema.Top + (TotalSistema.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalSistema.Left - LabelTotais.Width - 50
  
    iAlterado = 0
    iVAAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 20213, 20234
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160393)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colFluxoAnalitico As New Collection
Dim dtData As Date

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then gError 20214

    Set objFluxo1 = objFluxo

    'le os pagamentos selecionados
    lErro = CF("FluxoForn_Le", colFluxoAnalitico, objFluxo1.lFluxoId, objFluxo1.dtData, FLUXOANALITICO_TIPOREG_PAGTO)
    If lErro <> SUCESSO And lErro <> 20205 Then gError 133480
    
    If colFluxoAnalitico.Count = 0 Then
    
        dtData = objFluxo1.dtData

        'le os recebimentos selecionados
        lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 1)
        If lErro <> SUCESSO And lErro <> 133475 Then gError 133481

        If lErro = SUCESSO Then objFluxo1.dtData = dtData
    
    End If

    Data.Text = Format(objFluxo1.dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click
    
    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Pagamentos por Fornecedor"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 20214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 133480, 133481

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160394)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Inicializa_GridFCaixa() As Long
   
Dim iIndice As Integer
   
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Usuário")
    objGrid1.colColuna.Add ("Fornecedor")
    objGrid1.colColuna.Add ("Valor Sistema")
    objGrid1.colColuna.Add ("Valor Ajustado")
    objGrid1.colColuna.Add ("Valor Real")
    
   'campos de edição do grid
    objGrid1.colCampo.Add (Usuario.Name)
    objGrid1.colCampo.Add (Fornecedor.Name)
    objGrid1.colCampo.Add (ValorSistema.Name)
    objGrid1.colCampo.Add (ValorAjustado.Name)
    objGrid1.colCampo.Add (ValorReal.Name)
    
    objGrid1.objGrid = GridFCaixa
   
    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 10
    
    objGrid1.objGrid.ColWidth(0) = 300
    
    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    
    objGrid1.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
        
    Call Grid_Inicializa(objGrid1)
    
    
    Inicializa_GridFCaixa = SUCESSO
    
End Function

Function Preenche_GridFCaixa(colFluxoForn As Collection) As Long
'preenche o grid com os pagamentos contidos na coleção colFluxoForn

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoForn As ClassFluxoForn
Dim dColunaSomaSistema As Double
Dim dColunaSomaAjustado As Double
Dim dColunaSomaReal As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoForn.Count < objGrid1.iLinhasVisiveis Then
        GridFCaixa.Rows = 100
    Else
        GridFCaixa.Rows = colFluxoForn.Count + 100
    End If
    
    Call Grid_Inicializa(objGrid1)
    
    objGrid1.iLinhasExistentes = colFluxoForn.Count

    dColunaSomaSistema = 0
    dColunaSomaAjustado = 0
    dColunaSomaReal = 0
    
    'preenche o grid com os dados retornados na coleção colFluxoForn
    For iIndice = 1 To colFluxoForn.Count

        Set objFluxoForn = colFluxoForn.Item(iIndice)
        
        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = CStr(objFluxoForn.iUsuario)
        GridFCaixa.TextMatrix(iIndice, GRID_FORNECEDOR_COL) = objFluxoForn.sNomeReduzido
        GridFCaixa.TextMatrix(iIndice, GRID_VALORSISTEMA_COL) = Format(-objFluxoForn.dTotalSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORAJUSTADO_COL) = Format(-objFluxoForn.dTotalAjustado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORREAL_COL) = Format(-objFluxoForn.dTotalReal, "Standard")
        dColunaSomaSistema = dColunaSomaSistema + objFluxoForn.dTotalSistema
        dColunaSomaAjustado = dColunaSomaAjustado + objFluxoForn.dTotalAjustado
        dColunaSomaReal = dColunaSomaReal + objFluxoForn.dTotalReal
        
        
    Next

    For iIndice = colFluxoForn.Count + 1 To GridFCaixa.Rows - 1
        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO
    Next

    Call Grid_Refresh_Checkbox(objGrid1)

    TotalSistema.Caption = Format(-dColunaSomaSistema, "Standard")
    TotalAjustado.Caption = Format(-dColunaSomaAjustado, "Standard")
    TotalReal.Caption = Format(-dColunaSomaReal, "Standard")

    Preenche_GridFCaixa = SUCESSO
    
    iAlterado = 0
    iVAAlterado = 0

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160395)

    End Select

    Exit Function

End Function

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Data_Validate

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 20215
        
        dtData = CDate(Data.Text)
        
        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 20216
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 20215

        Case 20216
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160396)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub Fornecedor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Fornecedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Fornecedor
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridFCaixa_DblClick()

    If GridFCaixa.Row > 0 And Len(Data.ClipText) > 0 Then

        objFluxo1.dtData = StrParaDate(Data.Text)
    
        Call Chama_Tela("FluxoPag", objFluxo1)
    
    End If

End Sub

Private Sub ValorAjustado_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorAjustado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub ValorAjustado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorAjustado
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub


Private Sub ListaFornecedores_DblClick()

Dim iLinha As Integer

On Error GoTo Erro_ListaFornecedores_DblClick

    If GridFCaixa.Col = GRID_FORNECEDOR_COL And GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
    
        'verifica se há alguma linha que já exibe esta conta
        For iLinha = 1 To objGrid1.iLinhasExistentes
        
            If GridFCaixa.TextMatrix(iLinha, GRID_FORNECEDOR_COL) = ListaFornecedores.Text Then Error 55915

        Next
    
        GridFCaixa.TextMatrix(GridFCaixa.Row, GridFCaixa.Col) = ListaFornecedores.Text
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL) = Format(0, "Standard")
    
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
    
    End If
    
    Exit Sub
    
Erro_ListaFornecedores_DblClick:

    Select Case Err
    
        Case 55915
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_REPETIDO", Err, GridFCaixa.TextMatrix(iLinha, GRID_FORNECEDOR_COL), iLinha)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160397)
    
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_UpDown1_DownClick

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoAnalitico_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 0)
    If lErro <> SUCESSO And lErro <> 133191 Then gError 133462

    If lErro = 133191 Then gError 133463

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 133462

        Case 133463
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_AQUEM_DESTA_DATA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160398)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_UpDown1_UpClick

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoAnalitico_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 1)
    If lErro <> SUCESSO And lErro <> 133191 Then gError 133460

    If lErro = 133191 Then gError 133461

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 133460

        Case 133461
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_ALEM_DESTA_DATA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160399)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoDataDown_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataDown_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 55935

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 0)
    If lErro <> SUCESSO And lErro <> 133475 Then gError 133476

    If lErro = 133475 Then gError 133477

    Data.Text = Format(dtData, "dd/mm/yy")
    
    Call Botao_ExibeFluxo_Click
    
    Exit Sub
    
Erro_BotaoDataDown_Click:
    
    Select Case gErr
    
        Case 55935, 133476
        
        Case 133477
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_AQUEM_DESTA_DATA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160400)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDataUp_Click()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_BotaoDataUp_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 55936

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 1)
    If lErro <> SUCESSO And lErro <> 133475 Then gError 133478

    If lErro = 133475 Then gError 133479

    Data.Text = Format(dtData, "dd/mm/yy")
    
    Call Botao_ExibeFluxo_Click

    Exit Sub
    
Erro_BotaoDataUp_Click:
    
    Select Case gErr
    
        Case 55936, 133478
        
        Case 133479
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_ALEM_DESTA_DATA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160401)
        
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

Dim dColunaSoma As Double
Dim iLinhasExistentes As Integer

    iLinhasExistentes = objGrid1.iLinhasExistentes
    
    If GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
        Call Grid_Trata_Tecla2(KeyCode, objGrid1)
        
        If iLinhasExistentes <> objGrid1.iLinhasExistentes Then iAlterado = REGISTRO_ALTERADO
        
        dColunaSoma = GridColuna_Soma(GRID_VALORAJUSTADO_COL)
        TotalAjustado.Caption = Format(dColunaSoma, "Standard")
        
    End If
    
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

        Select Case GridFCaixa.Col
        
            Case GRID_FORNECEDOR_COL
            
                    lErro = Saida_Celula_Fornecedor(objGridInt)
                    If lErro <> SUCESSO Then Error 20221
                    
            Case GRID_VALORAJUSTADO_COL
            
                lErro = Saida_Celula_ValorAjustado(objGridInt)
                If lErro <> SUCESSO Then Error 20222

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 20223

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 20221, 20222

        Case 20223
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160402)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fornecedor(objGridInt As AdmGrid) As Long
'faz a critica da celula fornecedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor
Dim iFilial As Integer

On Error GoTo Erro_Saida_Celula_Fornecedor
    
    Set objGridInt.objControle = Fornecedor
    
    'Se o fornecedor foi preenchido
    If Len(Fornecedor.ClipText) > 0 Then
    
        'Verifica se o fornecedor está cadastrado
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iFilial, 0)
        If lErro <> SUCESSO And lErro <> 6663 And lErro <> 6664 And lErro <> 6660 And lErro <> 6675 Then Error 20229
    
        'Fornecedor não cadastrado
        If lErro = 6663 Or lErro = 6664 Or lErro = 6660 Or lErro = 6675 Then Error 20230
            
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If Fornecedor.Text = GridFCaixa.TextMatrix(iIndice, GRID_FORNECEDOR_COL) And iIndice <> GridFCaixa.Row Then Error 21223
        Next

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20231

    Saida_Celula_Fornecedor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = Err
    
    Select Case Err
    
        Case 20229, 20231
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20230
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Fornecedores", objFornecedor)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                    
            End If
            
        Case 21223
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_REPETIDO", Fornecedor.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160403)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorAjustado(objGridInt As AdmGrid) As Long
'faz a critica da celula valor ajustado do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_ValorAjustado

    Set objGridInt.objControle = ValorAjustado
    
    If Len(ValorAjustado.Text) > 0 Then
        
        lErro = Valor_NaoNegativo_Critica(ValorAjustado.Text)
        If lErro <> SUCESSO Then Error 20232
                
        If GridFCaixa.Row - GridFCaixa.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    Else
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_FORNECEDOR_COL)) > 0 Then ValorAjustado.Text = "0"

    End If
        
    If Format(ValorAjustado.Text, "Standard") <> GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) Then
        iVAAlterado = REGISTRO_ALTERADO
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20233
        
    If iVAAlterado = REGISTRO_ALTERADO Then
        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO
        
        iVAAlterado = 0
        
        Call Grid_Refresh_Checkbox(objGridInt)
        
    End If
            
    dColunaSoma = GridColuna_Soma(GRID_VALORAJUSTADO_COL)
    TotalAjustado.Caption = Format(dColunaSoma, "Standard")
   
    Saida_Celula_ValorAjustado = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ValorAjustado:

    Saida_Celula_ValorAjustado = Err
    
    Select Case Err
    
        Case 20233
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20232
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160404)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If Len(GridFCaixa.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(GridFCaixa.TextMatrix(iLinha, iColuna))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Private Sub ValorAjustado_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

    'se estiver trabalhando na campo de fornecedor
    If objControl.Name = "Fornecedor" Then
    
        'se for uma linha criada pelo usuario ==> habilita o campo
        If GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    ElseIf objControl.Name = "ValorAjustado" Then
    
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_FORNECEDOR_COL)) > 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    End If

End Sub

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoForn As New ClassFluxoForn
Dim colFluxoForn As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOPAGFORN_CPR")
    If lErro <> SUCESSO Then Error 47906
    
    'obter dados comuns a todas as linhas do grid
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoForn)
    If lErro <> SUCESSO Then Error 47907
    
    For iIndice1 = 1 To colFluxoForn.Count
    
        Set objFluxoForn = colFluxoForn.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoForn.iUsuario)
        Call colTemp.Add(objFluxoForn.sNomeReduzido)
        Call colTemp.Add(objFluxoForn.dTotalSistema)
        Call colTemp.Add(objFluxoForn.dTotalAjustado)
        Call colTemp.Add(objFluxoForn.dTotalReal)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47908
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47909
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47906, 47907, 47908, 47909
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160405)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoForn As Collection) As Long

Dim objFluxoForn As ClassFluxoForn
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoForn = New ClassFluxoForn
        
        objFluxoForn.iUsuario = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoForn.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_FORNECEDOR_COL)
        objFluxoForn.dTotalSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORSISTEMA_COL))
        objFluxoForn.dTotalAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        objFluxoForn.dTotalReal = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORREAL_COL))
        
        colFluxoForn.Add objFluxoForn
        
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
        Select Case Err
            
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160406)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_PAGAMENTO_FORNECEDOR
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Pagamentos por Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoPagForn"
    
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

Private Sub TotalSistema_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalSistema, Source, X, Y)
End Sub

Private Sub TotalSistema_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalSistema, Button, Shift, X, Y)
End Sub

Private Sub TotalReal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalReal, Source, X, Y)
End Sub

Private Sub TotalReal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalReal, Button, Shift, X, Y)
End Sub

Private Sub TotalAjustado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalAjustado, Source, X, Y)
End Sub

Private Sub TotalAjustado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalAjustado, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

