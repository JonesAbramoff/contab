VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FluxoRecebCliOcx 
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   ScaleHeight     =   4350
   ScaleWidth      =   9195
   Begin VB.CommandButton BotaoDataDown 
      Height          =   150
      Left            =   2070
      Picture         =   "FluxoRecebCliOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   405
      Width           =   240
   End
   Begin VB.CommandButton BotaoDataUp 
      Height          =   150
      Left            =   2070
      Picture         =   "FluxoRecebCliOcx.ctx":005A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   255
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7425
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "FluxoRecebCliOcx.ctx":00B4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoRecebCliOcx.ctx":0232
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoRecebCliOcx.ctx":0764
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox Usuario 
      Enabled         =   0   'False
      Height          =   210
      Left            =   810
      TabIndex        =   3
      Top             =   3870
      Width           =   870
   End
   Begin VB.ListBox ListaClientes 
      Height          =   3180
      Left            =   6900
      TabIndex        =   9
      Top             =   975
      Width           =   2220
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
      Left            =   2880
      Picture         =   "FluxoRecebCliOcx.ctx":08BE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   1350
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
      Left            =   4650
      Picture         =   "FluxoRecebCliOcx.ctx":0C0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   1350
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   210
      Left            =   1695
      TabIndex        =   4
      Top             =   3855
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ValorAjustado 
      Height          =   225
      Left            =   4410
      TabIndex        =   6
      Top             =   3855
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
      Top             =   975
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
      Left            =   3240
      TabIndex        =   5
      Top             =   3855
      Width           =   1170
      _ExtentX        =   2064
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
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clientes"
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
      Left            =   6900
      TabIndex        =   16
      Top             =   765
      Width           =   690
   End
   Begin VB.Label TotalAjustado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4380
      TabIndex        =   17
      Top             =   4125
      Width           =   1065
   End
   Begin VB.Label TotalReal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5505
      TabIndex        =   18
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label TotalSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3255
      TabIndex        =   19
      Top             =   4170
      Width           =   1065
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
      Left            =   2445
      TabIndex        =   20
      Top             =   4155
      Width           =   600
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
      TabIndex        =   21
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "FluxoRecebCliOcx"
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
Const GRID_CLIENTE_COL = 2
Const GRID_VALORSISTEMA_COL = 3
Const GRID_VALORAJUSTADO_COL = 4
Const GRID_VALORREAL_COL = 5

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long
Dim colFluxoCli As New Collection

On Error GoTo Erro_ExibeFluxo_Click

    'se a data da tela não estiver preenchida ==> não exibe os dados no grid
    If Len(Data.ClipText) = 0 Then Error 21000

    'le os recebimentos consolidados por cliente
    lErro = CF("FluxoForn_Le", colFluxoCli, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_RECEBTO)
    If lErro <> SUCESSO And lErro <> 20205 Then Error 21001
    
    If lErro = 20205 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOCLI_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 55925

    'preenche o grid com os recebimentos lidos
    lErro = Preenche_GridFCaixa(colFluxoCli)
    If lErro <> SUCESSO Then Error 21002
    
    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub

Erro_ExibeFluxo_Click:

    Select Case Err

        Case 21000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 21001, 21002, 55925
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160420)

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
    If lErro <> SUCESSO Then Error 21003
    
    Call Limpa_Tela_FluxoPagCli
    
    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 21003
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160421)

    End Select

    Exit Sub
    
End Sub

Function Gravar_Registro() As Long
'grava os registros da tela

Dim lErro As Long
Dim colFluxoCli As New Collection
Dim objFluxoCli As ClassFluxoForn
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se a data da tela não estiver preenchida ==> erro
    If Len(Data.ClipText) = 0 Then Error 21004

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoCli = New ClassFluxoForn

        objFluxoCli.iUsuario = CInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoCli.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_CLIENTE_COL)
        objFluxoCli.dTotalAjustado = CDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))

        colFluxoCli.Add objFluxoCli

    Next

    lErro = CF("FluxoForn_Grava", colFluxoCli, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_RECEBTO)
    If lErro <> SUCESSO Then Error 21005

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 21005

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160422)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21006

    Call Limpa_Tela_FluxoPagCli

    iAlterado = 0
    iVAAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21006

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160423)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_FluxoPagCli()

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
Dim colClientes As New AdmCollCodigoNome
Dim objCliente As AdmlCodigoNome

On Error GoTo Erro_Form_Load

    'Le cada Codigo e NomeReduzido da tabela de Clientes e poe na colecao
    lErro = CF("LCod_Nomes_Le", "Clientes", "Codigo", "NomeReduzido", STRING_CLIENTE_NOME_REDUZIDO, colClientes)
    If lErro <> SUCESSO Then Error 21007

    'preenche a ListBox ListaClientes com os objetos da colecao
    For Each objCliente In colClientes
        ListaClientes.AddItem objCliente.sNome
    Next

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21008

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
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21007, 21008
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160424)

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
    If (objFluxo Is Nothing) Then gError 21009

    Set objFluxo1 = objFluxo

    'le os pagamentos selecionados
    lErro = CF("FluxoForn_Le", colFluxoAnalitico, objFluxo1.lFluxoId, objFluxo1.dtData, FLUXOANALITICO_TIPOREG_RECEBTO)
    If lErro <> SUCESSO And lErro <> 20205 Then gError 133485
    
    If colFluxoAnalitico.Count = 0 Then
    
        dtData = objFluxo1.dtData

        'le os recebimentos selecionados
        lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 1)
        If lErro <> SUCESSO And lErro <> 133475 Then gError 133486

        If lErro = SUCESSO Then objFluxo1.dtData = dtData
    
    End If

    Data.Text = Format(objFluxo1.dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Recebimentos por Cliente"

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 21009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", gErr)

        Case 133485, 133486

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160425)

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
    objGrid1.colColuna.Add ("Cliente")
    objGrid1.colColuna.Add ("Valor Sistema")
    objGrid1.colColuna.Add ("Valor Ajustado")
    objGrid1.colColuna.Add ("Valor Real")

   'campos de edição do grid
    objGrid1.colCampo.Add (Usuario.Name)
    objGrid1.colCampo.Add (Cliente.Name)
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

Function Preenche_GridFCaixa(colFluxoCli As Collection) As Long
'preenche o grid com os recebimentos contidos na coleção colFluxoCli

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoCli As ClassFluxoForn
Dim dColunaSomaSistema As Double
Dim dColunaSomaAjustado As Double
Dim dColunaSomaReal As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoCli.Count < objGrid1.iLinhasVisiveis Then
        GridFCaixa.Rows = 100
    Else
        GridFCaixa.Rows = colFluxoCli.Count + 100
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoCli.Count

    dColunaSomaSistema = 0
    dColunaSomaAjustado = 0
    dColunaSomaReal = 0

    'preenche o grid com os dados retornados na coleção colFluxoCli
    For iIndice = 1 To colFluxoCli.Count

        Set objFluxoCli = colFluxoCli.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = CStr(objFluxoCli.iUsuario)
        GridFCaixa.TextMatrix(iIndice, GRID_CLIENTE_COL) = objFluxoCli.sNomeReduzido
        GridFCaixa.TextMatrix(iIndice, GRID_VALORSISTEMA_COL) = Format(objFluxoCli.dTotalSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORAJUSTADO_COL) = Format(objFluxoCli.dTotalAjustado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORREAL_COL) = Format(objFluxoCli.dTotalReal, "Standard")
        dColunaSomaSistema = dColunaSomaSistema + objFluxoCli.dTotalSistema
        dColunaSomaAjustado = dColunaSomaAjustado + objFluxoCli.dTotalAjustado
        dColunaSomaReal = dColunaSomaReal + objFluxoCli.dTotalReal


    Next

    For iIndice = colFluxoCli.Count + 1 To GridFCaixa.Rows - 1
        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO
    Next

    Call Grid_Refresh_Checkbox(objGrid1)

    TotalSistema.Caption = Format(dColunaSomaSistema, "Standard")
    TotalAjustado.Caption = Format(dColunaSomaAjustado, "Standard")
    TotalReal.Caption = Format(dColunaSomaReal, "Standard")

    Preenche_GridFCaixa = SUCESSO

    iAlterado = 0
    iVAAlterado = 0

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160426)

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
        If lErro <> SUCESSO Then Error 21011

        dtData = CDate(Data.Text)

        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 21012

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 21011

        Case 21012
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160427)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
        
End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub GridFCaixa_DblClick()

    If GridFCaixa.Row > 0 And Len(Data.ClipText) > 0 Then

        objFluxo1.dtData = StrParaDate(Data.Text)
    
        Call Chama_Tela("FluxoReceb", objFluxo1)
    
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

Private Sub ListaClientes_DblClick()

Dim iLinha As Integer

On Error GoTo Erro_ListaClientes_DblClick

    If GridFCaixa.Col = GRID_CLIENTE_COL And GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then

        'verifica se há alguma linha que já exibe esta conta
        For iLinha = 1 To objGrid1.iLinhasExistentes
        
            If GridFCaixa.TextMatrix(iLinha, GRID_CLIENTE_COL) = ListaClientes.Text Then Error 55916

        Next
    
        GridFCaixa.TextMatrix(GridFCaixa.Row, GridFCaixa.Col) = ListaClientes.Text
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL) = Format(0, "Standard")
    
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

    End If
    
    Exit Sub
    
Erro_ListaClientes_DblClick:

    Select Case Err
    
        Case 55916
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REPETIDO", Err, GridFCaixa.TextMatrix(iLinha, GRID_CLIENTE_COL), iLinha)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160428)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDataDown_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataDown_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 55937

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 0)
    If lErro <> SUCESSO And lErro <> 133475 Then gError 133481

    If lErro = 133475 Then gError 133482

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataDown_Click:

    Select Case gErr

        Case 55937, 133481

        Case 133482
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_AQUEM_DESTA_DATA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160429)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataUp_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataUp_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 55938

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 1)
    If lErro <> SUCESSO And lErro <> 133475 Then gError 133483

    If lErro = 133475 Then gError 133484

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataUp_Click:

    Select Case gErr

        Case 55938, 133483

        Case 133484
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_ALEM_DESTA_DATA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160430)

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

    If iExecutaEntradaCelula = 1 Then Call Grid_Entrada_Celula(objGrid1, iAlterado)

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

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridFCaixa.Col

            Case GRID_CLIENTE_COL

                lErro = Saida_Celula_Cliente(objGridInt)
                If lErro <> SUCESSO Then Error 21015

            Case GRID_VALORAJUSTADO_COL

                lErro = Saida_Celula_ValorAjustado(objGridInt)
                If lErro <> SUCESSO Then Error 21016

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 21017

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21015, 21016

        Case 21017
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160431)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cliente(objGridInt As AdmGrid) As Long
'faz a critica da célula cliente do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente
Dim iFilial As Integer

On Error GoTo Erro_Saida_Celula_Cliente


    Set objGridInt.objControle = Cliente

    'se o cliente foi preenchido
    If Len(Cliente.ClipText) > 0 Then

        'Verifica se o cliente está cadastrado
        lErro = TP_Cliente_Le(Cliente, objCliente, iFilial, 0)
        If lErro <> SUCESSO And lErro <> 6668 And lErro <> 6676 And lErro <> 6701 And lErro <> 6704 Then Error 21021

        'Cliente não cadastrado
        If lErro = 6668 Or lErro = 6676 Or lErro = 6701 Or lErro = 6704 Then Error 21022
        
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If Cliente.Text = GridFCaixa.TextMatrix(iIndice, GRID_CLIENTE_COL) And iIndice <> GridFCaixa.Row Then Error 21221
        Next

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21023

    Saida_Celula_Cliente = SUCESSO

    Exit Function

Erro_Saida_Celula_Cliente:

    Saida_Celula_Cliente = Err

    Select Case Err

        Case 21021, 21023
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21022
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_1", Cliente.Text)

            If vbMsgRes = vbYes Then

                objCliente.sNomeReduzido = Cliente.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Clientes", objCliente)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 21221
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REPETIDO", Cliente.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160432)

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
        If lErro <> SUCESSO Then Error 21024

        If GridFCaixa.Row - GridFCaixa.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_CLIENTE_COL)) > 0 Then ValorAjustado.Text = "0"

    End If

    If Format(ValorAjustado.Text, "Standard") <> GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) Then
        iVAAlterado = REGISTRO_ALTERADO
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21025

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

        Case 21024
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21025
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160433)

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
    If objControl.Name = "Cliente" Then
    
        'se for uma linha criada pelo usuario ==> habilita o campo
        If GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    ElseIf objControl.Name = "ValorAjustado" Then
    
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_CLIENTE_COL)) > 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
            
    End If

End Sub

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoCli As New ClassFluxoForn
Dim colFluxoCli As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXORECEBCLI_CPR")
    If lErro <> SUCESSO Then Error 47914
    
    'obter dados comuns a todas as linhas do grid
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoCli)
    If lErro <> SUCESSO Then Error 47915
    
    For iIndice1 = 1 To colFluxoCli.Count
    
        Set objFluxoCli = colFluxoCli.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoCli.iUsuario)
        Call colTemp.Add(objFluxoCli.sNomeReduzido)
        Call colTemp.Add(objFluxoCli.dTotalSistema)
        Call colTemp.Add(objFluxoCli.dTotalAjustado)
        Call colTemp.Add(objFluxoCli.dTotalReal)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47916
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47917
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47914, 47915, 47916, 47917
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160434)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoCli As Collection) As Long

Dim objFluxoCli As ClassFluxoForn
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoCli = New ClassFluxoForn
        
        objFluxoCli.iUsuario = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoCli.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_CLIENTE_COL)
        objFluxoCli.dTotalSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORSISTEMA_COL))
        objFluxoCli.dTotalAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        objFluxoCli.dTotalReal = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORREAL_COL))
        
        colFluxoCli.Add objFluxoCli
        
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
        Select Case Err
            
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160435)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_RECEBIMENTO_TIPO_CLIENTE
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Recebimentos por Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoRecebCli"
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub TotalAjustado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalAjustado, Source, X, Y)
End Sub

Private Sub TotalAjustado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalAjustado, Button, Shift, X, Y)
End Sub

Private Sub TotalReal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalReal, Source, X, Y)
End Sub

Private Sub TotalReal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalReal, Button, Shift, X, Y)
End Sub

Private Sub TotalSistema_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalSistema, Source, X, Y)
End Sub

Private Sub TotalSistema_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalSistema, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

