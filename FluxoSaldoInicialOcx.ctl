VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FluxoSaldoInicialOcx 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   4305
   ScaleWidth      =   9510
   Begin MSMask.MaskEdBox ValorSistema 
      Height          =   225
      Left            =   3120
      TabIndex        =   5
      Top             =   3495
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
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7740
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   180
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "FluxoSaldoInicialOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoSaldoInicialOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoSaldoInicialOcx.ctx":06B0
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
      Left            =   3900
      Picture         =   "FluxoSaldoInicialOcx.ctx":080A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   195
      Width           =   1320
   End
   Begin VB.ListBox ListaCCI 
      Height          =   3180
      Left            =   7440
      TabIndex        =   9
      Top             =   1020
      Width           =   1995
   End
   Begin VB.CheckBox Usuario 
      Enabled         =   0   'False
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   3465
      Width           =   600
   End
   Begin VB.TextBox CCINomeReduzido 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3480
      Width           =   1440
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
      Left            =   1440
      Picture         =   "FluxoSaldoInicialOcx.ctx":090C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   195
      Width           =   1320
   End
   Begin MSMask.MaskEdBox CCICodigo 
      Height          =   225
      Left            =   765
      TabIndex        =   3
      Top             =   3480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorAjustado 
      Height          =   225
      Left            =   4305
      TabIndex        =   6
      Top             =   3480
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
   Begin MSMask.MaskEdBox ValorReal 
      Height          =   225
      Left            =   5535
      TabIndex        =   7
      Top             =   3495
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
   Begin MSFlexGridLib.MSFlexGrid GridFCaixa 
      Height          =   2805
      Left            =   120
      TabIndex        =   8
      Top             =   1005
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
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
      Left            =   2415
      TabIndex        =   14
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label TotalSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3225
      TabIndex        =   15
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Label TotalReal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5475
      TabIndex        =   16
      Top             =   3825
      Width           =   1065
   End
   Begin VB.Label TotalAjustado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4350
      TabIndex        =   17
      Top             =   3810
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contas Correntes"
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
      Left            =   7440
      TabIndex        =   18
      Top             =   810
      Width           =   1470
   End
End
Attribute VB_Name = "FluxoSaldoInicialOcx"
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
Const GRID_CODIGO_COL = 2
Const GRID_NOMEREDUZIDO_COL = 3
Const GRID_VALORSISTEMA_COL = 4
Const GRID_VALORAJUSTADO_COL = 5
Const GRID_VALORREAL_COL = 6

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long
Dim colFluxoSldIni As New Collection

On Error GoTo Erro_ExibeFluxo_Click

    'le os saldos das contas
    lErro = CF("FluxoSaldosIniciais_Le",colFluxoSldIni, objFluxo1.lFluxoId)
    If lErro <> SUCESSO And lErro <> 21141 Then Error 21120

    If lErro = 21141 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOSALDOSINICIAIS_ULTRAPASSOU_LIMITE", objFluxo1.lFluxoId, MAX_FLUXO)

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 55926

    'preenche o grid com os saldos iniciais lidos
    lErro = Preenche_GridFCaixa(colFluxoSldIni)
    If lErro <> SUCESSO Then Error 21121

    GridFCaixa.Enabled = True

    iAlterado = 0

    Exit Sub

Erro_ExibeFluxo_Click:

    Select Case Err

        Case 21120, 21121, 55926

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160449)

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
    If lErro <> SUCESSO Then Error 21122

    Call Limpa_Tela_FluxoSldIni
    
    GridFCaixa.Enabled = False
                 
    iAlterado = 0
    iVAAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21122

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160450)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'grava os registros da tela

Dim lErro As Long
Dim colFluxoSldIni As New Collection
Dim objFluxoSldIni As ClassFluxoSldIni
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoSldIni = New ClassFluxoSldIni

        'se o codigo da conta não estiver preenchido
        If Len(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL)) = 0 Then Error 55913

        objFluxoSldIni.iCodConta = CInt(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL))
        objFluxoSldIni.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_NOMEREDUZIDO_COL)
        objFluxoSldIni.dSaldoAjustado = CDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        objFluxoSldIni.iUsuario = CInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoSldIni.lFluxoId = objFluxo1.lFluxoId
        
        colFluxoSldIni.Add objFluxoSldIni

    Next

    lErro = CF("FluxoSaldosIniciais_Grava1",colFluxoSldIni, objFluxo1.lFluxoId)
    If lErro <> SUCESSO Then Error 21123

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21123

        Case 55913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_CONTA_NAO_PREENCHIDA", Err, iLinha)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160451)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21124

    Call Limpa_Tela_FluxoSldIni
    
    GridFCaixa.Enabled = False

    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21124

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160452)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_FluxoSldIni()

    Call Limpa_GridFCaixa
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

Private Sub CCICodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'Carrega a Lista de CCI
    lErro = Carrega_ContasCorrente()
    If lErro <> SUCESSO Then Error 21125

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21126

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

        Case 21125, 21126

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160453)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_ContasCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_ContasCorrente

    'Leitura dos códigos e descrições das Contas existentes no BD
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed",colCodigoDescricao)
    If lErro <> SUCESSO Then Error 23450

    For Each objCodigoNome In colCodigoDescricao

        'Insere na combo de contas correntes
        ListaCCI.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        ListaCCI.ItemData(ListaCCI.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_ContasCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContasCorrente:

    Carrega_ContasCorrente = Err

    Select Case Err

        Case 23450

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160454)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then Error 21127

    Set objFluxo1 = objFluxo

    Call Botao_ExibeFluxo_Click

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Saldos Iniciais"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 21127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160455)

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
    objGrid1.colColuna.Add ("Conta")
    objGrid1.colColuna.Add ("Descrição")
    objGrid1.colColuna.Add ("Valor Sistema")
    objGrid1.colColuna.Add ("Valor Ajustado")
    objGrid1.colColuna.Add ("Valor Real")

    'campos de edição do grid
    objGrid1.colCampo.Add (Usuario.Name)
    objGrid1.colCampo.Add (CCICodigo.Name)
    objGrid1.colCampo.Add (CCINomeReduzido.Name)
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

Function Preenche_GridFCaixa(colFluxoSldIni As Collection) As Long
'preenche o grid com os saldos contidos na coleção colFluxoSldIni

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoSldIni As ClassFluxoSldIni
Dim dColunaSomaSistema As Double
Dim dColunaSomaAjustado As Double
Dim dColunaSomaReal As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoSldIni.Count < objGrid1.iLinhasVisiveis Then
        GridFCaixa.Rows = 100
    Else
        GridFCaixa.Rows = colFluxoSldIni.Count + 100
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoSldIni.Count

    dColunaSomaSistema = 0
    dColunaSomaAjustado = 0
    dColunaSomaReal = 0

    'preenche o grid com os dados retornados na coleção colFluxoSldIni
    For iIndice = 1 To colFluxoSldIni.Count

        Set objFluxoSldIni = colFluxoSldIni.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = CStr(objFluxoSldIni.iUsuario)
        GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_COL) = CStr(objFluxoSldIni.iCodConta)
        GridFCaixa.TextMatrix(iIndice, GRID_NOMEREDUZIDO_COL) = objFluxoSldIni.sNomeReduzido
        GridFCaixa.TextMatrix(iIndice, GRID_VALORSISTEMA_COL) = Format(objFluxoSldIni.dSaldoSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORAJUSTADO_COL) = Format(objFluxoSldIni.dSaldoAjustado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORREAL_COL) = Format(objFluxoSldIni.dSaldoReal, "Standard")
        dColunaSomaSistema = dColunaSomaSistema + objFluxoSldIni.dSaldoSistema
        dColunaSomaAjustado = dColunaSomaAjustado + objFluxoSldIni.dSaldoAjustado
        dColunaSomaReal = dColunaSomaReal + objFluxoSldIni.dSaldoReal

    Next

    For iIndice = colFluxoSldIni.Count + 1 To GridFCaixa.Rows - 1
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
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160456)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub CCICodigo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub CCICodigo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub CCICodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = CCICodigo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub ValorAjustado_Change()

    iAlterado = REGISTRO_ALTERADO

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

Private Sub ListaCCI_DblClick()

Dim iLinha As Integer

On Error GoTo Erro_ListaCCI_DblClick

    If GridFCaixa.Col = GRID_CODIGO_COL And GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then

        'verifica se há alguma linha que já exibe esta conta
        For iLinha = 1 To objGrid1.iLinhasExistentes
        
            If GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL) = CStr(ListaCCI.ItemData(ListaCCI.ListIndex)) Then Error 55914

        Next
        
        GridFCaixa.TextMatrix(GridFCaixa.Row, GridFCaixa.Col) = CStr(ListaCCI.ItemData(ListaCCI.ListIndex))
        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_NOMEREDUZIDO_COL) = Mid(ListaCCI.Text, InStr(ListaCCI.Text, SEPARADOR) + 1)
        
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL) = Format(0, "Standard")
        
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
        
    End If

    Exit Sub
    
Erro_ListaCCI_DblClick:

    Select Case Err
    
        Case 55914
             Call Rotina_Erro(vbOKOnly, "ERRO_CCI_EXISTENTE_GRID", Err, GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL), iLinha)
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160457)
    
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

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridFCaixa.Col

            Case GRID_CODIGO_COL

                lErro = Saida_Celula_CCICodigo(objGridInt)
                If lErro <> SUCESSO Then Error 21129

            Case GRID_VALORAJUSTADO_COL

                lErro = Saida_Celula_ValorAjustado(objGridInt)
                If lErro <> SUCESSO Then Error 21130

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 21131

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21129, 21130

        Case 21131
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160458)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CCICodigo(objGridInt As AdmGrid) As Long
'faz a critica da célula codigo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sListBoxItem As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCCI As New ClassContasCorrentesInternas
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_CCICodigo

    Set objGridInt.objControle = CCICodigo

    'se o codigo foi preenchido
    If Len(CCICodigo.Text) > 0 Then

        lErro = Valor_Critica(CCICodigo.Text)
        If lErro <> SUCESSO Then Error 21132
        
        'verifica se a conta está cadastrada
        lErro = CF("ContaCorrenteInt_Le",CInt(CCICodigo.Text), objCCI)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 21133

        'conta não cadastrada
        If lErro = 11807 Then Error 21134
        
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridFCaixa.Row And GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_COL) = CCICodigo.Text Then Error 21160
        Next

        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_NOMEREDUZIDO_COL) = objCCI.sNomeReduzido
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL) = Format(0, "Standard")

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21135

    Saida_Celula_CCICodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_CCICodigo:

    Saida_Celula_CCICodigo = Err

    Select Case Err

        Case 21132, 21133, 21135
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21134
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", CInt(CCICodigo.Text))

            If vbMsgRes = vbYes Then

                objCCI.iCodigo = CInt(CCICodigo.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("CtaCorrenteInt", objCCI)
                
            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
            
        Case 21160
            Call Rotina_Erro(vbOKOnly, "ERRO_CCI_EXISTENTE_GRID", GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_COL), iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160459)

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

        lErro = Valor_Critica(ValorAjustado.Text)
        If lErro <> SUCESSO Then Error 21136

        If GridFCaixa.Row - GridFCaixa.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_CODIGO_COL)) > 0 Then ValorAjustado.Text = "0"
    End If

    If Format(ValorAjustado.Text, "Standard") <> GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) Then
        iVAAlterado = REGISTRO_ALTERADO
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21137
    
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

        Case 21136, 21137
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160460)

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

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

    'se estiver trabalhando na campo de fornecedor
    If objControl.Name = "CCICodigo" Then

        'se for uma linha criada pelo usuario ==> habilita o campo
        If GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = "ValorAjustado" Then
    
        If Len(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL)) > 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    End If

End Sub

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoSldIni As New ClassFluxoSldIni
Dim colFluxoSldIni As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOSLDINI_CPR")
    If lErro <> SUCESSO Then Error 47918
    
    lErro = Grid_FCaixa_Obter(colFluxoSldIni)
    If lErro <> SUCESSO Then Error 47919
    
    For iIndice1 = 1 To colFluxoSldIni.Count
    
        Set objFluxoSldIni = colFluxoSldIni.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(objFluxoSldIni.iUsuario)
        Call colTemp.Add(objFluxoSldIni.iCodConta)
        Call colTemp.Add(objFluxoSldIni.sNomeReduzido)
        Call colTemp.Add(objFluxoSldIni.dSaldoSistema)
        Call colTemp.Add(objFluxoSldIni.dSaldoAjustado)
        Call colTemp.Add(objFluxoSldIni.dSaldoReal)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47920
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47921
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47918, 47919, 47920, 47921
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160461)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoSldIni As Collection) As Long

Dim objFluxoSldIni As ClassFluxoSldIni
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoSldIni = New ClassFluxoSldIni
        
        objFluxoSldIni.iUsuario = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoSldIni.iCodConta = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL))
        objFluxoSldIni.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_NOMEREDUZIDO_COL)
        objFluxoSldIni.dSaldoSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORSISTEMA_COL))
        objFluxoSldIni.dSaldoAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        objFluxoSldIni.dSaldoReal = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORREAL_COL))
        
        colFluxoSldIni.Add objFluxoSldIni
        
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160462)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_SALDOS_INICIAIS
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Saldos Iniciais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoSaldoInicial"
    
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

