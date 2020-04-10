VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FluxoTipoAplicOcx 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   4560
   ScaleWidth      =   9480
   Begin VB.CommandButton BotaoDataDown 
      Height          =   150
      Left            =   2160
      Picture         =   "FluxoTipoAplicOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   405
      Width           =   240
   End
   Begin VB.CommandButton BotaoDataUp 
      Height          =   150
      Left            =   2160
      Picture         =   "FluxoTipoAplicOcx.ctx":005A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   255
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7650
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "FluxoTipoAplicOcx.ctx":00B4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton botaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoTipoAplicOcx.ctx":0232
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoTipoAplicOcx.ctx":0764
         Style           =   1  'Graphical
         TabIndex        =   12
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
      Height          =   585
      Left            =   4470
      Picture         =   "FluxoTipoAplicOcx.ctx":08BE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
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
      Height          =   585
      Left            =   2820
      Picture         =   "FluxoTipoAplicOcx.ctx":09C0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1290
   End
   Begin VB.ListBox ListaTiposAplicacao 
      Height          =   3180
      Left            =   7125
      TabIndex        =   10
      Top             =   1095
      Width           =   2220
   End
   Begin VB.CheckBox Usuario 
      Enabled         =   0   'False
      Height          =   210
      Left            =   315
      TabIndex        =   3
      Top             =   3915
      Width           =   615
   End
   Begin VB.TextBox DescTipoAplic 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1830
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3855
      Width           =   1380
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   255
      Left            =   915
      TabIndex        =   4
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorAjustado 
      Height          =   225
      Left            =   4530
      TabIndex        =   7
      Top             =   3840
      Width           =   1110
      _ExtentX        =   1958
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
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   990
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
      Left            =   3300
      TabIndex        =   6
      Top             =   3885
      Width           =   1110
      _ExtentX        =   1958
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
      Left            =   5685
      TabIndex        =   8
      Top             =   3885
      Width           =   1110
      _ExtentX        =   1958
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
      TabIndex        =   9
      Top             =   885
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      HighLight       =   0
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
      Left            =   420
      TabIndex        =   17
      Top             =   300
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
      Left            =   2535
      TabIndex        =   18
      Top             =   4155
      Width           =   600
   End
   Begin VB.Label TotalSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   4140
      Width           =   1050
   End
   Begin VB.Label TotalReal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5595
      TabIndex        =   20
      Top             =   4155
      Width           =   1050
   End
   Begin VB.Label TotalAjustado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4485
      TabIndex        =   21
      Top             =   4140
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipos de Aplicação"
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
      Left            =   7140
      TabIndex        =   22
      Top             =   810
      Width           =   1650
   End
End
Attribute VB_Name = "FluxoTipoAplicOcx"
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
Const GRID_DESCTIPOAPLIC_COL = 3
Const GRID_VALORSISTEMA_COL = 4
Const GRID_VALORAJUSTADO_COL = 5
Const GRID_VALORREAL_COL = 6

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long
Dim colFluxoTipoAplic As New Collection

On Error GoTo Erro_ExibeFluxo_Click

    'se a data da tela não estiver preenchida ==> não exibe os dados no grid
    If Len(Data.ClipText) = 0 Then Error 21177

    'le os FluxoTipoAplic por data e por fluxo
    lErro = CF("FluxoTipoAplic_Le",colFluxoTipoAplic, objFluxo1.lFluxoId, CDate(Data.Text))
    If lErro <> SUCESSO And lErro <> 21204 Then Error 21178

    If lErro = 21204 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOTIPOAPLIC_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 55929

    'preenche o grid com as aplicações lidas
    lErro = Preenche_GridFCaixa(colFluxoTipoAplic)
    If lErro <> SUCESSO Then Error 21179

    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub

Erro_ExibeFluxo_Click:

    Select Case Err

        Case 21177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 21178, 21179, 55929

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160491)

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
    If lErro <> SUCESSO Then Error 21180

    Call Limpa_Tela_FluxoTipoAplic

    iAlterado = 0
    
    iVAAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160492)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'grava os registros da tela

Dim lErro As Long
Dim colFluxoTipoAplic As New Collection
Dim objFluxoTipoAplic As ClassFluxoTipoAplic
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se a data da tela não estiver preenchida ==> erro
    If Len(Data.ClipText) = 0 Then Error 21181

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoTipoAplic = New ClassFluxoTipoAplic

        objFluxoTipoAplic.lFluxoId = objFluxo1.lFluxoId
        objFluxoTipoAplic.dtData = CDate(Data.Text)
        objFluxoTipoAplic.iUsuario = CInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoTipoAplic.iTipoAplicacao = CInt(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL))
        objFluxoTipoAplic.dTotalAjustado = CDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))

        colFluxoTipoAplic.Add objFluxoTipoAplic

    Next

    lErro = CF("FluxoTipoAplic_Grava",colFluxoTipoAplic, objFluxo1.lFluxoId, CDate(Data.Text))
    If lErro <> SUCESSO Then Error 21182

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 21182

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160493)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21183

    Call Limpa_Tela_FluxoTipoAplic

    iAlterado = 0
    iVAAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21183

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160494)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_FluxoTipoAplic()

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
Dim colTipoAplic As New AdmColCodigoNome
Dim objTipoAplic As AdmCodigoNome
Dim sListBoxItem As String

On Error GoTo Erro_Form_Load

    'Le cada Codigo e Descrição da tabela de Tipos de Aplicação e poe na colecao
    lErro = CF("Cod_Nomes_Le","TiposDeAplicacao", "Codigo", "Descricao", STRING_TIPO_APLICACAO_DESCRICAO, colTipoAplic)
    If lErro <> SUCESSO Then Error 21191

    'preenche a ListBox ListaTiposAplicacao com os objetos da colecao
    For Each objTipoAplic In colTipoAplic

        'Espaços que faltam para completar tamanho STRING_CODIGO_APLICACAO
        sListBoxItem = Space(STRING_CODIGO_APLICACAO - Len(CStr(objTipoAplic.iCodigo)))

        'Concatena Código e a Descrição da Aplicação
        sListBoxItem = sListBoxItem & CStr(objTipoAplic.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & Trim(objTipoAplic.sNome)

        ListaTiposAplicacao.AddItem sListBoxItem
        ListaTiposAplicacao.ItemData(ListaTiposAplicacao.NewIndex) = objTipoAplic.iCodigo
        
    Next
    
    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21192

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

        Case 21191, 21192

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160495)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then Error 21184

    Set objFluxo1 = objFluxo

    Data.Text = Format(objFluxo.dtDataBase, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Resgates por Tipo de Aplicação"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 21184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160496)

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
    objGrid1.colColuna.Add ("Código")
    objGrid1.colColuna.Add ("Tipo de Aplicação")
    objGrid1.colColuna.Add ("Valor Sistema")
    objGrid1.colColuna.Add ("Valor Ajustado")
    objGrid1.colColuna.Add ("Valor Real")

   'campos de edição do grid
    objGrid1.colCampo.Add (Usuario.Name)
    objGrid1.colCampo.Add (Codigo.Name)
    objGrid1.colCampo.Add (DescTipoAplic.Name)
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

Function Preenche_GridFCaixa(colFluxoTipoAplic As Collection) As Long
'Preenche o grid com as aplicações contidas na coleção colFluxoTipoAplic

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoTipoAplic As ClassFluxoTipoAplic
Dim dColunaSomaSistema As Double
Dim dColunaSomaAjustado As Double
Dim dColunaSomaReal As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoTipoAplic.Count < objGrid1.iLinhasVisiveis Then
        GridFCaixa.Rows = 100
    Else
        GridFCaixa.Rows = colFluxoTipoAplic.Count + 100
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoTipoAplic.Count

    dColunaSomaSistema = 0
    dColunaSomaAjustado = 0
    dColunaSomaReal = 0

    'preenche o grid com os dados retornados na coleção colFluxoTipoAplic
    For iIndice = 1 To colFluxoTipoAplic.Count

        Set objFluxoTipoAplic = colFluxoTipoAplic.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_USUARIO_COL) = CStr(objFluxoTipoAplic.iUsuario)
        GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_COL) = CStr(objFluxoTipoAplic.iTipoAplicacao)
        GridFCaixa.TextMatrix(iIndice, GRID_DESCTIPOAPLIC_COL) = objFluxoTipoAplic.sDescricao
        GridFCaixa.TextMatrix(iIndice, GRID_VALORSISTEMA_COL) = Format(objFluxoTipoAplic.dTotalSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORAJUSTADO_COL) = Format(objFluxoTipoAplic.dTotalAjustado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALORREAL_COL) = Format(objFluxoTipoAplic.dTotalReal, "Standard")
        dColunaSomaSistema = dColunaSomaSistema + objFluxoTipoAplic.dTotalSistema
        dColunaSomaAjustado = dColunaSomaAjustado + objFluxoTipoAplic.dTotalAjustado
        dColunaSomaReal = dColunaSomaReal + objFluxoTipoAplic.dTotalReal


    Next

    For iIndice = colFluxoTipoAplic.Count + 1 To GridFCaixa.Rows - 1
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
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160497)

    End Select

    Exit Function

End Function

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Data_Validate

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 21185

        dtData = CDate(Data.Text)

        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 21186

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 21185

        Case 21186
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160498)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Codigo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing

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

Private Sub ListaTiposAplicacao_DblClick()

Dim iLinha As Integer

On Error GoTo Erro_ListaTiposAplicacao_DblClick

    If GridFCaixa.Col = GRID_CODIGO_COL And GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then

        'verifica se há alguma linha que já exibe este tipo de aplicacao
        For iLinha = 1 To objGrid1.iLinhasExistentes
        
            If GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL) = CStr(ListaTiposAplicacao.ItemData(ListaTiposAplicacao.ListIndex)) Then Error 55918

        Next
    
        GridFCaixa.TextMatrix(GridFCaixa.Row, GridFCaixa.Col) = CStr(ListaTiposAplicacao.ItemData(ListaTiposAplicacao.ListIndex))
        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_DESCTIPOAPLIC_COL) = Mid(ListaTiposAplicacao.Text, InStr(ListaTiposAplicacao.Text, SEPARADOR) + 1)
        
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORSISTEMA_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORREAL_COL) = Format(0, "Standard")
    
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

    End If

    Exit Sub
    
Erro_ListaTiposAplicacao_DblClick:

    Select Case Err
    
        Case 55918
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEAPLICACAO_CODIGO_REPETIDO", Err, GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL), iLinha)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160499)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDataDown_Click()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_BotaoDataDown_Click

    If Len(Data.ClipText) > 0 Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 55939

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 21187

        dtData = CDate(sData)

        If dtData < objFluxo1.dtDataBase Then
            Data.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
        ElseIf dtData > objFluxo1.dtDataFinal Then
            Data.Text = Format(objFluxo1.dtDataFinal, "dd/mm/yy")
        Else
            Data.Text = sData
        End If

    Else
        Data.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")

    End If

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataDown_Click:

    Select Case Err

        Case 21187, 55939

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160500)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataUp_Click()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_BotaoDataUp_Click

    If Len(Data.ClipText) > 0 Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 55940

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 21188

        dtData = CDate(sData)

        If dtData < objFluxo1.dtDataBase Then
            Data.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")
        ElseIf dtData > objFluxo1.dtDataFinal Then
            Data.Text = Format(objFluxo1.dtDataFinal, "dd/mm/yy")
        Else
            Data.Text = sData
        End If

    Else
        Data.Text = Format(objFluxo1.dtDataBase, "dd/mm/yy")

    End If

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataUp_Click:

    Select Case Err

        Case 21188, 55940

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160501)

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

                lErro = Saida_Celula_Codigo(objGridInt)
                If lErro <> SUCESSO Then Error 21194

            Case GRID_VALORAJUSTADO_COL

                lErro = Saida_Celula_ValorAjustado(objGridInt)
                If lErro <> SUCESSO Then Error 21195

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 21196

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21194, 21195

        Case 21196
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160502)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Codigo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula código do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoAplic As New ClassTiposDeAplicacao

On Error GoTo Erro_Saida_Celula_Codigo

    Set objGridInt.objControle = Codigo

    'Se o código foi preenchido
    If Len(Codigo.Text) > 0 Then

        objTipoAplic.iCodigo = CInt(Codigo.Text)

        'Verifica se o código está cadastrado
        lErro = CF("TiposDeAplicacao_Le",objTipoAplic)
        If lErro <> SUCESSO And lErro <> 15068 Then Error 21197

        'Código não cadastrado
        If lErro = 15068 Then Error 21198
        
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If Codigo.Text = GridFCaixa.TextMatrix(iIndice, GRID_CODIGO_COL) And iIndice <> GridFCaixa.Row Then Error 21220
        Next

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_DESCTIPOAPLIC_COL) = objTipoAplic.sDescricao
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL)) = 0 Then GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALORAJUSTADO_COL) = Format(0, "Standard")
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21199

    Saida_Celula_Codigo = SUCESSO

    Exit Function

Erro_Saida_Celula_Codigo:

    Saida_Celula_Codigo = Err

    Select Case Err

        Case 21197, 21199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21198
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOAPLICACAO", objTipoAplic.iCodigo)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("TipoAplicacao", objTipoAplic)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
            
        Case 21220
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEAPLICACAO_CODIGO_REPETIDO", objTipoAplic.iCodigo, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160503)

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
        If lErro <> SUCESSO Then Error 21189

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
    If lErro <> SUCESSO Then Error 21190
    
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

        Case 21189, 21190
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160504)

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

    'se estiver trabalhando na campo de Código
    If objControl.Name = "Codigo" Then

        'se for uma linha criada pelo usuario ==> habilita o campo
        If GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL) = GRID_CHECKBOX_ATIVO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    ElseIf objControl.Name = "ValorAjustado" Then
    
        If Len(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_CODIGO_COL)) > 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    End If

End Sub

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoTipoAplic As New ClassFluxoTipoAplic
Dim colFluxoTipoAplic As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOTIPOAPLIC_CPR")
    If lErro <> SUCESSO Then Error 47930
    
    'obter dados comuns a todas as linhas do grid
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoTipoAplic)
    If lErro <> SUCESSO Then Error 47931
    
    For iIndice1 = 1 To colFluxoTipoAplic.Count
    
        Set objFluxoTipoAplic = colFluxoTipoAplic.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoTipoAplic.iUsuario)
        Call colTemp.Add(objFluxoTipoAplic.iTipoAplicacao)
        Call colTemp.Add(objFluxoTipoAplic.sDescricao)
        Call colTemp.Add(objFluxoTipoAplic.dTotalSistema)
        Call colTemp.Add(objFluxoTipoAplic.dTotalAjustado)
        Call colTemp.Add(objFluxoTipoAplic.dTotalReal)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47932
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47933
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47930, 47931, 47932, 47933
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160505)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoTipoAplic As Collection) As Long

Dim objFluxoTipoAplic As ClassFluxoTipoAplic
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoTipoAplic = New ClassFluxoTipoAplic
        
        objFluxoTipoAplic.iUsuario = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_USUARIO_COL))
        objFluxoTipoAplic.iTipoAplicacao = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_CODIGO_COL))
        objFluxoTipoAplic.sDescricao = GridFCaixa.TextMatrix(iLinha, GRID_DESCTIPOAPLIC_COL)
        objFluxoTipoAplic.dTotalSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORSISTEMA_COL))
        objFluxoTipoAplic.dTotalAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORAJUSTADO_COL))
        objFluxoTipoAplic.dTotalReal = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALORREAL_COL))
        
        colFluxoTipoAplic.Add objFluxoTipoAplic
 
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160506)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_RESGATE_TIPO_APLICACAO
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Resgates por Tipo de Aplicação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoTipoAplic"
    
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

