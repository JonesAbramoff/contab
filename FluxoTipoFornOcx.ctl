VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FluxoTipoFornOcx 
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ScaleHeight     =   4590
   ScaleWidth      =   6720
   Begin VB.CommandButton BotaoDataUp 
      Height          =   150
      Left            =   1950
      Picture         =   "FluxoTipoFornOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   315
      Width           =   240
   End
   Begin VB.CommandButton BotaoDataDown 
      Height          =   150
      Left            =   1950
      Picture         =   "FluxoTipoFornOcx.ctx":005A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   465
      Width           =   240
   End
   Begin VB.PictureBox Picture6 
      Height          =   555
      Left            =   5340
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1170
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoTipoFornOcx.ctx":00B4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoTipoFornOcx.ctx":0232
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox DescTipoForn 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   780
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3900
      Width           =   2175
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
      Height          =   555
      Left            =   2400
      Picture         =   "FluxoTipoFornOcx.ctx":0764
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1290
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
      Height          =   555
      Left            =   3870
      Picture         =   "FluxoTipoFornOcx.ctx":0AB2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1290
   End
   Begin MSMask.MaskEdBox ValorAjustado 
      Height          =   225
      Left            =   4185
      TabIndex        =   5
      Top             =   3870
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
      TabIndex        =   10
      Top             =   1020
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
      Left            =   795
      TabIndex        =   0
      Top             =   315
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
      Left            =   3000
      TabIndex        =   4
      Top             =   3915
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
      Left            =   5400
      TabIndex        =   6
      Top             =   3870
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
   Begin VB.Label TotalAjustado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4155
      TabIndex        =   11
      Top             =   4155
      Width           =   1065
   End
   Begin VB.Label TotalReal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   4170
      Width           =   1065
   End
   Begin VB.Label TotalSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3015
      TabIndex        =   13
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
      Left            =   2220
      TabIndex        =   14
      Top             =   4185
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
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FluxoTipoFornOcx"
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
Const GRID_DESCRICAO_TIPO_FORNECEDOR_COL = 1
Const GRID_VALOR_SISTEMA_COL = 2
Const GRID_VALOR_AJUSTADO_COL = 3
Const GRID_VALOR_REAL_COL = 4

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long

On Error GoTo Erro_Botao_ExibeFluxo_Click

    'se a data da tela não estiver preenchido ==> erro
    If Len(Data.ClipText) = 0 Then Error 21044

    lErro = Ordenados()
    If lErro <> SUCESSO Then Error 55933

    Exit Sub

Erro_Botao_ExibeFluxo_Click:

    Select Case Err

        Case 21044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 55933

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160518)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataDown_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataDown_Click

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoTipoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 0)
    If lErro <> SUCESSO And lErro <> 133491 Then gError 133500

    If lErro = 133491 Then gError 133501

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataDown_Click:

    Select Case gErr

        Case 133500

        Case 133501
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_AQUEM_DESTA_DATA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160519)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataUp_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataUp_Click

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoTipoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 1)
    If lErro <> SUCESSO And lErro <> 133491 Then gError 133502

    If lErro = 133491 Then gError 133503

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataUp_Click:

    Select Case gErr

        Case 133502
        
        Case 133503
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_ALEM_DESTA_DATA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160520)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_FluxoTipoForn

End Sub

Sub Limpa_Tela_FluxoTipoForn()

    Call Grid_Limpa(objGrid1)
    Data.PromptInclude = False
    Data.Text = ""
    Data.PromptInclude = True
    TotalSistema.Caption = ""
    TotalReal.Caption = ""
    TotalAjustado.Caption = ""

End Sub

Private Sub Data_Change()

    If objGrid1.iLinhasExistentes > 0 Then
        Call Grid_Limpa(objGrid1)
    End If

    TotalSistema.Caption = ""
    TotalReal.Caption = ""
    TotalAjustado.Caption = ""

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21045
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21045

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160521)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colFluxoAnalitico As New Collection
Dim dtData As Date
Dim sOrdenacao As String

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then gError 21046

    Set objFluxo1 = objFluxo

    'le os pagamentos selecionados
    lErro = CF("FluxoTipoForn_Le", colFluxoAnalitico, sOrdenacao, objFluxo1.lFluxoId, objFluxo1.dtData, FLUXOANALITICO_TIPOREG_PAGTO)
    If lErro <> SUCESSO And lErro <> 21059 Then gError 133492
    
    If colFluxoAnalitico.Count = 0 Then
    
        dtData = objFluxo1.dtData

        'le os recebimentos selecionados
        lErro = CF("FluxoTipoForn_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_PAGTO, 1)
        If lErro <> SUCESSO And lErro <> 133491 Then gError 133493

        If lErro = SUCESSO Then objFluxo1.dtData = dtData
    
    End If

    Data.Text = Format(objFluxo1.dtData, "dd/mm/yy")

    lErro = Ordenados()
    If lErro <> SUCESSO Then gError 55934

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Pagamentos por Tipo de Fornecedor"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 21046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", gErr)

        Case 55934, 133492, 133493

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160522)

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
    objGrid1.colColuna.Add ("Descr. Tipo Fornecedor")
    objGrid1.colColuna.Add ("Valor Sistema")
    objGrid1.colColuna.Add ("Valor Ajustado")
    objGrid1.colColuna.Add ("Valor Real")

   'campos de edição do grid
    objGrid1.colCampo.Add (DescTipoForn.Name)
    objGrid1.colCampo.Add (ValorSistema.Name)
    objGrid1.colCampo.Add (ValorAjustado.Name)
    objGrid1.colCampo.Add (ValorReal.Name)

    objGrid1.objGrid = GridFCaixa

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 10

    objGrid1.objGrid.ColWidth(0) = 300

    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid1)

     'Posiciona o totalizador
    TotalSistema.Top = GridFCaixa.Top + GridFCaixa.Height
    TotalSistema.Left = GridFCaixa.Left
    
    For iIndice = 0 To GRID_VALOR_SISTEMA_COL - 1
        TotalSistema.Left = TotalSistema.Left + GridFCaixa.ColWidth(iIndice) + GridFCaixa.GridLineWidth + 10
    Next
    
    TotalSistema.Width = GridFCaixa.ColWidth(GRID_VALOR_SISTEMA_COL)
    
    TotalAjustado.Top = TotalSistema.Top
    TotalAjustado.Left = TotalSistema.Left + TotalSistema.Width + GridFCaixa.GridLineWidth
    TotalAjustado.Width = GridFCaixa.ColWidth(GRID_VALOR_AJUSTADO_COL)
    
    TotalReal.Top = TotalAjustado.Top
    TotalReal.Left = TotalAjustado.Left + TotalAjustado.Width + GridFCaixa.GridLineWidth
    TotalReal.Width = GridFCaixa.ColWidth(GRID_VALOR_REAL_COL)
    
    LabelTotais.Top = TotalSistema.Top + (TotalSistema.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalSistema.Left - LabelTotais.Width - 50

    Inicializa_GridFCaixa = SUCESSO

End Function

Function Preenche_GridFCaixa(colFluxoTipoForn As Collection) As Long
'preenche o grid com os pagamentos contidos na coleção colFluxoTipoForn

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoTipoForn As ClassFluxoTipoForn
Dim dColunaSomaAjustado As Double
Dim dColunaSomaReal As Double
Dim dColunaSomaSistema As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoTipoForn.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colFluxoTipoForn.Count + 1
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoTipoForn.Count

    dColunaSomaReal = 0
    dColunaSomaAjustado = 0
    dColunaSomaSistema = 0

    'preenche o grid com os dados retornados na coleção colFluxoTipoForn
    For iIndice = 1 To colFluxoTipoForn.Count

        Set objFluxoTipoForn = colFluxoTipoForn.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_DESCRICAO_TIPO_FORNECEDOR_COL) = objFluxoTipoForn.sDescricao
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_COL) = Format(-objFluxoTipoForn.dTotalSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_COL) = Format(-objFluxoTipoForn.dTotalAjustado, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_REAL_COL) = Format(-objFluxoTipoForn.dTotalReal, "Standard")
        dColunaSomaReal = dColunaSomaReal + objFluxoTipoForn.dTotalReal
        dColunaSomaAjustado = dColunaSomaAjustado + objFluxoTipoForn.dTotalAjustado
        dColunaSomaSistema = dColunaSomaSistema + objFluxoTipoForn.dTotalSistema
        
    Next

    TotalReal.Caption = Format(-dColunaSomaReal, "Standard")
    TotalAjustado.Caption = Format(-dColunaSomaAjustado, "Standard")
    TotalSistema.Caption = Format(-dColunaSomaSistema, "Standard")
    
    Preenche_GridFCaixa = SUCESSO

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160523)

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
        If lErro <> SUCESSO Then Error 21048

        dtData = CDate(Data.Text)

        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 21049

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 21048

        Case 21049
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160524)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing

End Sub

Private Sub GridFCaixa_DblClick()

    If GridFCaixa.Row > 0 And Len(Data.ClipText) > 0 Then

        objFluxo1.dtData = StrParaDate(Data.Text)
    
        Call Chama_Tela("FluxoPag", objFluxo1)
    
    End If

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
'faz a critica da celula do grid que está deixando de ser a corrente /m

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 21052

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21052
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160525)

    End Select

    Exit Function

End Function

Private Function Ordenados() As Long

Dim sOrdenacao As String
Dim lErro As Long
Dim colFluxoTipoForn As New Collection

On Error GoTo Erro_Ordenados

    'se a data da tela não estiver preenchido ==> não exibe os dados no grid
    If Len(Data.ClipText) = 0 Then Exit Function

    'monta a expressão SQL de Ordenação
    sOrdenacao = " ORDER BY TipoFornecedor"

    'le os pagamentos selecionados
    lErro = CF("FluxoTipoForn_Le", colFluxoTipoForn, sOrdenacao, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOTIPOFORN_TIPOREG_PAGTO)
    If lErro <> SUCESSO And lErro <> 21059 Then Error 21061

    If lErro = 21059 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOTIPOFORN_PAGTO_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)

    'preenche o grid com os pagamentos lidos
    lErro = Preenche_GridFCaixa(colFluxoTipoForn)
    If lErro <> SUCESSO Then Error 21062

    Ordenados = SUCESSO

    Exit Function

Erro_Ordenados:

    Ordenados = Err

    Select Case Err

        Case 21061, 21062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160526)

    End Select

    Exit Function

End Function

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoTipoForn As New ClassFluxoTipoForn
Dim colFluxoTipoForn As New Collection
Dim dtData As Date

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOTIPOFORN_CPR")
    If lErro <> SUCESSO Then Error 47934
    
    'obter dados comuns a todas as linhas do grid
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoTipoForn)
    If lErro <> SUCESSO Then Error 47935
    
    For iIndice1 = 1 To colFluxoTipoForn.Count
    
        Set objFluxoTipoForn = colFluxoTipoForn.Item(iIndice1)
        
        Set colTemp = New Collection
                
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoTipoForn.sDescricao)
        Call colTemp.Add(objFluxoTipoForn.dTotalSistema)
        Call colTemp.Add(objFluxoTipoForn.dTotalAjustado)
        Call colTemp.Add(objFluxoTipoForn.dTotalReal)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47936
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47937
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47934, 47935, 47936, 47937
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160527)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoTipoForn As Collection) As Long

Dim objFluxoTipoForn As ClassFluxoTipoForn
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoTipoForn = New ClassFluxoTipoForn
        
        objFluxoTipoForn.sDescricao = GridFCaixa.TextMatrix(iLinha, GRID_DESCRICAO_TIPO_FORNECEDOR_COL)
        objFluxoTipoForn.dTotalSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_SISTEMA_COL))
        objFluxoTipoForn.dTotalAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_AJUSTADO_COL))
        objFluxoTipoForn.dTotalReal = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_REAL_COL))
        
        colFluxoTipoForn.Add objFluxoTipoForn
 
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
        Select Case Err
            
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160528)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_PAGAMENTO_TIPO_FORNECEDOR
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Pagamentos por Tipo de Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoTipoForn"
    
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
Private Sub ValorAjustado_Change()

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

