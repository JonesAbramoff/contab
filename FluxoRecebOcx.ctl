VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl FluxoRecebOcx 
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   4785
   ScaleWidth      =   8655
   Begin VB.CommandButton BotaoDataUp 
      Height          =   150
      Left            =   2010
      Picture         =   "FluxoRecebOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Width           =   240
   End
   Begin VB.CommandButton BotaoDataDown 
      Height          =   150
      Left            =   2010
      Picture         =   "FluxoRecebOcx.ctx":005A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   240
   End
   Begin VB.ComboBox Ordenados 
      Height          =   315
      ItemData        =   "FluxoRecebOcx.ctx":00B4
      Left            =   1620
      List            =   "FluxoRecebOcx.ctx":00C1
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   3390
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
      Left            =   2670
      Picture         =   "FluxoRecebOcx.ctx":00F9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1425
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
      Left            =   4410
      Picture         =   "FluxoRecebOcx.ctx":0447
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1425
   End
   Begin VB.PictureBox Picture6 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoRecebOcx.ctx":0549
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoRecebOcx.ctx":0A7B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Height          =   555
      Left            =   165
      Picture         =   "FluxoRecebOcx.ctx":0BF9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4155
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid GridFCaixa 
      Height          =   2805
      Left            =   150
      TabIndex        =   5
      Top             =   1290
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   855
      TabIndex        =   0
      Top             =   810
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   5400
      TabIndex        =   11
      Top             =   4125
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   225
      Left            =   570
      TabIndex        =   12
      Top             =   4125
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumTitulo 
      Height          =   225
      Left            =   3690
      TabIndex        =   13
      Top             =   4125
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SiglaDocumento 
      Height          =   225
      Left            =   2790
      TabIndex        =   14
      Top             =   4125
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
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
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Filial 
      Height          =   225
      Left            =   2055
      TabIndex        =   15
      Top             =   4125
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   6180
      TabIndex        =   16
      Top             =   4125
      Width           =   990
      _ExtentX        =   1746
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
   Begin MSMask.MaskEdBox Item 
      Height          =   225
      Left            =   4890
      TabIndex        =   17
      Top             =   4125
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   7230
      TabIndex        =   22
      Top             =   4125
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Label TotalValor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3780
      TabIndex        =   21
      Top             =   4380
      Width           =   1155
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
      Left            =   3135
      TabIndex        =   20
      Top             =   4365
      Width           =   600
   End
   Begin VB.Label Label59 
      Caption         =   "Ordenadas por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   19
      Top             =   120
      Width           =   1380
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
      Left            =   270
      TabIndex        =   18
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "FluxoRecebOcx"
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
Const GRID_CLIENTE_COL = 1
Const GRID_FILIAL_COL = 2
Const GRID_SIGLA_DOCUMENTO_COL = 3
Const GRID_NUMTITULO_COL = 4
Const GRID_ITEM_COL = 5
Const GRID_PARCELA_COL = 6
Const GRID_VALOR_COL = 7
Const GRID_FILIALEMPRESA_COL = 8

'tipos de ordenação dos grids
Const ORDENACAO_CLIENTEFILIAL = 1
Const ORDENACAO_SIGLADOCUMENTO = 2
Const ORDENACAO_TITULOPARCELA = 3

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long

On Error GoTo Erro_Botao_ExibeFluxo_Click

    'se a data da tela não estiver preenchido ==> erro
    If Len(Data.ClipText) = 0 Then Error 20201

    Call Ordenados_Click

    Exit Sub

Erro_Botao_ExibeFluxo_Click:

    Select Case Err

        Case 20201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160436)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataDown_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataDown_Click

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoAnalitico_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 0)
    If lErro <> SUCESSO And lErro <> 133191 Then gError 133455

    If lErro = 133191 Then gError 133456

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataDown_Click:

    Select Case gErr

        Case 133455

        Case 133456
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_AQUEM_DESTA_DATA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160437)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataUp_Click()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_BotaoDataUp_Click

    dtData = StrParaDate(Data.Text)

    'le os recebimentos selecionados
    lErro = CF("FluxoAnalitico_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 1)
    If lErro <> SUCESSO And lErro <> 133191 Then gError 133458

    If lErro = 133191 Then gError 133459

    Data.Text = Format(dtData, "dd/mm/yy")

    Call Botao_ExibeFluxo_Click

    Exit Sub

Erro_BotaoDataUp_Click:

    Select Case gErr

        Case 133458

        Case 133459
            Call Rotina_Erro(vbOKOnly, "NAO_HA_FLUXO_ALEM_DESTA_DATA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160438)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_FluxoReceb

End Sub

Sub Limpa_Tela_FluxoReceb()

    Call Grid_Limpa(objGrid1)
    Data.PromptInclude = False
    Data.Text = ""
    Data.PromptInclude = True

End Sub

Private Sub Data_Change()

    If objGrid1.iLinhasExistentes > 0 Then
        Call Grid_Limpa(objGrid1)
        TotalValor.Caption = ""
    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 20181
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 20181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160439)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sOrdenacao As String
Dim colFluxoAnalitico As New Collection
Dim dtData As Date

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then gError 20182

    Set objFluxo1 = objFluxo

    'le os pagamentos selecionados
    lErro = CF("FluxoAnalitico_Le", colFluxoAnalitico, sOrdenacao, objFluxo1.lFluxoId, objFluxo.dtData, FLUXOANALITICO_TIPOREG_RECEBTO)
    If lErro <> SUCESSO And lErro <> 20170 Then gError 133468
    
    If colFluxoAnalitico.Count = 0 Then
    
        dtData = objFluxo1.dtData

        'le os recebimentos selecionados
        lErro = CF("FluxoAnalitico_Le_ProxAnt", objFluxo1.lFluxoId, dtData, FLUXOANALITICO_TIPOREG_RECEBTO, 1)
        If lErro <> SUCESSO And lErro <> 133191 Then gError 133469

        If lErro = SUCESSO Then objFluxo1.dtData = dtData
    
    End If

    Data.Text = Format(objFluxo1.dtData, "dd/mm/yy")

    Ordenados.ListIndex = -1

    'seta a ordenacao Cliente + Filial como a ordenacao inicial e inicializa o grid
    For iIndice = 0 To Ordenados.ListCount - 1
        If Ordenados.ItemData(iIndice) = ORDENACAO_CLIENTEFILIAL Then
            Ordenados.ListIndex = iIndice
            Exit For
        End If
    Next

    Parent.Caption = "Fluxo de Caixa " & objFluxo.sFluxo & " - Recebimentos por Título"
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 20182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", gErr)

        Case 133468, 133469

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160440)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Monta_Ordenacao(sOrdenacao As String, Ordenacao As ComboBox) As Long
'monta a expressão de ordenação SQL

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Ordenacao

    Select Case Ordenacao.ItemData(Ordenacao.ListIndex)

        Case ORDENACAO_CLIENTEFILIAL

            sOrdenacao = " ORDER BY Fornecedor, Filial"

        Case ORDENACAO_SIGLADOCUMENTO

            sOrdenacao = " ORDER BY SiglaDocumento"

        Case ORDENACAO_TITULOPARCELA

            sOrdenacao = " ORDER BY NumTitulo, Item, NumParcela"

    End Select

    Monta_Ordenacao = SUCESSO

    Exit Function

Erro_Monta_Ordenacao:

    Monta_Ordenacao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160441)

    End Select

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
    objGrid1.colColuna.Add ("Cliente")
    objGrid1.colColuna.Add ("Filial")
    objGrid1.colColuna.Add ("Sigla")
    objGrid1.colColuna.Add ("Título")
    objGrid1.colColuna.Add ("Item")
    objGrid1.colColuna.Add ("Parcela")
    objGrid1.colColuna.Add ("Valor")
    objGrid1.colColuna.Add ("FilialEmpresa")

   'campos de edição do grid
    objGrid1.colCampo.Add (Cliente.Name)
    objGrid1.colCampo.Add (Filial.Name)
    objGrid1.colCampo.Add (SiglaDocumento.Name)
    objGrid1.colCampo.Add (NumTitulo.Name)
    objGrid1.colCampo.Add (Item.Name)
    objGrid1.colCampo.Add (Parcela.Name)
    objGrid1.colCampo.Add (Valor.Name)
    objGrid1.colCampo.Add (FilialEmpresa.Name)

    objGrid1.objGrid = GridFCaixa

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 10

    objGrid1.objGrid.ColWidth(0) = 300

    objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid1)

    'Posiciona o totalizador
    TotalValor.Top = GridFCaixa.Top + GridFCaixa.Height
    TotalValor.Left = GridFCaixa.Left

    For iIndice = 0 To GRID_VALOR_COL - 1
        TotalValor.Left = TotalValor.Left + GridFCaixa.ColWidth(iIndice) + GridFCaixa.GridLineWidth + 10
    Next

    TotalValor.Width = GridFCaixa.ColWidth(GRID_VALOR_COL)

    LabelTotais.Top = TotalValor.Top + (TotalValor.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalValor.Left - LabelTotais.Width - 50

    Inicializa_GridFCaixa = SUCESSO

End Function

Function Preenche_GridFCaixa(colFluxoAnalitico As Collection) As Long
'preenche o grid com os recebimentos contidos na coleção colFluxoAnalitico

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoAnalitico As ClassFluxoAnalitico
Dim dColunaSoma As Double

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoAnalitico.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 1
    Else
        objGrid1.objGrid.Rows = colFluxoAnalitico.Count + 1
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoAnalitico.Count

    dColunaSoma = 0

    'preenche o grid com os dados retornados na coleção colFluxoAnalitico
    For iIndice = 1 To colFluxoAnalitico.Count

        Set objFluxoAnalitico = colFluxoAnalitico.Item(iIndice)

        GridFCaixa.TextMatrix(iIndice, GRID_CLIENTE_COL) = objFluxoAnalitico.sNomeReduzido
        GridFCaixa.TextMatrix(iIndice, GRID_FILIAL_COL) = CStr(objFluxoAnalitico.iFilial)
        GridFCaixa.TextMatrix(iIndice, GRID_SIGLA_DOCUMENTO_COL) = objFluxoAnalitico.sSiglaDocumento
        
        If objFluxoAnalitico.sSiglaDocumento <> TIPODOC_CREDITOSRECCLI And objFluxoAnalitico.sSiglaDocumento <> TIPODOC_RECEBIMENTO_ANTECIPADO Then
            GridFCaixa.TextMatrix(iIndice, GRID_NUMTITULO_COL) = objFluxoAnalitico.sTitulo
            GridFCaixa.TextMatrix(iIndice, GRID_PARCELA_COL) = CStr(objFluxoAnalitico.iNumParcela)
        End If
        
        If objFluxoAnalitico.sSiglaDocumento = TIPODOC_CONTRATO_REC Then
            GridFCaixa.TextMatrix(iIndice, GRID_ITEM_COL) = CStr(objFluxoAnalitico.iItem)
        End If
        
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_COL) = Format(objFluxoAnalitico.dValor, "Standard")
        
'        If objFluxoAnalitico.sSiglaDocumento = TIPODOC_CONTRATO_REC Then
'            GridFCaixa.TextMatrix(iIndice, GRID_DATAREFENCIA_COL) = Format(objFluxoAnalitico.dtDataReferencia, "dd/mm/yyyy")
'        End If
        
        If objFluxoAnalitico.sSiglaDocumento = TIPODOC_PV Then
            GridFCaixa.TextMatrix(iIndice, GRID_FILIALEMPRESA_COL) = CStr(objFluxoAnalitico.iFilialEmpresa)
        End If
        
        dColunaSoma = dColunaSoma + objFluxoAnalitico.dValor

    Next

    TotalValor.Caption = Format(dColunaSoma, "Standard")

    Preenche_GridFCaixa = SUCESSO

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160442)

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
        If lErro <> SUCESSO Then Error 20184

        dtData = CDate(Data.Text)

        If dtData < objFluxo1.dtDataBase Or dtData > objFluxo1.dtDataFinal Then Error 20218

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 20184

        Case 20218
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATA_FORA_FAIXA", Err, CStr(dtData), CStr(objFluxo1.dtDataBase), CStr(objFluxo1.dtDataFinal))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160443)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub GridFCaixa_DblClick()

    Call BotaoDocOriginal_Click

End Sub

Private Sub Ordenados_Click()

Dim sOrdenacao As String
Dim lErro As Long
Dim colFluxoAnalitico As New Collection

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex >= 0 Then

        'se a data da tela não estiver preenchido ==> não exibe os dados no grid
        If Len(Data.ClipText) = 0 Then Exit Sub
    
        'monta a expressão SQL de Ordenação
        lErro = Monta_Ordenacao(sOrdenacao, Ordenados)
        If lErro <> SUCESSO Then Error 20185
    
        'le os recebimentos selecionados
        lErro = CF("FluxoAnalitico_Le", colFluxoAnalitico, sOrdenacao, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_RECEBTO)
        If lErro <> SUCESSO And lErro <> 20170 Then Error 20191
    
        If lErro = 20170 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXO_RECEB_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yy"), MAX_FLUXO)
    
        'preenche o grid com os recebimentos lidos
        lErro = Preenche_GridFCaixa(colFluxoAnalitico)
        If lErro <> SUCESSO Then Error 20192

    End If

    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 20185, 20191, 20192

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160444)

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
'faz a critica da celula do grid que está deixando de ser a corrente /m

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 20195

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 20195
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160445)

    End Select

    Exit Function

End Function

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoAnalitico As New ClassFluxoAnalitico
Dim colFluxoAnalitico As New Collection
Dim dtData As Date
Dim sOrdenados As String

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXORECEB_CPR")
    If lErro <> SUCESSO Then Error 47910
    
    'obter dados comuns a todas as linhas do grid
    sOrdenados = Ordenados.List(Ordenados.ListIndex)
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_FCaixa_Obter(colFluxoAnalitico)
    If lErro <> SUCESSO Then Error 47911
    
    For iIndice1 = 1 To colFluxoAnalitico.Count
    
        Set objFluxoAnalitico = colFluxoAnalitico.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(sOrdenados)
        Call colTemp.Add(dtData)
        Call colTemp.Add(objFluxoAnalitico.sNomeReduzido)
        Call colTemp.Add(objFluxoAnalitico.iFilial)
        Call colTemp.Add(objFluxoAnalitico.sSiglaDocumento)
        Call colTemp.Add(objFluxoAnalitico.sTitulo)
        Call colTemp.Add(objFluxoAnalitico.iItem)
        Call colTemp.Add(objFluxoAnalitico.iNumParcela)
        Call colTemp.Add(objFluxoAnalitico.dValor)
        Call colTemp.Add(objFluxoAnalitico.iFilialEmpresa)
        
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47912
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47913
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47910, 47911, 47912, 47913
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160446)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoAnalitico As Collection) As Long

Dim objFluxoAnalitico As ClassFluxoAnalitico
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objFluxoAnalitico = New ClassFluxoAnalitico
        
        objFluxoAnalitico.sNomeReduzido = GridFCaixa.TextMatrix(iLinha, GRID_CLIENTE_COL)
        objFluxoAnalitico.iFilial = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_FILIAL_COL))
        objFluxoAnalitico.sSiglaDocumento = GridFCaixa.TextMatrix(iLinha, GRID_SIGLA_DOCUMENTO_COL)
        objFluxoAnalitico.sTitulo = GridFCaixa.TextMatrix(iLinha, GRID_NUMTITULO_COL)
        objFluxoAnalitico.iItem = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_ITEM_COL))
        objFluxoAnalitico.iNumParcela = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_PARCELA_COL))
        objFluxoAnalitico.dValor = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_COL))
        objFluxoAnalitico.iFilialEmpresa = StrParaInt(GridFCaixa.TextMatrix(iLinha, GRID_FILIALEMPRESA_COL))
        
        colFluxoAnalitico.Add objFluxoAnalitico
        
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160447)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA_RECEBIMENTOS_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa - Recebimentos por Título"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoReceb"
    
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

Private Sub Label59_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label59, Source, X, Y)
End Sub

Private Sub Label59_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label59, Button, Shift, X, Y)
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

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objFluxoAnalitico As New ClassFluxoAnalitico
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber
Dim colFluxoAnalitico As New Collection
Dim sOrdenacao As String
Dim objContrato As New ClassContrato
Dim objItemContrato As New ClassItensDeContrato
Dim objFluxoContratoItemNFRec As New ClassFluxoContratoItemNFRec
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridFCaixa.Row = 0 Then Error 123510
        
    'Se foi selecionada uma linha que está preenchida
    If GridFCaixa.Row <= objGrid1.iLinhasExistentes Then
    
        'monta a expressão SQL de Ordenação
        lErro = Monta_Ordenacao(sOrdenacao, Ordenados)
        If lErro <> SUCESSO Then gError 123511
        
        'le os pagamentos selecionados
        lErro = CF("FluxoAnalitico_Le", colFluxoAnalitico, sOrdenacao, objFluxo1.lFluxoId, CDate(Data.Text), FLUXOANALITICO_TIPOREG_RECEBTO)
        If lErro <> SUCESSO And lErro <> 20170 Then gError 123512
        
        If lErro = 20170 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXO_PAG_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)
        
        'pega o objeto referente a linha selecionada no grid
        For Each objFluxoAnalitico In colFluxoAnalitico
            If objFluxoAnalitico.sTitulo = GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_NUMTITULO_COL) And _
            objFluxoAnalitico.sSiglaDocumento = GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SIGLA_DOCUMENTO_COL) Then
                If objFluxoAnalitico.sSiglaDocumento = TIPODOC_PV Then
                    If objFluxoAnalitico.iFilialEmpresa = GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_FILIALEMPRESA_COL) Then Exit For
                Else
                    Exit For
                End If
            End If
        Next
        
        If objFluxoAnalitico.sSiglaDocumento = TIPODOC_CONTRATO_REC Then
            
            sSelecao = "FluxoID = ? AND Contrato = ? AND SeqContrato = ? AND DataRec = ?"
            
            colSelecao.Add objFluxo1.lFluxoId
            colSelecao.Add objFluxoAnalitico.sTitulo
            colSelecao.Add objFluxoAnalitico.iItem
            colSelecao.Add objFluxoAnalitico.dtDataReferencia
            
            'Chama a tela de Browser
            Call Chama_Tela("FluxoContratoItemNFRecLista", colSelecao, Nothing, Nothing, sSelecao)
            
        ElseIf objFluxoAnalitico.sSiglaDocumento = TIPODOC_PV Then
        
            objPedidoVenda.iFilialEmpresa = objFluxoAnalitico.iFilialEmpresa
            objPedidoVenda.lCodigo = StrParaLong(objFluxoAnalitico.sTitulo)
            
            Call Chama_Tela("PedidoVenda", objPedidoVenda)
        
        Else
        
            objParcelaReceber.lNumIntDoc = objFluxoAnalitico.lNumIntDoc
            
            'Le o NumInterno do Titulo para passar no objParcelaReceber
            lErro = CF("ParcelaReceber_Le", objParcelaReceber)
            If lErro <> SUCESSO And lErro <> 19147 Then gError 123513
            If lErro <> SUCESSO Then
    
                'Se não encontrar
                lErro = CF("ParcelaReceber_Baixada_Le", objParcelaReceber)
                If lErro <> SUCESSO Then gError 123514
    
            End If
    
            objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
            
            'lê os dados do título
            lErro = CF("TituloReceber_Le", objTituloReceber)
            If lErro <> SUCESSO And lErro <> 26061 Then gError 123515
            
            If lErro <> SUCESSO Then
            
                'se não encontrar
                lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
                If lErro <> SUCESSO Then gError 123516
            
            End If
            
            Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
    
        End If
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr
    
        Case 123510
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
         
        Case 123511, 123512, 123513, 123515, 133356
        
        Case 123514
             Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", gErr)
         
        Case 123516
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case 133357
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCONTRATO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160448)

    End Select

    Exit Sub

End Sub

