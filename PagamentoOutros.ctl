VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoOutros 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   KeyPreview      =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   6135
   Begin VB.CommandButton BotaoIncluir 
      Caption         =   "(F6)  Incluir"
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
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   195
      Width           =   1350
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3255
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1725
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1725
   End
   Begin VB.ComboBox AdmMeioPagto 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   195
      Width           =   1875
   End
   Begin VB.Frame FrameOutros 
      Caption         =   "Outros"
      Height          =   2565
      Left            =   105
      TabIndex        =   0
      Top             =   1410
      Width           =   5910
      Begin MSMask.MaskEdBox AutorizacaoGrid 
         Height          =   255
         Left            =   4650
         TabIndex        =   18
         Top             =   675
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox AdmMeioPagtoGrid 
         Height          =   240
         Left            =   630
         TabIndex        =   13
         Top             =   735
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   240
         Left            =   3405
         TabIndex        =   6
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox ClienteGrid 
         Height          =   240
         Left            =   1995
         TabIndex        =   7
         Top             =   705
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid GridOutros 
         Height          =   1710
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   300
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   3016
         _Version        =   393216
         Rows            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Outros: "
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
         Left            =   2685
         TabIndex        =   15
         Top             =   2235
         Width           =   1185
      End
      Begin VB.Label TotalOutros 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3855
         TabIndex        =   14
         Top             =   2190
         Width           =   1230
      End
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1605
      TabIndex        =   2
      Top             =   675
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   675
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
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
   Begin MSMask.MaskEdBox Autorizacao 
      Height          =   300
      Left            =   3840
      TabIndex        =   16
      Top             =   1095
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Autorização:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   17
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   3105
      TabIndex        =   12
      Top             =   705
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
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
      Left            =   1035
      TabIndex        =   9
      Top             =   705
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Administradora:"
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
      Left            =   225
      TabIndex        =   8
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "PagamentoOutros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Dim iAlterado As Integer
Dim giTipo As Integer

'Variável que guarda as características do grid da tela
Dim objGridOutros As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_AdmMeioPagto_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Autorizacao_Col As Integer

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim colAdmMeioPagto As New Collection
    
    giTipo = MOVIMENTOCAIXA_RECEB_OUTROS
    
    'Se o projeto <> SGEECF
    If gsNomePrinc <> "SGEECF" Then
        Call CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
        Set gcolAdmMeioPagto = colAdmMeioPagto
        giTipo = MOVIMENTOCAIXA_RECEB_CARNE_OUTROS
    End If
    
    'Adiciona na combo de AdmMeioPagto todos
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_OUTROS And objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
            AdmMeioPagto.AddItem objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
            AdmMeioPagto.ItemData(AdmMeioPagto.NewIndex) = objAdmMeioPagto.iCodigo
        End If
    Next
    
    Set gobjVenda = objVenda
    
    'Joga na tela todos os dados referentes a Contra-vale, Convenio e Vale
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = giTipo And objMovimento.iAdmMeioPagto <> 0 Then
        
            objGridOutros.iLinhasExistentes = objGridOutros.iLinhasExistentes + 1
            
            GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Valor_Col) = Format(objMovimento.dValor, "standard")
            GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Cliente_Col) = objMovimento.sFavorecido
            GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Autorizacao_Col) = objMovimento.sAutorizacao
            
            For iIndice = 0 To AdmMeioPagto.ListCount - 1
                If AdmMeioPagto.ItemData(iIndice) = objMovimento.iAdmMeioPagto Then
                    GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_AdmMeioPagto_Col) = AdmMeioPagto.List(iIndice)
                    Exit For
                End If
            Next
        End If
    Next
        
    'Atualiza o total do troco
    Call Atualiza_Total
        
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
    
    Set objGridOutros = New AdmGrid
        
    Call Inicializa_Grid_Outros(objGridOutros)
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Function Inicializa_Grid_Outros(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Administradora")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Autorização")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AdmMeioPagtoGrid.Name)
    objGridInt.colCampo.Add (ClienteGrid.Name)
    objGridInt.colCampo.Add (ValorGrid.Name)
    objGridInt.colCampo.Add (AutorizacaoGrid.Name)
    
    'Colunas do Grid
    iGrid_AdmMeioPagto_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Autorizacao_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridOutros

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_OUTROS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridOutros.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Outros = SUCESSO

    Exit Function

End Function


Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoIncluir_Click
    
    'se AdmMeioPagto não selecionado --> Erro.
    If AdmMeioPagto.ListIndex = -1 Then gError 99653
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99655
    
    'verifica se o valor pago ultrapassa o valor minimo da condicao de pagto
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iCodigo = AdmMeioPagto.ItemData(AdmMeioPagto.ListIndex) Then
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                    If StrParaDbl(Valor.Text) < objAdmMeioPagtoCondPagto.dValorMinimo Then gError 126815
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    objGridOutros.iLinhasExistentes = objGridOutros.iLinhasExistentes + 1
    
    GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_AdmMeioPagto_Col) = AdmMeioPagto.Text
    GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Cliente_Col) = Cliente.Text
    GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridOutros.TextMatrix(objGridOutros.iLinhasExistentes, iGrid_Autorizacao_Col) = Autorizacao.Text
    
    'Atualiza o total do troco
    Call Atualiza_Total
        
    'Limpa os campos da tela
    Valor.Text = ""
    AdmMeioPagto.ListIndex = -1
    Cliente.Text = ""
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 99653
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMMEIOPAGTO_NAO_SELECIONADO, gErr)
            
        Case 99655
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case 126815
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORMINIMO_CONDPAGTO, gErr, objAdmMeioPagtoCondPagto.dValorMinimo, Valor.Text)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164188)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalOutros.Caption = ""
    
    For iIndice = 1 To objGridOutros.iLinhasExistentes
        TotalOutros.Caption = Format(StrParaDbl(TotalOutros.Caption) + StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_Valor_Col)), "standard")
    Next
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objAdm As ClassAdmMeioPagto
Dim iTipo As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoOk_Click
    
    'Exclui todos os movimentos de Outros
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = giTipo And objMovimento.iAdmMeioPagto <> 0 Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada linha do grid...
    For iIndice = 1 To objGridOutros.iLinhasExistentes
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iAdmMeioPagto = Codigo_Extrai(GridOutros.TextMatrix(iIndice, iGrid_AdmMeioPagto_Col))
        For Each objAdmMeioPagtoCondPagto In gcolOutros
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimento.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                objMovimento.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                Exit For
            End If
        Next
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_Valor_Col))
        objMovimento.dHora = CDbl(Time)
        objMovimento.sFavorecido = GridOutros.TextMatrix(iIndice, iGrid_Cliente_Col)
        objMovimento.iTipo = giTipo
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        objMovimento.sAutorizacao = GridOutros.TextMatrix(iIndice, iGrid_Autorizacao_Col)
        
        
        If objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_CONTRAVALE Then
            objMovimento.iTipoCartao = TIPO_MANUAL
        Else
            objMovimento.iTipoCartao = TIPO_POS
        End If
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164189)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99656
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99656
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164190)

    End Select

    Exit Sub
    
End Sub

Private Sub GridOutros_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOutros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridOutros, iAlterado)
    End If

End Sub

Private Sub GridOutros_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridOutros, iAlterado)

End Sub

Private Sub GridOutros_GotFocus()

    Call Grid_Recebe_Foco(objGridOutros)

End Sub

Private Sub GridOutros_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim dValor As Double

    If GridOutros.Row <> 0 Then
        dValor = StrParaDbl(GridOutros.TextMatrix(GridOutros.Row, iGrid_Valor_Col))
    End If
    
    Call Grid_Trata_Tecla1(KeyCode, objGridOutros)
    
    If KeyCode = vbKeyDelete Then
        TotalOutros.Caption = Format(StrParaDbl(TotalOutros.Caption) - dValor, "standard")
    End If
    
End Sub

Private Sub GridOutros_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOutros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOutros, iAlterado)
    End If
        
End Sub

Private Sub GridOutros_LeaveCell()

    Call Saida_Celula(objGridOutros)

End Sub

Private Sub GridOutros_LostFocus()

    Call Grid_Libera_Foco(objGridOutros)

End Sub

Private Sub GridOutros_RowColChange()

    Call Grid_RowColChange(objGridOutros)

End Sub

Private Sub GridOutros_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridOutros)
        
End Sub

Private Sub GridOutros_Scroll()

    Call Grid_Scroll(objGridOutros)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99659

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99659
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164191)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a referência da tela
    Set gobjVenda = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not gobjVenda Is Nothing Then
    
    Select Case KeyCode
    
    
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
            Call BotaoOk_Click

        Case vbKeyEscape
            If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
            Call BotaoCancelar_Click

        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
            Call BotaoIncluir_Click
    
        Case vbKeyF7
            GridOutros.SetFocus
            
    End Select
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pagamentos em Outros Meios"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoOutros"
    
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

