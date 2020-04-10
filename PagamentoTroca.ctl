VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoTroca 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   6630
   Begin VB.Frame FrameTroca 
      Caption         =   "Produtos Trocados"
      Height          =   2625
      Left            =   90
      TabIndex        =   14
      Top             =   960
      Width           =   6390
      Begin MSMask.MaskEdBox TotalGrid 
         Height          =   240
         Left            =   4140
         TabIndex        =   8
         Top             =   900
         Width           =   1245
         _ExtentX        =   2196
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   240
         Left            =   3135
         TabIndex        =   7
         Top             =   870
         Width           =   975
         _ExtentX        =   1720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantidadeGrid 
         Height          =   240
         Left            =   1905
         TabIndex        =   6
         Top             =   825
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoGrid 
         Height          =   240
         Left            =   585
         TabIndex        =   5
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTroca 
         Height          =   1860
         Left            =   225
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   5
         Cols            =   4
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
      Begin VB.Label TotalTroca 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4410
         TabIndex        =   16
         Top             =   2175
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Troca: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3225
         TabIndex        =   15
         Top             =   2235
         Width           =   1125
      End
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
      Height          =   360
      Left            =   1125
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3735
      Width           =   1725
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
      Height          =   360
      Left            =   3240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3735
      Width           =   1725
   End
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
      Left            =   5145
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1350
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   300
      Left            =   1215
      TabIndex        =   0
      Top             =   165
      Width           =   2955
      _ExtentX        =   5212
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
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   300
      Left            =   1215
      TabIndex        =   1
      Top             =   630
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   2625
      TabIndex        =   2
      Top             =   630
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label LabelTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4935
      TabIndex        =   18
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      Left            =   4380
      TabIndex        =   17
      Top             =   675
      Width           =   510
   End
   Begin VB.Label Label1 
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
      Index           =   4
      Left            =   2070
      TabIndex        =   13
      Top             =   675
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
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
      Index           =   5
      Left            =   105
      TabIndex        =   10
      Top             =   675
      Width           =   1050
   End
   Begin VB.Label LabelProdutoBrowse 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
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
      Left            =   435
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   195
      Width           =   735
   End
End
Attribute VB_Name = "PagamentoTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Public iAlterado As Integer
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Variável que guarda as características do grid da tela
Dim objGridTroca As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Total_Col As Integer


Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim sServ As String
Dim objTroca As ClassTroca
Dim lErro As Long
Dim objProduto As ClassProduto

    Set gobjVenda = objVenda
    
    'Joga na tela todas as Trocas
    For Each objTroca In gobjVenda.colTroca
        
        objGridTroca.iLinhasExistentes = objGridTroca.iLinhasExistentes + 1
            
        GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Quantidade_Col) = objTroca.dQuantidade
        GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Valor_Col) = Format(objTroca.dValor / objTroca.dQuantidade, "standard")
        GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Total_Col) = Format(objTroca.dValor, "standard")
        GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Produto_Col) = objTroca.sProduto
        
    Next
        
    'Atualiza o total da troca
    Call Atualiza_Total
        
    Trata_Parametros = SUCESSO

    Exit Function
    
Erro_Trata_Parametros:
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164208)

    End Select
    
    Exit Function
        
End Function

Public Sub Form_Load()
    
Dim lErro As Long

    Set objEventoProduto = New AdmEvento
    
    Set objGridTroca = New AdmGrid
        
    Call Inicializa_Grid_Troca(objGridTroca)
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:
    
    Select Case gErr
        
        Case 99673
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164209)

    End Select
    
    Exit Sub

End Sub

Function Inicializa_Grid_Troca(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Total")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (ProdutoGrid.Name)
    objGridInt.colCampo.Add (QuantidadeGrid.Name)
    objGridInt.colCampo.Add (ValorGrid.Name)
    objGridInt.colCampo.Add (TotalGrid.Name)
    
    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Quantidade_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Total_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridTroca

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_TROCA + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridTroca.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Troca = SUCESSO

    Exit Function

End Function

Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoIncluir_Click
    
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99660
    
    'Se quantidade não preenchido --> Erro.
    If Len(Trim(Quantidade.Text)) = 0 Then gError 105799
    
    'Se quantidade não preenchido --> Erro.
    If Len(Trim(Produto.Text)) = 0 Then gError 105800
    
    objGridTroca.iLinhasExistentes = objGridTroca.iLinhasExistentes + 1
        
    'Se a quantidade não foi preenchida --> valor default
    If Len(Trim(Quantidade.Text)) = 0 Then Quantidade.Text = 1
    
    GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Produto_Col) = Produto.Text
    GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Quantidade_Col) = Quantidade.Text
    GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridTroca.TextMatrix(objGridTroca.iLinhasExistentes, iGrid_Total_Col) = Format(Valor.Text * Quantidade.Text, "standard")
    
    'Atualiza o total do troco
    Call Atualiza_Total
        
    'Limpa os campos da tela
    Call Limpa_Tela(Me)
    LabelTotal.Caption = ""
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr
            
        Case 99660
            Call Rotina_ErroECF(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO2", gErr)
           
        Case 105799
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_NAO_PREENCHIDO1, gErr)
           
        Case 105800
            Call Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_PREENCHIDO1, gErr)

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164210)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalTroca.Caption = ""
    
    For iIndice = 1 To objGridTroca.iLinhasExistentes
        TotalTroca.Caption = Format(StrParaDbl(TotalTroca.Caption) + (StrParaDbl(GridTroca.TextMatrix(iIndice, iGrid_Total_Col))), "standard")
    Next
    
End Sub

Private Sub BotaoOK_Click()

Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objTroca As ClassTroca
Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoOK_Click
    
    If Not gobjVenda Is Nothing Then
    
    'Exclui todos os movimentos de Troca
    Set gobjVenda.colTroca = New Collection
            
    'Exclui todos os movimentos de Troca
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_TROCA Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
            
    'Para cada linha do grid...
    For iIndice = 1 To objGridTroca.iLinhasExistentes
            
        Set objTroca = New ClassTroca
    
        'Insere um novo movimento
        objTroca.dValor = StrParaDbl(GridTroca.TextMatrix(iIndice, iGrid_Total_Col))
        objTroca.iFilialEmpresa = giFilialEmpresa
        objTroca.dQuantidade = StrParaDbl(GridTroca.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objTroca.sProduto = GridTroca.TextMatrix(iIndice, iGrid_Produto_Col)
        
        Set objProduto = gaobjProdutosNome.Busca(objTroca.sProduto)
        
        objTroca.sCodProduto = objProduto.sCodigo
        objTroca.sUnidadeMed = objProduto.sSiglaUMVenda
        
        gobjVenda.colTroca.Add objTroca
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_TROCA
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_TROCA
        objMovimento.iParcelamento = COD_A_VISTA
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridTroca.TextMatrix(iIndice, iGrid_Total_Col))
        objMovimento.dHora = CDbl(Time)
        objMovimento.lNumRefInterna = gobjVenda.colTroca.Count
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    End If
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164211)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoBrowse_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoBrowse_Click
    
    Call Chama_TelaECF_Modal("ProdutosLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_LabelProdutoBrowse_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164212)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(objProduto.sReferencia)) > 0 Then
        Produto.Text = objProduto.sReferencia
    Else
        Produto.Text = objProduto.sCodigoBarras
    End If
    Call Produto_Validate(False)

'    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214935)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_Produto_Validate
    
    'Se o produto não está preenchido
    If Len(Trim(Produto.Text)) <> 0 Then
    
        sProduto = Produto.Text
    
        lErro = TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
        If lErro <> SUCESSO Then gError 112079
        If Not (objProduto Is Nothing) Then
            Produto.Text = objProduto.sNomeReduzido
            Valor.Text = objProduto.dPrecoLoja
        End If
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr
                
        Case 112079
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164213)

    End Select
    
    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99661
        
        'Recalcula o valor total
        If Len(Trim(Quantidade.Text)) > 0 Then
            LabelTotal.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(Valor.Text), "Standard")
        End If
    
    Else
        LabelTotal.Caption = ""
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99661
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164214)

    End Select

    Exit Sub
    
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Quantidade_Validate
    
    If Len(Trim(Quantidade.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 99662
        
        'Recalcula o Quantidade total
        If Len(Trim(Quantidade.Text)) > 0 Then
            LabelTotal.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(Valor.Text), "Standard")
        End If
        
    Else
        LabelTotal.Caption = ""
        
    End If
        
    Exit Sub
    
Erro_Quantidade_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99662
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164215)

    End Select

    Exit Sub
    
End Sub

Private Sub GridTroca_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridTroca, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridTroca, iAlterado)
    End If

End Sub

Private Sub GridTroca_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridTroca, iAlterado)

End Sub

Private Sub GridTroca_GotFocus()

    Call Grid_Recebe_Foco(objGridTroca)

End Sub

Private Sub GridTroca_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Call Grid_Trata_Tecla1(KeyCode, objGridTroca)
    
    Call Atualiza_Total
    
End Sub

Private Sub GridTroca_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTroca, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTroca, iAlterado)
    End If
        
End Sub

Private Sub GridTroca_LeaveCell()

    Call Saida_Celula(objGridTroca)

End Sub

Private Sub GridTroca_LostFocus()

    Call Grid_Libera_Foco(objGridTroca)

End Sub

Private Sub GridTroca_RowColChange()

    Call Grid_RowColChange(objGridTroca)

End Sub

Private Sub GridTroca_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridTroca)
        
End Sub

Private Sub GridTroca_Scroll()

    Call Grid_Scroll(objGridTroca)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99663

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99663
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164216)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a referência da tela
    Set gobjVenda = Nothing
    Set objEventoProduto = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not gobjVenda Is Nothing Then
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call LabelProdutoBrowse_Click
        End If
    End If
    
    'Clique em f5
    If KeyCode = vbKeyF5 Then
        If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
        Call BotaoOK_Click
    End If

    'Clique em esc
    If KeyCode = vbKeyEscape Then
        If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
        Call BotaoCancelar_Click
    End If

    'Clique em F6
    If KeyCode = vbKeyF6 Then
        If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
        Call BotaoIncluir_Click
    End If
    
    If KeyCode = vbKeyF7 Then
        GridTroca.SetFocus
    End If
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pagamentos em Troca"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoTroca"
    
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
