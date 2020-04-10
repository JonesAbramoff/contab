VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LocalizacaoProdutoOcx 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ScaleHeight     =   5760
   ScaleWidth      =   6750
   Begin VB.ComboBox Produto 
      Height          =   315
      Left            =   1725
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   825
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "LocalizacaoProdutoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "LocalizacaoProdutoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "LocalizacaoProdutoOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoSubstituir 
      Caption         =   "Substituição do Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4125
      TabIndex        =   5
      Top             =   5250
      Width           =   2430
   End
   Begin VB.Frame Frame7 
      Caption         =   "Localização do Produto"
      Height          =   2760
      Left            =   120
      TabIndex        =   10
      Top             =   2295
      Width           =   6420
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   225
         Left            =   1485
         TabIndex        =   1
         Top             =   495
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantDisponivel 
         Height          =   225
         Left            =   2895
         TabIndex        =   2
         Top             =   510
         Width           =   1575
         _ExtentX        =   2778
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantAlocada 
         Height          =   225
         Left            =   4485
         TabIndex        =   3
         Top             =   510
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   "FORMATO_ESTOQUE"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
         Height          =   1860
         Left            =   390
         TabIndex        =   4
         Top             =   420
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantTotalReserva 
         AutoSize        =   -1  'True
         Caption         =   "Quant. Alocada:"
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
         Left            =   1845
         TabIndex        =   12
         Top             =   2295
         Width           =   1395
      End
      Begin VB.Label TotalAlocado 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3375
         TabIndex        =   11
         Top             =   2265
         Width           =   1440
      End
   End
   Begin VB.ComboBox Item 
      Height          =   315
      Left            =   1725
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   345
      Width           =   750
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1725
      TabIndex        =   19
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unidade:"
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
      Left            =   885
      TabIndex        =   18
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quant. a Alocar:"
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
      Left            =   255
      TabIndex        =   17
      Top             =   1800
      Width           =   1410
   End
   Begin VB.Label QuantAlocar 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1725
      TabIndex        =   16
      Top             =   1755
      Width           =   1440
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3420
      TabIndex        =   15
      Top             =   825
      Width           =   3120
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   930
      TabIndex        =   14
      Top             =   870
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item:"
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
      Left            =   1230
      TabIndex        =   13
      Top             =   390
      Width           =   435
   End
End
Attribute VB_Name = "LocalizacaoProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
 
Dim iAlterado As Integer
Dim gcolItemNF As ColItensNF
Dim giCodigo As Integer

Dim objGridAlocacao As AdmGrid
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_QuantDisp_Col As Integer
Dim iGrid_QuantAloc_Col As Integer

Private Sub Almoxarifado_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_Alocacao

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 39394

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162397)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSubstituir_Click()

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objSubstProdutoNF As New ClassSubstProdutoNF
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dQuantAlocar As Double
Dim dFator As Double
Dim colEstoqueProduto As New colEstoqueProduto
Dim objEstoqueProduto As ClassEstoqueProduto
Dim iItem As Integer
Dim sProduto As String
Dim iEscaninho As Integer

On Error GoTo Erro_BotaoSubstituir_Click
    
    'Verifica se há algum item selecionado na combo de itens
    If Item.ListIndex = -1 Then Exit Sub
    
    iItem = Item.ListIndex
    'Recolhe o item
    Set objItemNF = gcolItemNF(Item.ListIndex + 1)
    'Guarda os outros produto que já participam da nota
    For iIndice = 1 To gcolItemNF.Count
        If gcolItemNF.Item(iIndice).iItem <> objItemNF.iItem Then
            objSubstProdutoNF.colOutrosProdutosNF.Add gcolItemNF.Item(iIndice).sProduto
        End If
    Next
    
    objSubstProdutoNF.sProduto = objItemNF.sProduto
    'Chama a tela de substituição de produto
    Call Chama_Tela_Modal("SubstProdutoNF", objSubstProdutoNF, giCodigo)
    'Se não substitui --> sia da rotina
    If giRetornoTela <> vbOK Then Exit Sub
    
    objProduto.sCodigo = objSubstProdutoNF.sProdutoSubstituto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 39384
    If lErro = 28030 Then Error 39385
    'Guarda os dados no novo produto
    objItemNF.sProduto = objProduto.sCodigo
    objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
    objItemNF.sDescricaoItem = objProduto.sDescricao
    objItemNF.iClasseUM = objProduto.iClasseUM
    
    dQuantAlocar = StrParaDbl(QuantAlocar.Caption)
    
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUMEstoque, objItemNF.sUnidadeMed, dFator)
    If lErro <> SUCESSO Then Error 39386
    'Passa quantidade a alocar para UM Estoque
    objItemNF.dQuantidade = dQuantAlocar * dFator
    
    Set objItemNF.colAlocacoes = New ColAlocacoesItemNF
    'Lê os estoque do produto
    lErro = CF("EstoquesProduto_Le_Filial", objItemNF.sProduto, colEstoqueProduto)
    If lErro <> SUCESSO Then Error 39387
    
    lErro = TipoDocInfo_Escaninho(giCodigo, iEscaninho)
    If lErro <> SUCESSO Then Error 39387

    If iEscaninho = ESCANINHO_CONSIG_NOSSO Then

        For Each objEstoqueProduto In colEstoqueProduto
            iIndice = iIndice + 1
            If objEstoqueProduto.dQuantConsig = 0 Then
                colEstoqueProduto.Remove (iIndice)
                iIndice = iIndice - 1
            End If
        Next

    ElseIf iEscaninho = ESCANINHO_BENEF_3 Then 'Incluido por Leo em 11/01/02
    
        For Each objEstoqueProduto In colEstoqueProduto
            iIndice = iIndice + 1
            If objEstoqueProduto.dQuantBenef3 = 0 Then
                colEstoqueProduto.Remove (iIndice)
                iIndice = iIndice - 1
            End If
        Next 'Leo até aqui
    
    Else
    
        For Each objEstoqueProduto In colEstoqueProduto
            iIndice = iIndice + 1
            If objEstoqueProduto.dQuantDisponivel = 0 Then
                colEstoqueProduto.Remove (iIndice)
                iIndice = iIndice - 1
            End If
        Next
    
    End If
    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 39013
    
    'Coloca o produto na tela
    Produto.List(Produto.ListIndex) = sProduto

    Call GridAlocacao_Limpa
    'Carrega a tela com os dados do novo produto
    lErro = Preenche_Tela2(objItemNF, colEstoqueProduto, objProduto.sSiglaUMEstoque)
    If lErro <> SUCESSO Then Error 39388

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_BotaoSubstituir_Click:
    
    Select Case Err
    
        Case 39013, 39384, 39386, 39387, 39388
        
        Case 39385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    QuantDisponivel.Format = FORMATO_ESTOQUE
    QuantAlocada.Format = FORMATO_ESTOQUE
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err
        
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162398)
            
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variaveis globais
    Set gcolItemNF = Nothing
    Set objGridAlocacao = Nothing
    
End Sub

Private Sub GridAlocacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAlocacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Private Sub GridAlocacao_EnterCell()

    Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)

End Sub

Private Sub GridAlocacao_GotFocus()

    Call Grid_Recebe_Foco(objGridAlocacao)

End Sub

Private Sub GridAlocacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAlocacao, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Private Sub GridAlocacao_LeaveCell()

    Call Saida_Celula(objGridAlocacao)

End Sub

Private Sub GridAlocacao_Validate(Cancel As Boolean)
        
    Call Grid_Libera_Foco(objGridAlocacao)
        
End Sub

Private Sub GridAlocacao_RowColChange()

    Call Grid_RowColChange(objGridAlocacao)

End Sub

Private Sub GridAlocacao_Scroll()

    Call Grid_Scroll(objGridAlocacao)

End Sub

Private Sub Item_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Item_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim colEstoqueProduto As New colEstoqueProduto
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim objItemPV As New ClassItemPedido
Dim colReserva As New colReservaItem
Dim objReservaItem As ClassReservaItem
Dim vbMsg As VbMsgBoxResult
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim sProdutoMascarado As String

On Error GoTo Erro_Item_Click

    
    'Limpa a COmbo de produtos
    Produto.Clear
    
    If Item.ListIndex = -1 Then Exit Sub

    'Verificar se o item anterior sofreu alguma alteração
    If iAlterado = REGISTRO_ALTERADO Then
        'Verificar se o usuário deseja salvar as alterações
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ITEM_ANTERIOR_ALTERADO")
                
        If vbMsg = vbYes Then
            Call BotaoGravar_Click
        End If

    End If

    'Se o produto for de Grade
    If gcolItemNF(Item.ListIndex + 1).iPossuiGrade = MARCADO Then
        
        BotaoSubstituir.Enabled = False

        For Each objItemRomaneio In gcolItemNF(Item.ListIndex + 1).colItensRomaneioGrade
        
            'Preenche dados do Item Pedido
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 23800
            
            Produto.AddItem sProdutoMascarado
        
        Next

    Else
    
        'Preenche dados do Item Pedido
        
        BotaoSubstituir.Enabled = True
        
        lErro = Mascara_MascararProduto(gcolItemNF(Item.ListIndex + 1).sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 23800
        
        Produto.AddItem sProdutoMascarado
    
    End If

    Produto.ListIndex = 0
    
    iAlterado = 0

    
    Exit Sub
    
Erro_Item_Click:

    Select Case Err

        Case 23800

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162399)

    End Select
        
    Exit Sub
    
End Sub

Private Sub Produto_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim colEstoqueProduto As New colEstoqueProduto
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim objItemPV As New ClassItemPedido
Dim colReserva As New colReservaItem
Dim objReservaItem As ClassReservaItem
Dim sProduto As String
Dim lNumIntItemPedVenda As Long
Dim iEscaninho As Integer

On Error GoTo Erro_Item_Click

    Set objItemNF = gcolItemNF.Item(Item.ListIndex + 1)
    
    If objItemNF.iPossuiGrade <> MARCADO Then
        sProduto = objItemNF.sProduto
        lNumIntItemPedVenda = objItemNF.lNumIntItemPedVenda
    Else
        sProduto = objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1).sProduto
        lNumIntItemPedVenda = objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1).lNumIntItemPV
    End If
    
    'Lê os estoque do produto
    lErro = CF("EstoquesProduto_Le_Filial", sProduto, colEstoqueProduto)
    If lErro <> SUCESSO Then Error 39373
    
    'Se o itemNF foi gerado por um itemPV
    If objItemNF.lNumIntItemPedVenda > 0 Then
        objItemPV.lNumIntDoc = lNumIntItemPedVenda
        objItemPV.sProduto = sProduto
        objItemPV.iPossuiGrade = objItemNF.iPossuiGrade
        
        'Lê as reservas do item do pedido
        lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReserva)
        If lErro <> SUCESSO And lErro <> 51601 Then Error 51596

    End If
    
    iIndice = 0
    
    lErro = TipoDocInfo_Escaninho(giCodigo, iEscaninho)
    If lErro <> SUCESSO Then Error 51596
    
    'Exclui almoxarifados sem estoque disponível
    For Each objEstoqueProduto In colEstoqueProduto
        iIndice = iIndice + 1
        'Verifica se há uma reserva para esse item nesse almoxarifado
        'e inclui a qtd reservada como disponivel p\ esse pedido
        Call Procura_Almoxarifado(objEstoqueProduto, colReserva)
        'SE não tiver quantidade disponível e nem reserva nesse almoxarifado
        
        Select Case iEscaninho
        
            Case ESCANINHO_CONSIG_NOSSO
        
                If objEstoqueProduto.dQuantConsig <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
        
            Case ESCANINHO_BENEF_3
                
                If objEstoqueProduto.dQuantBenef3 <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
            
            Case Else
        
                If objEstoqueProduto.dQuantDisponivel <= 0 And gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
                    'Retira o estoque produto da coleção
                    colEstoqueProduto.Remove (iIndice)
                    iIndice = iIndice - 1
                End If
        
        End Select
        
    Next
    
    objProduto.sCodigo = sProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 39374
    If lErro = 28030 Then Error 39375
    
    objItemNF.iClasseUM = objProduto.iClasseUM
    objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
    
    Call GridAlocacao_Limpa
    'Prenche a tela com os dados do estoque do produto
    lErro = Preenche_Tela(objItemNF, colEstoqueProduto)
    If lErro <> SUCESSO Then Error 39376
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Item_Click:

    Select Case Err
    
        Case 39373, 39374, 51596
        
        Case 39375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objItemNF.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162400)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub QuantAlocada_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub QuantAlocada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAlocacao)

End Sub

Private Sub QuantAlocada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAlocacao)

End Sub

Private Sub QuantAlocada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAlocacao.objControle = QuantAlocada
    lErro = Grid_Campo_Libera_Foco(objGridAlocacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantDisponivel_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Inicializa_Grid_Alocacao(objGridInt As AdmGrid, lNumAlmoxarifados As Long) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer
Dim iEscaninho As Integer

    Set objGridAlocacao.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Almoxarifado")
    
    Call TipoDocInfo_Escaninho(giCodigo, iEscaninho)
    
    If iEscaninho = ESCANINHO_CONSIG_NOSSO Then
        objGridInt.colColuna.Add ("Quant. Consignada")
    
    ElseIf iEscaninho = ESCANINHO_BENEF_3 Then 'inserido por Leo em 15/01/02
        objGridInt.colColuna.Add ("Quant. Benef. 3º")
    
    Else
        objGridInt.colColuna.Add ("Quant. Disponivel")
    
    End If
    objGridInt.colColuna.Add ("Quant. Alocada")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (QuantDisponivel.Name)
    objGridInt.colCampo.Add (QuantAlocada.Name)

    'Colunas da Grid
    iGrid_Almoxarifado_Col = 1
    iGrid_QuantDisp_Col = 2
    iGrid_QuantAloc_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridAlocacao

    objGridInt.iLinhasVisiveis = 7
    If lNumAlmoxarifados > 7 Then
        objGridInt.objGrid.Rows = lNumAlmoxarifados + 1
    Else
        objGridInt.objGrid.Rows = 8
    End If

    'Largura da primeira coluna
    GridAlocacao.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridAlocacao)

    'Posiciona o totalizador e seu label
    TotalAlocado.top = GridAlocacao.top + GridAlocacao.Height
    TotalAlocado.left = GridAlocacao.left
    For iIndice = 0 To iGrid_QuantAloc_Col - 1
        TotalAlocado.left = TotalAlocado.left + GridAlocacao.ColWidth(iIndice) + GridAlocacao.GridLineWidth + 20
    Next

    TotalAlocado.Width = GridAlocacao.ColWidth(iGrid_QuantAloc_Col)

    QuantTotalReserva.top = TotalAlocado.top + (TotalAlocado.Height - QuantTotalReserva.Height) / 2
    QuantTotalReserva.left = TotalAlocado.left - QuantTotalReserva.Width

    Inicializa_Grid_Alocacao = SUCESSO

    Exit Function

End Function


Public Function Trata_Parametros(colItemNF As ColItensNF, iCodigo As Integer) As Long

Dim lNumAlmoxarifados As Long
Dim objItemNF As ClassItemNF
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giCodigo = iCodigo

    'Lê quantos Almoxarifados tem na filialempresa
    lErro = CF("AlmoxarifadosFilial_Le_Quantidade", giFilialEmpresa, lNumAlmoxarifados)
    If lErro <> SUCESSO Then Error 39369
    
    If lNumAlmoxarifados = 0 Then Error 39370
    
    Set objGridAlocacao = New AdmGrid
    'Inicializa o grid de alocacoes
    lErro = Inicializa_Grid_Alocacao(objGridAlocacao, lNumAlmoxarifados)
    If lErro <> SUCESSO Then Error 39371

    Set gcolItemNF = colItemNF
    
    For Each objItemNF In colItemNF
        Item.AddItem objItemNF.iItem
    Next
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 39369, 39371
        
        Case 39371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ALMOXARIFADO_FILIAL", Err, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162401)
            
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function


Private Sub GridAlocacao_Limpa()
    
    Call Grid_Limpa(objGridAlocacao)
    TotalAlocado.Caption = ""

End Sub

Private Function Preenche_Tela(objItemNF As ClassItemNF, colEstoqueProduto As colEstoqueProduto) As Long


Dim lErro As Long
Dim dFator As Double
Dim objEstoqueProduto As ClassEstoqueProduto
Dim iIndice As Integer
Dim dQuantAlocada As Double
Dim objAlocacao As ClassItemNFAlocacao
Dim dTotalAlocado As Double
Dim sProduto As String
Dim dQuantAlocar As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim sUMEstoque As String
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_Tela

    lErro = Mascara_MascararProduto(objItemNF.sProduto, sProduto)
    If lErro <> SUCESSO Then Error 22222
    
    objProduto.sCodigo = objItemNF.sProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 22222
    
    If objItemNF.iPossuiGrade = MARCADO Then
        
        Set objItemRomaneio = objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1)
        
        Descricao.Caption = objItemRomaneio.sDescricao
        dQuantAlocar = objItemRomaneio.dQuantidade
        sUMEstoque = objItemRomaneio.sUMEstoque
        
    Else
        Descricao.Caption = objItemNF.sDescricaoItem
        dQuantAlocar = objItemNF.dQuantidade
        sUMEstoque = objItemNF.sUMEstoque
    
    End If

    If objProduto.iKitVendaComp <> MARCADO Then
        lErro = CF("UM_Conversao_Trans", objItemNF.iClasseUM, objItemNF.sUnidadeMed, sUMEstoque, dFator)
        If lErro <> SUCESSO Then Error 39377
    Else
        dFator = 1
    End If

    '###################################################################
    'Alterado por Wagner 16/11/04
'    dQuantAlocar = dQuantAlocar * dFator
'
'    If StrParaDbl(Formata_Estoque(dQuantAlocar)) < dQuantAlocar Then
'
'        'Calcula o número de casas decimais do Formato de Estoque
'        iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'        dQuantAlocar = dQuantAlocar + dAcrescimo
'    End If
    
    dQuantAlocar = Arredonda_Estoque(dQuantAlocar * dFator)
    '###################################################################
    
    QuantAlocar.Caption = Formata_Estoque(dQuantAlocar)
    
    lErro = Preenche_Tela2(objItemNF, colEstoqueProduto, sUMEstoque)
    If lErro <> SUCESSO Then Error 42763
    
    Preenche_Tela = SUCESSO
    
    Exit Function
    
Erro_Preenche_Tela:

    Select Case Err
    
        Case 39377, 42763
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162402)
    
    End Select
            
End Function

Private Sub SomaAlocado(dTotalAlocado)

Dim iIndice As Integer

    dTotalAlocado = 0
    
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
        dTotalAlocado = dTotalAlocado + StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
    Next

    Exit Sub
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col
            'Produto
            Case iGrid_QuantAloc_Col
                lErro = Saida_Celula_QuantAlocada(objGridInt)
                If lErro <> SUCESSO Then Error 39378
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39379

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 39378

        Case 39379
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162403)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_QuantAlocada(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantDisponivel As Double
Dim dQuantAlocada As Double
Dim dTotalAlocado As Double
Dim dQuantAlocar As Double

On Error GoTo Erro_Saida_Celula_QuantAlocada

    Set objGridInt.objControle = QuantAlocada

    If Len(Trim(QuantAlocada.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(QuantAlocada.Text)
        If lErro <> SUCESSO Then Error 39380
        
        dQuantAlocada = StrParaDbl(QuantAlocada.Text)
        
        dQuantDisponivel = StrParaDbl(GridAlocacao.TextMatrix(GridAlocacao.Row, iGrid_QuantDisp_Col))
        
        If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO And dQuantDisponivel < dQuantAlocada Then Error 39382
                
        GridAlocacao.TextMatrix(GridAlocacao.Row, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocada)
        
        Call SubTotal_Calcula(objGridAlocacao, iGrid_QuantAloc_Col, dTotalAlocado)
        
        dQuantAlocar = StrParaDbl(QuantAlocar.Caption)
        
        If dTotalAlocado > dQuantAlocar + QTDE_ESTOQUE_DELTA Then Error 39383
                
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 39381
    
    Call SubTotal_Calcula(objGridAlocacao, iGrid_QuantAloc_Col, dTotalAlocado)
    TotalAlocado.Caption = Formata_Estoque(dTotalAlocado)
    
    Saida_Celula_QuantAlocada = SUCESSO

    Exit Function
    
Erro_Saida_Celula_QuantAlocada:

    Saida_Celula_QuantAlocada = Err
    
    Select Case Err
    
        Case 39380, 39381
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 39382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_ALOCADA_MAIOR_DISPONIVEL", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 39383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_ALOCACAO_SUPERIOR_ALOCAR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162404)
            
    End Select
 
    Exit Function
 
End Function

Private Function Preenche_Tela2(objItemNF As ClassItemNF, colEstoqueProduto As colEstoqueProduto, sUMEstoque As String) As Long

Dim objEstoqueProduto As ClassEstoqueProduto
Dim iIndice As Integer
Dim dQuantAlocada As Double
Dim objAlocacao As ClassItemNFAlocacao
Dim dTotalAlocado As Double
Dim objReservaItem As ClassReservaItem
Dim iEscaninho As Integer

    UnidadeMedida.Caption = sUMEstoque
    
    objGridAlocacao.iLinhasExistentes = colEstoqueProduto.Count
    
    Call TipoDocInfo_Escaninho(giCodigo, iEscaninho)

    iIndice = 1
    For Each objEstoqueProduto In colEstoqueProduto
        
        GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        
        Select Case iEscaninho
        
            Case ESCANINHO_CONSIG_NOSSO
            
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantConsig)
                
            Case ESCANINHO_BENEF_3
                
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantBenef3)
                
            Case Else
                GridAlocacao.TextMatrix(iIndice, iGrid_QuantDisp_Col) = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
        
        End Select
        
        dQuantAlocada = 0
        
        If objItemNF.iPossuiGrade = DESMARCADO Then
        
            For Each objAlocacao In objItemNF.colAlocacoes
                If objAlocacao.iAlmoxarifado = objEstoqueProduto.iAlmoxarifado Then dQuantAlocada = dQuantAlocada + objAlocacao.dQuantidade
            Next
        
        Else
            For Each objReservaItem In objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1).colLocalizacao
                If objReservaItem.iAlmoxarifado = objEstoqueProduto.iAlmoxarifado Then dQuantAlocada = dQuantAlocada + objReservaItem.dQuantidade
            Next
        End If
        
        If dQuantAlocada > 0 Then GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocada)
        
        iIndice = iIndice + 1
    Next
    
    Call SomaAlocado(dTotalAlocado)
    TotalAlocado.Caption = Formata_Estoque(dTotalAlocado)
    
    Preenche_Tela2 = SUCESSO
    
    Exit Function
    
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    If Item.ListIndex = -1 Then Exit Sub

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 39389

    Item.ListIndex = -1
    Call GridAlocacao_Limpa
    Descricao.Caption = ""
    QuantAlocar.Caption = ""
    UnidadeMedida.Caption = ""

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 39389

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162405)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dTotalAlocado As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Call SubTotal_Calcula(objGridAlocacao, iGrid_QuantAloc_Col, dTotalAlocado)
    
    dQuantAlocar = StrParaDbl(QuantAlocar.Caption)
    
    If dTotalAlocado <> dQuantAlocar Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALOCADO_MENOR_ALOCAR", dTotalAlocado, dQuantAlocar)
        If vbMsgRes = vbNo Then Error 39390
        
    End If
    
    Set objItemNF = gcolItemNF.Item(Item.ListIndex + 1)
    
    lErro = Move_Tela_Memoria(objItemNF)
    If lErro <> SUCESSO Then Error 39391
    
    giRetornoTela = vbOK
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 39390, 39391
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162406)
            
    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objItemNF As ClassItemNF) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAlocada As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_Move_Tela_Memoria

    
    If objItemNF.iPossuiGrade = DESMARCADO Then
    
        Set objItemNF.colAlocacoes = New ColAlocacoesItemNF
    
    Else
        
        Set objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1).colLocalizacao = New Collection
    
    End If
    
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
    
        dQuantAlocada = StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
    
        If dQuantAlocada > 0 Then
            
            objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
                    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then Error 39392
            If lErro <> SUCESSO Then Error 39393
            
            If objItemNF.iPossuiGrade = DESMARCADO Then
            
                objItemNF.colAlocacoes.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, dQuantAlocada
            Else
            
                Set objReservaItem = New ClassReservaItem
                
                objReservaItem.dQuantidade = dQuantAlocada
                objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo
                objReservaItem.iFilialEmpresa = objAlmoxarifado.iFilialEmpresa
                objReservaItem.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
                
                objItemNF.colItensRomaneioGrade(Produto.ListIndex + 1).colLocalizacao.Add objReservaItem
                               
            End If
        End If
    Next
            
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
        
        Case 39392
        
        Case 39393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162407)
    End Select
    
    Exit Function

End Function


Private Sub Limpa_Tela_Alocacao()

Dim iIndice As Integer
    
    TotalAlocado.Caption = ""
        
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
    
        GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col) = ""
        
    Next
    
    Exit Sub
    
End Sub

Private Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubTotal As Double)
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer

    dSubTotal = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
        'Acumula em dSubTotal
        dSubTotal = dSubTotal + StrParaDbl(IIf(Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna))) > 0, objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna), 0))

    Next

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOCALIZACAO_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Localização de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LocalizacaoProduto"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub Procura_Almoxarifado(objEstoqueProduto As ClassEstoqueProduto, colReserva As colReservaItem)
'Percorre uma colecao de reservas e verifica se o almoxarifado passado está na coleção.

Dim iIndice As Integer

    For iIndice = 1 To colReserva.Count
    
        If objEstoqueProduto.iAlmoxarifado = colReserva(iIndice).iAlmoxarifado Then
            objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReserva(iIndice).dQuantidade
            Exit For
        End If

    Next
    
    Exit Sub
    
End Sub

Private Sub QuantTotalReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantTotalReserva, Source, X, Y)
End Sub

Private Sub QuantTotalReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantTotalReserva, Button, Shift, X, Y)
End Sub

Private Sub TotalAlocado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalAlocado, Source, X, Y)
End Sub

Private Sub TotalAlocado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalAlocado, Button, Shift, X, Y)
End Sub

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub QuantAlocar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantAlocar, Source, X, Y)
End Sub

Private Sub QuantAlocar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantAlocar, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Function TipoDocInfo_Escaninho(ByVal iTipo As Integer, iEscaninho As Integer)

    Select Case iTipo
    
        Case DOCINFO_NFISPC, DOCINFO_NFFISPC
            iEscaninho = ESCANINHO_CONSIG_NOSSO
        
        Case DOCINFO_NFISRMB3PV, DOCINFO_NFISBF, DOCINFO_NFISFBF, DOCINFO_NFISFBFPV
            iEscaninho = ESCANINHO_BENEF_3
        
        Case Else
            iEscaninho = ESCANINHO_DISPONIVEL
        
    End Select

End Function

