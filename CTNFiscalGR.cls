VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const PRODUTO_ADVALOREN = "999999"

'Constante relacionada ao GridComprovServ
Private Const NUM_MAXIMO_CONHECFRETE = 240

Private WithEvents objEventoComprovServ As AdmEvento
Attribute objEventoComprovServ.VB_VarHelpID = -1
Private gobjCT As Object

Public Property Get GridComprovServ(ByVal objCT As Object) As Object
     Set GridComprovServ = objCT.objUserControl.Controls("GridComprovServ")
End Property

Public Property Get ComprovServCon(ByVal objCT As Object) As Object
     Set ComprovServCon = objCT.objUserControl.Controls("ComprovServCon")
End Property

Public Property Get DataCon(ByVal objCT As Object) As Object
     Set DataCon = objCT.objUserControl.Controls("DataCon")
End Property

Public Property Get ServicoCon(ByVal objCT As Object) As Object
     Set ServicoCon = objCT.objUserControl.Controls("ServicoCon")
End Property

Public Property Get DescricaoCon(ByVal objCT As Object) As Object
     Set DescricaoCon = objCT.objUserControl.Controls("DescricaoCon")
End Property

Public Property Get QuantCon(ByVal objCT As Object) As Object
     Set QuantCon = objCT.objUserControl.Controls("QuantCon")
End Property

Public Property Get PrecoCon(ByVal objCT As Object) As Object
     Set PrecoCon = objCT.objUserControl.Controls("PrecoCon")
End Property

Public Property Get AdValorenCon(ByVal objCT As Object) As Object
     Set AdValorenCon = objCT.objUserControl.Controls("AdValorenCon")
End Property

Public Property Get ValorMercadoriaCon(ByVal objCT As Object) As Object
     Set ValorMercadoriaCon = objCT.objUserControl.Controls("ValorMercadoriaCon")
End Property

Public Property Get ValorContainerCon(ByVal objCT As Object) As Object
     Set ValorContainerCon = objCT.objUserControl.Controls("ValorContainerCon")
End Property

Public Property Get LabelComprovante(ByVal objCT As Object) As Object
     Set LabelComprovante = objCT.objUserControl.Controls("LabelComprovante")
End Property

Public Property Get BotaoComprovante(ByVal objCT As Object) As Object
     Set BotaoComprovante = objCT.objUserControl.Controls("BotaoComprovante")
End Property

Public Property Get ValorMercadoria(ByVal objCT As Object) As Object
     Set ValorMercadoria = objCT.objUserControl.Controls("ValorMercadoria")
End Property

Public Property Get ValorContainer(ByVal objCT As Object) As Object
     Set ValorContainer = objCT.objUserControl.Controls("ValorContainer")
End Property

Public Property Get Origem(ByVal objCT As Object) As Object
     Set Origem = objCT.objUserControl.Controls("Origem")
End Property

Public Property Get Destino(ByVal objCT As Object) As Object
     Set Destino = objCT.objUserControl.Controls("Destino")
End Property

'--- fim dos properties get dos controles da tela

Public Function NF_CargaPosFormLoad(ByVal objCT As Object) As Long

Dim iSubTipo As Integer
Dim lErro As Long

On Error GoTo Erro_NF_CargaPosFormLoad

    Set objCT.gobjInfoUsu.objGridComprovServ = New AdmGrid

    'Inicializa Grid de Comprovante
    lErro = Inicializa_Grid_ComprovServ(objCT)
    If lErro <> SUCESSO Then gError 184457

    'Inicializa a máscara do Produto do grid de Comprovante
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ServicoCon(objCT))
    If lErro <> SUCESSO Then gError 184458

    NF_CargaPosFormLoad = SUCESSO

    Exit Function

Erro_NF_CargaPosFormLoad:

    NF_CargaPosFormLoad = gErr

    Select Case gErr

        Case 184457, 184458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184459)

    End Select

    Exit Function

End Function

Public Function NF_Form_UnLoad(ByVal objCT As Object) As Long

    'Libera as variáveis globais da tela
    Set objEventoComprovServ = Nothing
    Set objCT.gobjInfoUsu.objGridComprovServ = Nothing

    NF_Form_UnLoad = SUCESSO

End Function

'Cyntia
Private Function Inicializa_Grid_ComprovServ(ByVal objCT As Object) As Long
'Inicializa o Grid de ComprovServ

Dim objGridInt As Object

    Set objGridInt = New AdmGrid
    Set objCT.gobjInfoUsu.objGridComprovServ = objGridInt

    'tela em questão
    Set objGridInt.objForm = objCT

    With objGridInt

        'Títulos das colunas
        .colColuna.Add (" ")
        .colColuna.Add ("Comprovante")
        .colColuna.Add ("Data")
        .colColuna.Add ("Serviço")
        .colColuna.Add ("Descrição do Serviço")
        .colColuna.Add ("Quantidade")
        .colColuna.Add ("Preço Unitário")
        .colColuna.Add ("Ad Valoren")
        .colColuna.Add ("Valor Mercadoria")
        .colColuna.Add ("Valor Container")

        'Controles que participam do Grid
        .colCampo.Add (ComprovServCon(objCT).Name)
        .colCampo.Add (DataCon(objCT).Name)
        .colCampo.Add (ServicoCon(objCT).Name)
        .colCampo.Add (DescricaoCon(objCT).Name)
        .colCampo.Add (QuantCon(objCT).Name)
        .colCampo.Add (PrecoCon(objCT).Name)
        .colCampo.Add (AdValorenCon(objCT).Name)
        .colCampo.Add (ValorMercadoriaCon(objCT).Name)
        .colCampo.Add (ValorContainerCon(objCT).Name)

    End With

    With objCT.gobjInfoUsu

        'Colunas do Grid
        .iGrid_ComprovServCon_col = 1
        .iGrid_DataCon_Col = 2
        .iGrid_ServicoCon_Col = 3
        .iGrid_DescricaoCon_Col = 4
        .iGrid_QuantCon_Col = 5
        .iGrid_PrecoCon_Col = 6
        .iGrid_AdValorenCon_Col = 7
        .iGrid_ValorMercadoriaCon_col = 8
        .iGrid_ValorContainerCon_col = 9

    End With

    'Grid do GridInterno
    objGridInt.objGrid = GridComprovServ(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_CONHECFRETE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridComprovServ(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Usado para que se possa utilizar a Rotina_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ComprovServ = SUCESSO

End Function

Public Sub ComprovServCon_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComprovServCon_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComprovServ)

    Call MaskEdBox_TrataGotFocus(ComprovServCon(objCT), objCT.iAlterado)

End Sub

Public Sub ComprovServCon_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComprovServ)

End Sub

Public Sub ComprovServCon_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridComprovServ.objControle = ComprovServCon(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComprovServ)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function NF_Traz_NF_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NF_Traz_NF_Tela

    If objNFiscal.dValorMercadoria > 0 Then ValorMercadoria(objCT).Caption = Format(objNFiscal.dValorMercadoria, "Standard")
    If objNFiscal.dValorContainer > 0 Then ValorContainer(objCT).Caption = Format(objNFiscal.dValorContainer, "Standard")

     'cyntia
    'Lê a parte de Comprovante da Nota Fiscal
    lErro = CF("NFiscalComprovante_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 99094 Then gError 99086

    Origem(objCT).Text = objNFiscal.sOrigem
    Destino(objCT).Text = objNFiscal.sDestino

    'cyntia
    'Preenche o Grid com os Comprovantes da Nota Fiscal
    lErro = Preenche_GridComprovante(objCT, objNFiscal)
    If lErro <> SUCESSO Then gError 99087

    NF_Traz_NF_Tela = SUCESSO

    Exit Function

Erro_NF_Traz_NF_Tela:

    NF_Traz_NF_Tela = gErr

    Select Case gErr

        Case 25828, 25829, 39031, 35705, 39023, 39024, 39342, 39343, 39661, 51609, 83325, 83326, 500116, 99086, 99087

        Case 39025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'Cyntia
Private Function Rotina_Comprovantes_Tela(ByVal objCT As Object, objComprovante As ClassCompServ) As Long
'Faz o Tratamento do Comprovante

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim dTotal As Double
Dim dValorAdValoren As Double
Dim sProdutoEnxuto As String
Dim bExiste_AdValoren As Boolean
Dim iLinhas As Integer
Dim iIndice2 As Integer
Dim bAchou_Comp As Boolean
Dim dTotalComp As Double
Dim dQuantAdValoren As Double
Dim dValorTotAdValoren As Double
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Rotina_Comprovantes_Tela

    lErro = CF2(objCT, "NF_Limpa_GridItens")
    If lErro <> SUCESSO Then gError 99078
    
    LabelComprovante(objCT).Caption = ""
    
    dValorAdValoren = 0
    bExiste_AdValoren = False

    For iIndice = 1 To objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes

        objComprovante.lCodigo = StrParaLong(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col))
        objComprovante.iFilialEmpresa = giFilialEmpresa

        lErro = CF("CompServGR_Le", objComprovante)
        If lErro <> SUCESSO And lErro <> 97419 Then gError 99045

        If lErro = 97419 Then gError 99051

        bAchou_Comp = False

        For iIndice2 = 1 To objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes
            If iIndice2 <> GridComprovServ(objCT).Row Then
                If objComprovante.lCodigo = GridComprovServ(objCT).TextMatrix(iIndice2, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) Then
                    bAchou_Comp = True
                    Exit For
                End If
            End If
        Next

        objProduto.sCodigo = objComprovante.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 99046

        If lErro = 28030 Then gError 99047

        If bAchou_Comp = False Then

            'coloca a mascara no produto
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 99049

            'Coloca o produto já com a máscara no controle
            ServicoCon(objCT).PromptInclude = False
            ServicoCon(objCT).Text = sProdutoEnxuto
            ServicoCon(objCT).PromptInclude = True

            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) = objComprovante.lCodigo
            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_DataCon_Col) = Format(objComprovante.dtDataEmissao, "dd/mm/yyyy")
            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ServicoCon_Col) = ServicoCon(objCT).Text
            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_DescricaoCon_Col) = objProduto.sDescricao
            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_QuantCon_Col) = Format(objComprovante.dQuantidade, "STANDARD")
            GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_PrecoCon_Col) = Format(objComprovante.dFretePeso, "STANDARD")

            If objComprovante.dAdValoren <> 0 Then
                GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_AdValorenCon_Col) = Format(objComprovante.dAdValoren, "Percent")
                GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ValorMercadoriaCon_col) = Format(objComprovante.dValorMercadoria, "STANDARD")
                GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ValorContainerCon_col) = Format(objComprovante.dValorContainer, "STANDARD")
            End If
        End If

        If objComprovante.dAdValoren <> 0 Then

                dQuantAdValoren = dQuantAdValoren + objComprovante.dQuantidade
                dValorAdValoren = (objComprovante.dValorMercadoria + objComprovante.dValorContainer) * objComprovante.dAdValoren
                dValorTotAdValoren = dValorTotAdValoren + dValorAdValoren
                bExiste_AdValoren = True

        End If

        dTotalComp = (objComprovante.dFretePeso * objComprovante.dQuantidade)
        dTotal = dTotal + dTotalComp

        dValorAdValoren = 0

        lErro = Rotina_ItensServico_Tela(objCT, objComprovante, objProduto, dTotalComp)
        If lErro <> SUCESSO Then gError 99072

        dTotalComp = 0

    Next

    If bExiste_AdValoren = True Then

        'malandragem para mascarar o produto...
        objCT.Produto.PromptInclude = False
        objCT.Produto.Text = PRODUTO_ADVALOREN
        objCT.Produto.PromptInclude = True

        'Usado para recolher a descrição do produto
        lErro = CF("Produto_Critica_Filial", objCT.Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 Then gError 99073

        If lErro = 51381 Then gError 99074

        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 99075

        'Coloca o produto já com a máscara no controle
        objCT.Produto.PromptInclude = False
        objCT.Produto.Text = sProdutoEnxuto
        objCT.Produto.PromptInclude = True

        objCT.objGridItens.iLinhasExistentes = objCT.objGridItens.iLinhasExistentes + 1

        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_Produto_Col) = objCT.Produto.Text
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_DescProduto_Col) = objProduto.sDescricao
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_Quantidade_Col) = Format(dQuantAdValoren, "STANDARD")
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_PrecoTotal_Col) = Format(dValorTotAdValoren, "STANDARD")
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_ValorUnitario_Col) = IIf(Format((dValorTotAdValoren / dQuantAdValoren), "STANDARD") = 0, "", Format((dValorTotAdValoren / dQuantAdValoren), "STANDARD"))

        Call objCT.gobjTelaProjetoInfo.Adiciona_Linha(objCT.objGridItens.iLinhasExistentes)
        
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = objCT.gobjTribTab.Inclusao_Item_Grid(objCT.objGridItens.iLinhasExistentes, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 99079

        Set objItemNF = New ClassItemNF
        objCT.gobjNFiscal.ColItensNF.Add1 objItemNF
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).sUMEstoque = objProduto.sSiglaUMEstoque
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).iItem = objCT.objGridItens.iLinhasExistentes
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).sProduto = objProduto.sCodigo
        
        Call objCT.PrecoTotal_Calcula(objCT.objGridItens.iLinhasExistentes, False)

    End If

    LabelComprovante(objCT).Caption = Format(dTotal, "STANDARD")

    Call objCT.ValorTotal_Calcula

    Rotina_Comprovantes_Tela = SUCESSO

    Exit Function

Erro_Rotina_Comprovantes_Tela:

    Rotina_Comprovantes_Tela = gErr

    Select Case gErr

        Case 99045, 99046, 99072, 99073, 99078, 99079

        Case 99047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 99074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, PRODUTO_ADVALOREN)

        Case 99049, 99075
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 99051
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objComprovante.lCodigo, objComprovante.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Rotina_ItensServico_Tela(ByVal objCT As Object, objComprovante As ClassCompServ, objProduto As ClassProduto, ByVal dTotalComp As Double) As Long
'Faz o Tratamento do item do Produto

Dim lErro As Long
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim bAchou_Servico As Boolean
Dim sProdutoFormatado As String
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Rotina_ItensServico_Tela

    bAchou_Servico = False

    'Verifica se já existe este serviço em outra linha do Grid
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes

        lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99077

        If objProduto.sCodigo = sProdutoFormatado Then

            objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col) = Format(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col) + objComprovante.dQuantidade, "STANDARD")
            objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col) = Format(dTotalComp + objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col), "STANDARD")
            objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_ValorUnitario_Col) = Format(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col) / objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col), "STANDARD")
            bAchou_Servico = True

''*** incluidos p/tratamento de tributacao *******************************
        lErro = objCT.gobjTribTab.Alteracao_Item_Grid(iIndice)
        If lErro <> SUCESSO Then gError 99080
'*** fim tributacao

            Exit For

        End If

    Next

    If bAchou_Servico = False Then

        objCT.objGridItens.iLinhasExistentes = objCT.objGridItens.iLinhasExistentes + 1

        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 99056

        'Coloca o produto já com a máscara no controle
        objCT.Produto.PromptInclude = False
        objCT.Produto.Text = sProdutoEnxuto
        objCT.Produto.PromptInclude = True

        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_Produto_Col) = objCT.Produto.Text
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_DescProduto_Col) = objProduto.sDescricao
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_Quantidade_Col) = Format(objComprovante.dQuantidade, "STANDARD")
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_ValorUnitario_Col) = Format((dTotalComp / objComprovante.dQuantidade), "STANDARD")
        objCT.GridItens.TextMatrix(objCT.objGridItens.iLinhasExistentes, objCT.iGrid_PrecoTotal_Col) = Format(dTotalComp, "STANDARD")

        Call objCT.gobjTelaProjetoInfo.Adiciona_Linha(objCT.objGridItens.iLinhasExistentes)
        
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = objCT.gobjTribTab.Inclusao_Item_Grid(objCT.objGridItens.iLinhasExistentes, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 99081

        Set objItemNF = New ClassItemNF
        objCT.gobjNFiscal.ColItensNF.Add1 objItemNF
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).sUMEstoque = objProduto.sSiglaUMEstoque
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).iItem = objCT.objGridItens.iLinhasExistentes
        objCT.gobjNFiscal.ColItensNF(objCT.objGridItens.iLinhasExistentes).sProduto = objProduto.sCodigo
        
        iIndice = objCT.objGridItens.iLinhasExistentes

    End If

    Call objCT.PrecoTotal_Calcula(iIndice, False)

    Rotina_ItensServico_Tela = SUCESSO

    Exit Function

Erro_Rotina_ItensServico_Tela:

    Rotina_ItensServico_Tela = gErr

    Select Case gErr

        Case 99056
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 99077, 99080, 99081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

'Cyntia
Public Sub BotaoComprovante_Click(ByVal objCT As Object)
'Chama o browser de Comprovante
'Este Browser traz os dados do comprovante que não estejam associados
'a uma nota fiscal, que estejam abertos ou faturáveis e que a origem e o
'destino sejam os mesmos

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprovante As New ClassCompServ
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoComprovante_Click

    If Len(Trim(objCT.Cliente.Text)) = 0 Then gError 99113

    objCliente.sNomeReduzido = objCT.Cliente.Text

    lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
    If lErro <> SUCESSO Then gError 99122

    'Verifica se o browser está sendo chamado do controle(F3)ou pelo grid
    If Not (objCT.ActiveControl Is ComprovServCon(objCT)) Then

        'Verifica se tem alguma linha selecionada no Grid
        If GridComprovServ(objCT).Row = 0 Then gError 99042

        objComprovante.lCodigo = StrParaLong(ComprovServCon(objCT).Text)

    End If

    colSelecao.Add objCliente.lCodigo

    Set objEventoComprovServ = New AdmEvento
    Set gobjCT = objCT
    
    'Chama a tela de browse de Comprovante
    Call Chama_Tela("ComprovanteLista", colSelecao, objComprovante, objEventoComprovServ)

    Exit Sub

Erro_BotaoComprovante_Click:

    Select Case gErr

        Case 99042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 99113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 99122

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

'Cyntia
Private Sub objEventoComprovServ_evSelecao(obj1 As Object)

Dim objComprovante As ClassCompServ
Dim lErro As Long
Dim iIndice As Integer, objCT As Object

On Error GoTo Erro_objEventoComprovServ_evSelecao

    Set objComprovante = obj1
    Set objCT = gobjCT

    'Verifica se alguma linha está selecionada
    If GridComprovServ(objCT).Row < 1 Then Exit Sub

    For iIndice = 1 To objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes
        If iIndice <> GridComprovServ(objCT).Row Then
            If objComprovante.lCodigo = GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) Then gError 99043
        End If
    Next

    ComprovServCon(objCT).PromptInclude = False
    ComprovServCon(objCT).Text = objComprovante.lCodigo
    ComprovServCon(objCT).PromptInclude = True

    If Not (objCT.ActiveControl Is ComprovServCon(objCT)) Then

        'Atualiza o total de linhas existentes no grid
        If GridComprovServ(objCT).Row - GridComprovServ(objCT).FixedRows = objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes Then
            objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes = objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes + 1
        End If

        GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) = ComprovServCon(objCT).Text

        'Faz o Tratamento do Comprovante
        lErro = Rotina_Comprovantes_Tela(objCT, objComprovante)
        If lErro <> SUCESSO Then gError 99044

    End If

    objCT.Show

    Exit Sub

Erro_objEventoComprovServ_evSelecao:

    Select Case gErr

        Case 99043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTE_JA_EXISTENTE", gErr, objComprovante.lCodigo, iIndice)

        Case 99044, 99050

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub NF_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String

On Error GoTo Erro_NF_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        'Cyntia
        'Código do Comprovante
        Case ComprovServCon(objCT).Name

            If Len(Trim(GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ComprovServCon_col))) = 0 Then
                ComprovServCon(objCT).Enabled = True
            Else
                ComprovServCon(objCT).Enabled = False
            End If

    End Select

    Exit Sub

Erro_NF_Rotina_Grid_Enable:

    Select Case gErr

        Case 35768, 35769, 35770, 35771, 35772, 35773, 83327

        Case 35775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub GridComprovServ_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridComprovServ, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComprovServ, objCT.iAlterado)
    End If

End Sub

Public Sub GridComprovServ_EnterCell(ByVal objCT As Object)

    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComprovServ, objCT.iAlterado)

End Sub

Public Sub GridComprovServ_GotFocus(ByVal objCT As Object)

    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridComprovServ)

End Sub

Public Sub GridComprovServ_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridComprovServ, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComprovServ, objCT.iAlterado)
    End If

End Sub

Public Sub GridComprovServ_LeaveCell(ByVal objCT As Object)

    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridComprovServ)

End Sub

Public Sub GridComprovServ_Validate(ByVal objCT As Object, Cancel As Boolean)

    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridComprovServ)

End Sub

Public Sub GridComprovServ_RowColChange(ByVal objCT As Object)

    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridComprovServ)

End Sub

Public Sub GridComprovServ_Scroll(ByVal objCT As Object)

    Call Grid_Scroll(objCT.gobjInfoUsu.objGridComprovServ)

End Sub

'cyntia
Public Sub GridComprovServ_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim objComprovante As New ClassCompServ

On Error GoTo Erro_GridComprovServ_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridComprovServ)

    If KeyCode = vbKeyDelete Then

        lErro = Rotina_Comprovantes_Tela(objCT, objComprovante)
        If lErro <> SUCESSO Then gError 99066

    End If

    Exit Sub

Erro_GridComprovServ_KeyDown:

    Select Case gErr

        Case 99066

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Public Function NF_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_NF_Saida_Celula

    'Verifica qual o Grid em questão
    Select Case objGridInt.objGrid.Name

        'Se for o GridComprovServ
        Case GridComprovServ(objCT).Name

            lErro = Saida_Celula_GridComprovServ(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 99057

    End Select

    NF_Saida_Celula = SUCESSO

    Exit Function

Erro_NF_Saida_Celula:

    NF_Saida_Celula = gErr

    Select Case gErr

        Case 99057
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComprovServ(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComprovServ

    'Verifica qual a coluna atual do Grid
    If objGridInt.objGrid.Col = objCT.gobjInfoUsu.iGrid_ComprovServCon_col Then

        lErro = Saida_Celula_ComprovServCon(objCT, objGridInt)
        If lErro <> SUCESSO Then gError 99058

    End If

    Saida_Celula_GridComprovServ = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComprovServ:

    Saida_Celula_GridComprovServ = gErr

    Select Case gErr

        Case 99058
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
'cyntia
Private Function Saida_Celula_ComprovServCon(ByVal objCT As Object, objGridComprovServ As AdmGrid) As Long
'Faz a crítica da célula do Código do Comprovante que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objComprovante As New ClassCompServ
Dim objCliente As New ClassCliente

On Error GoTo Erro_Saida_Celula_ComprovServCon

    Set objGridComprovServ.objControle = ComprovServCon(objCT)

    'Verifica se o comprovante foi preenchido
    If Len(Trim(ComprovServCon(objCT).ClipText)) > 0 Then

        If Len(Trim(objCT.Cliente.Text)) = 0 Then gError 99121

        'Critica a existencia do comprovante
        objComprovante.lCodigo = StrParaLong(ComprovServCon(objCT).Text)
        objComprovante.iFilialEmpresa = giFilialEmpresa

        lErro = CF("CompServNFiscal_Le", objComprovante)
        If lErro <> SUCESSO And lErro <> 99070 Then gError 99082

        If lErro = 99166 Then gError 99070

        'comprovante já existe em outra nota fiscal --> erro
        If objComprovante.lNumIntNota <> 0 Then gError 99179

        objCliente.sNomeReduzido = objCT.Cliente.Text

        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
        If lErro <> SUCESSO Then gError 99123

        If objComprovante.lCliente <> objCliente.lCodigo Then gError 99120

        For iIndice = 1 To objGridComprovServ.iLinhasExistentes
            If iIndice <> GridComprovServ(objCT).Row Then
                If Trim(ComprovServCon(objCT).ClipText) = Trim(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col)) Then gError 99076
            End If
        Next

        'Atualiza o total de linhas existentes no grid
        If GridComprovServ(objCT).Row - GridComprovServ(objCT).FixedRows = objGridComprovServ.iLinhasExistentes Then
            objGridComprovServ.iLinhasExistentes = objGridComprovServ.iLinhasExistentes + 1
        End If

        GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) = objComprovante.lCodigo

        'Faz o Tratamento do Comprovante
        lErro = Rotina_Comprovantes_Tela(objCT, objComprovante)
        If lErro <> SUCESSO Then gError 99059

        GridComprovServ(objCT).TextMatrix(GridComprovServ(objCT).Row, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridComprovServ)
    If lErro <> SUCESSO Then gError 99061

    Saida_Celula_ComprovServCon = SUCESSO

    Exit Function

Erro_Saida_Celula_ComprovServCon:

    Saida_Celula_ComprovServCon = gErr

    Select Case gErr

        Case 99059, 99061, 99082, 99123
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case 99076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTE_JA_EXISTENTE", gErr, ComprovServCon(objCT).ClipText, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case 99083
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objComprovante.lCodigo, objComprovante.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case 99120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLI_NAO_RELACIONADO_COMP", gErr, objCliente.lCodigo, objComprovante.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case 99121
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case 99179
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FATURADO", gErr, objComprovante.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridComprovServ)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NF_Move_NF_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
'Move os dados dos Grids para memória

Dim lErro As Long

On Error GoTo Erro_NF_Move_NF_Memoria

    objNFiscal.dValorMercadoria = StrParaDbl(ValorMercadoria(objCT).Text)
    objNFiscal.dValorContainer = StrParaDbl(ValorContainer(objCT).Text)
    objNFiscal.sOrigem = Origem(objCT).Text
    objNFiscal.sDestino = Destino(objCT).Text
    
    'cyntia
    'move o grid de comprovantes de serviços para a memória
    lErro = Move_GridComprovServ_Memoria(objCT, objNFiscal)
    If lErro <> SUCESSO Then gError 99140

    NF_Move_NF_Memoria = SUCESSO

    Exit Function

Erro_NF_Move_NF_Memoria:

    NF_Move_NF_Memoria = gErr

    Select Case gErr

        Case 39041, 35933, 35937, 35938, 99140

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'cyntia
Private Function Move_GridComprovServ_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
'Move os dados de gridComprovServ para a memória

Dim iIndice As Integer
Dim objComprovServ As ClassCompServ
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridComprovServ_Memoria

    'Se não há comprovantes a recolher, sai da função
    If objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes = 0 Then Exit Function

    'Para cada comprovante do grid
    For iIndice = 1 To objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes

        'Verifica se o comprovante esta relacionado com o cliente
        lErro = CF("Verifica_Cliente1", StrParaLong(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col)), objNFiscal.lCliente)
        If lErro <> SUCESSO And lErro <> 99171 Then gError 99114

        'Comprovante não relacionado ao cliente
        If lErro = 99171 Then gError 99119

        Set objComprovServ = New ClassCompServ

        lErro = CF("Produto_Formata", GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ServicoCon_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 99100

        'recolhe os dados do grid de comprovantes de serviços
        objComprovServ.lCodigo = StrParaLong(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col))
        objComprovServ.dtDataEmissao = StrParaDate(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCon_Col))
        objComprovServ.sProduto = sProdutoFormatado
        objComprovServ.dQuantidade = StrParaDbl(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QuantCon_Col))
        objComprovServ.dFretePeso = StrParaDbl(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoCon_Col))
        objComprovServ.dAdValoren = PercentParaDbl(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_AdValorenCon_Col))
        objComprovServ.dValorMercadoria = StrParaDbl(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorMercadoriaCon_col))
        objComprovServ.dValorContainer = StrParaDbl(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorContainerCon_col))

        'Carrega o obj na coleção
        objNFiscal.colComprovServ.Add objComprovServ

    Next

    Move_GridComprovServ_Memoria = SUCESSO

    Exit Function

Erro_Move_GridComprovServ_Memoria:

    Move_GridComprovServ_Memoria = gErr

    Select Case gErr

        Case 99100, 99114

        Case 99119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLI_NAO_RELACIONADO_COMP", gErr, objNFiscal.lCliente, StrParaLong(GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col)))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NF_Limpa_Tela_Cust(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NF_Limpa_Tela_Cust

    'Indica que está limpando a tela
    LabelComprovante(objCT).Caption = ""

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridComprovServ)

    NF_Limpa_Tela_Cust = SUCESSO
    
    Exit Function

Erro_NF_Limpa_Tela_Cust:

    NF_Limpa_Tela_Cust = gErr

    Select Case gErr

        Case 500115

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Preenche_GridComprovante(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com as Parcelas da Nota Fiscal

Dim objComprovServ As ClassCompServ
Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objCodigoNome As New ClassCliente
Dim objCotacaoMoeda As New ClassCotacaoMoeda
Dim dTotal As Double

On Error GoTo Erro_Preenche_GridComprovante

    'Preenche o Grid
    For Each objComprovServ In objNFiscal.colComprovServ

        iIndice = iIndice + 1

        objProduto.sCodigo = objComprovServ.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 99088

        If lErro = 28030 Then gError 99089

        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 99090

        'Coloca o produto já com a máscara no controle
        ServicoCon(objCT).PromptInclude = False
        ServicoCon(objCT).Text = sProdutoEnxuto
        ServicoCon(objCT).PromptInclude = True

        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ComprovServCon_col) = objComprovServ.lCodigo
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCon_Col) = Format(objComprovServ.dtDataEmissao, "dd/mm/yyyy")
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ServicoCon_Col) = ServicoCon(objCT).Text
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DescricaoCon_Col) = objProduto.sDescricao
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QuantCon_Col) = Format(objComprovServ.dQuantidade, "STANDARD")
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoCon_Col) = Format(objComprovServ.dFretePeso, "STANDARD")
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_AdValorenCon_Col) = Format(objComprovServ.dAdValoren, "Percent")
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorMercadoriaCon_col) = Format(objComprovServ.dValorMercadoria, "STANDARD")
        GridComprovServ(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorContainerCon_col) = Format(objComprovServ.dValorContainer, "STANDARD")

        dTotal = dTotal + (objComprovServ.dFretePeso * objComprovServ.dQuantidade)

    Next

    LabelComprovante(objCT).Caption = Format(dTotal, "STANDARD")

    objCT.gobjInfoUsu.objGridComprovServ.iLinhasExistentes = iIndice

    Preenche_GridComprovante = SUCESSO

    Exit Function

Erro_Preenche_GridComprovante:

    Preenche_GridComprovante = gErr

    Select Case gErr

        Case 99088

        Case 99089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 99090
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

   End Select

    Exit Function

End Function

Public Sub NF_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If objCT.ActiveControl Is ComprovServCon(objCT) Then
            Call BotaoComprovante_Click(objCT)
        End If
        
    End If

End Sub

Public Sub ValorMercadoria_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorContainer_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Origem_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Destino_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

