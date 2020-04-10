VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastreamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Grid relacionado com Rastreamento
Private iGrid_ItemNFRastro_Col As Integer
Private iGrid_ProdutoRastro_Col As Integer
Private iGrid_AlmoxRastro_Col As Integer
Private iGrid_EscaninhoRastro_Col As Integer
Private iGrid_UMRastro_Col As Integer
Private iGrid_QuantRastro_Col As Integer
Private iGrid_LoteRastro_Col As Integer
Private iGrid_LoteDataRastro_Col As Integer
Private iGrid_QuantLoteRastro_Col As Integer
Private iGrid_FilialOPRastro_Col As Integer

Dim bEscaninhoVisivel As Boolean 'se o campo escaninho é visivel para a tela em questão
Public bSaidaMaterial As Boolean 'se é uma tela que trata da saida de material do estoque
Public bTelaManutencao As Boolean 'se é uma tela de manutenção do rastreamento, ou seja, não edita/ ve os dados da nota
Public bDevolucaoMaterial As Boolean 'se é uma tela de devolução de material de saida ==> true
Public iCodigo As Integer 'contem o codigo do documento (TipoDocInfo)
Public bEditavel As Boolean 'se verdadeiro indica que é editavel o rastreamento desde a tela inicial, mesmo sem estar cadastrado

Private WithEvents objEventoLoteRastro As AdmEvento
Attribute objEventoLoteRastro.VB_VarHelpID = -1
Public objTela As Object
Public objGridInt As AdmGrid
Private objUserControl As Object

Public gcolcolRastreamentoSerie As Collection 'Vai guardar o detalhamentos das séries usadas 'Inserido por Wagner 10/03/2006
Dim iTipoMovtoAnt As Integer

Const ESCANINHO_RASTRO_CONSIGNADO = "Consignado"
Const ESCANINHO_RASTRO_DISPONIVEL = "Disponível"

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get NFiscal() As Object
     Set NFiscal = objUserControl.Controls("NFiscal")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get LblDataEmissao() As Object
     Set LblDataEmissao = objUserControl.Controls("LblDataEmissao")
End Property

'Rastreamento
Public Property Get BotaoLotes() As Object
    Set BotaoLotes = objUserControl.Controls("BotaoLotes")
End Property

'#########################################################
'Inserido por Wagner 10/03/2006
Public Property Get BotaoSerie() As Object
    Set BotaoSerie = objUserControl.Controls("BotaoSerie")
End Property
'#########################################################

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

'Rastreamento
Public Property Get ItemNFRastro() As Object
     Set ItemNFRastro = objUserControl.Controls("ItemNFRastro")
End Property

'Rastreamento
Public Property Get ProdutoRastro() As Object
     Set ProdutoRastro = objUserControl.Controls("ProdutoRastro")
End Property

'Rastreamento
Public Property Get AlmoxRastro() As Object
     Set AlmoxRastro = objUserControl.Controls("AlmoxRastro")
End Property

'Rastreamento
Public Property Get EscaninhoRastro() As Object
     Set EscaninhoRastro = objUserControl.Controls("EscaninhoRastro")
End Property

'Rastreamento
Public Property Get UMRastro() As Object
     Set UMRastro = objUserControl.Controls("UMRastro")
End Property

'Rastreamento
Public Property Get QuantRastro() As Object
     Set QuantRastro = objUserControl.Controls("QuantRastro")
End Property

'Rastreamento
Public Property Get LoteRastro() As Object
     Set LoteRastro = objUserControl.Controls("LoteRastro")
End Property

'Rastreamento
Public Property Get FilialOPRastro() As Object
     Set FilialOPRastro = objUserControl.Controls("FilialOPRastro")
End Property

'Rastreamento
Public Property Get LoteDataRastro() As Object
     Set LoteDataRastro = objUserControl.Controls("LoteDataRastro")
End Property

'Rastreamento
Public Property Get QuantLoteRastro() As Object
     Set QuantLoteRastro = objUserControl.Controls("QuantLoteRastro")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

Public Function Carrega_RastroItensNF(objNFiscal As ClassNFiscal) As Long
'Carrega rastreamento dos itens da nota fiscal

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLote As New ClassRastreamentoLote
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim colRastreamentoMovto As Collection
Dim colMovEstoque As Collection
Dim sProdutoEnxuto As String
Dim objProduto As ClassProduto
Dim iIndice2 As Integer 'Inserido por Wagner

On Error GoTo Erro_Carrega_RastroItensNF

    Call Grid_Limpa(objGridInt)

    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 10/03/2006

    'Para cada item da nota fiscal
    For Each objItemNF In objNFiscal.colItensNF

        '##############################################################
        'Se ProdutoRastro for uma Combo => Limpa
        If TypeName(ProdutoRastro) = "ComboBox" Then
            ProdutoRastro.Clear
        End If
        
        'Para cada Item da Grade ou componente do Kit de Venda
        For iIndice2 = 1 To objItemNF.colItensRomaneioGrade.Count
        
            Set objProduto = New ClassProduto
        
            'Lê o Produto
            objProduto.sCodigo = objItemNF.colItensRomaneioGrade.Item(iIndice2).sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 180046
        
            'Se ele for rastreado adiciona na combo
            If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
            
                lErro = Mascara_RetornaProdutoTela(objItemNF.colItensRomaneioGrade.Item(iIndice2).sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 180047
                                        
                ProdutoRastro.AddItem sProdutoMascarado
            
            End If
            
        Next
        '################################################################


'        If objItemNF.iAlmoxarifado <> 0 Then

            'Lê item de movimento de estoque
            objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
            objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
            objItemMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
            
            Set colMovEstoque = New Collection
            
            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
            If lErro <> SUCESSO Then gError 75937
            
            '############################################################
            'Inserido por Wagner 02/06/2006
            'Le os movimento gerados por grade ou componentes de um kit de venda
            lErro = CF("MovEstoque_Le_ItemNF_Grade", objItemMovEstoque, colMovEstoque)
            If lErro <> SUCESSO Then gError 180048
            '############################################################
            
            For Each objItemMovEstoque In colMovEstoque
            
                'se forem os movimentos de transferencia de material consignado ==> não trata-os, pois a venda de material consignado está sendo tratada
                If objItemMovEstoque.iTipoMov <> MOV_EST_SAIDA_TRANSF_CONSIG_TERC And _
                   objItemMovEstoque.iTipoMov <> MOV_EST_ENTRADA_TRANSF_DISP1 Then
            
                    'Lê o Almoxarifado
                    objAlmoxarifado.iCodigo = objItemMovEstoque.iAlmoxarifado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError 75935
            
                    'Se não encontrou Almoxarifado --> Erro
                    If lErro = 25056 Then gError 75938
            
                    'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
                    Set colRastreamentoMovto = New Collection
                    
                    lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
                    If lErro <> SUCESSO Then gError 75934
            
                    Set objItemNF.colRastreamento = New Collection
'                    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 10/03/2006
             
                    'Guarda as quantidades alocadas dos lotes
                    For Each objRastreamentoMovto In colRastreamentoMovto
            
                        iLinha = iLinha + 1
                        
                        gcolcolRastreamentoSerie.Add objRastreamentoMovto.colRastreamentoSerie 'Inserido por Wagner 10/03/2006
            
                        'procura RastreamentoLote
'                         objRastreamentoLote.sProduto = objItemNF.sProduto 'O produto Rastreado pode não ser o item da Nota associado (Grade e Kit de Venda)
                        objRastreamentoLote.sProduto = objRastreamentoMovto.sProduto
                        objRastreamentoLote.iFilialOP = objRastreamentoMovto.iFilialOP
                        objRastreamentoLote.sCodigo = objRastreamentoMovto.sLote
                        
                        lErro = CF("RastreamentoLote_Le", objRastreamentoLote)
                        If lErro <> SUCESSO And lErro <> 75710 Then gError 75936
            
                        'Se não encontrou, erro
                        If lErro = 75710 Then gError 75939
            
                        'Lê FilialEmpresa
                        If objRastreamentoMovto.iFilialOP <> 0 Then
                        
                            objFilialEmpresa.iCodFilial = objRastreamentoMovto.iFilialOP
                            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                            If lErro <> SUCESSO And lErro <> 27378 Then gError 75696
                
                            'Se não encontrou a FilialEmpresa
                            If lErro = 27378 Then gError 75697
            
                        End If
            
                        'Formata o Produto
'                        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                        lErro = Mascara_RetornaProdutoEnxuto(objRastreamentoMovto.sProduto, sProdutoEnxuto)
                        If lErro <> SUCESSO Then gError 83198
                
                        'Tratamento diferenciado para quando o Produto for uma Combo
                        If TypeName(ProdutoRastro) <> "ComboBox" Then
                            ProdutoRastro.PromptInclude = False
                            ProdutoRastro.Text = sProdutoEnxuto
                            ProdutoRastro.PromptInclude = True
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = ProdutoRastro.Text
                        Else
                            lErro = Mascara_RetornaProdutoTela(objRastreamentoMovto.sProduto, sProdutoEnxuto)
                            If lErro <> SUCESSO Then gError 188057
                        
                            Call CF("sCombo_Seleciona2", ProdutoRastro, sProdutoEnxuto)
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = sProdutoEnxuto
                        End If
            
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objItemNF.iItem
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objAlmoxarifado.sNomeReduzido
                        If bEscaninhoVisivel = True Then
                            If objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG Then
                                objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_CONSIGNADO
                            Else
                                objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                            End If
                        End If
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantRastro_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastreamentoMovto.sLote
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteDataRastro_Col) = Format(objRastreamentoLote.dtDataEntrada, "dd/mm/yyyy")
                        If objRastreamentoMovto.iFilialOP <> 0 Then
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col) = CStr(objRastreamentoMovto.iFilialOP) & SEPARADOR & objFilialEmpresa.sNome
                        End If
                        
                        '###########################################################
                        'Alterado por Wagner 15/03/2006
                        Set objProduto = New ClassProduto
                        
                        objProduto.sCodigo = objRastreamentoMovto.sProduto
        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 141952
                        
                        If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantLoteRastro_Col) = Formata_Estoque(objRastreamentoMovto.dQuantidadeSerie)
                        Else
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantLoteRastro_Col) = Formata_Estoque(objRastreamentoMovto.dQuantidade)
                        End If
                        '###########################################################
                        
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_UMRastro_Col) = objRastreamentoMovto.sSiglaUM
                        
                        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
                    
                    Next
                    
                End If
    
            Next
            
 '       End If
        
    Next
    
    Carrega_RastroItensNF = SUCESSO

    Exit Function

Erro_Carrega_RastroItensNF:

    Carrega_RastroItensNF = gErr

    Select Case gErr

        Case 75934, 75935, 75936, 75937, 83199, 141952, 180046, 180047, 180048, 188057

        Case 75938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)

        Case 75939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO1", gErr, objRastreamentoLote.sProduto, objRastreamentoLote.sCodigo, objRastreamentoLote.iFilialOP)

        Case 83198
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)

        Case 83200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152757)
    
    End Select

    Exit Function

End Function

Public Function Saida_Celula() As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    Select Case objGridInt.objGrid.Col
    
    
        Case iGrid_ItemNFRastro_Col
            lErro = Saida_Celula_ItemNFRastro()
            If lErro <> SUCESSO Then gError 83202
    
        'Almoxarifado Rastreamento
        Case iGrid_AlmoxRastro_Col
            lErro = Saida_Celula_AlmoxRastro()
            If lErro <> SUCESSO Then gError 83161
        
        'Escaninho Rastreamento
        Case iGrid_EscaninhoRastro_Col
            lErro = Saida_Celula_EscaninhoRastro()
            If lErro <> SUCESSO Then gError 83313
        
        'Lote Rastreamento
        Case iGrid_LoteRastro_Col
            lErro = Saida_Celula_LoteRastro()
            If lErro <> SUCESSO Then gError 83162
    
        'FilialOP Rastreamento
        Case iGrid_FilialOPRastro_Col
            lErro = Saida_Celula_FilialOPRastro()
            If lErro <> SUCESSO Then gError 83163
        
        'QuantLoteRastro Rastreamento
        Case iGrid_QuantLoteRastro_Col
            lErro = Saida_Celula_QuantLoteRastro()
            If lErro <> SUCESSO Then gError 83164
    
    '#############################################
    'Inserido por Wagner 29/05/2006
        Case iGrid_ProdutoRastro_Col
            lErro = Saida_Celula_ProdutoRastro()
            If lErro <> SUCESSO Then gError 180049
    '##############################################
    
    End Select

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 83160, 83161, 83162, 83163, 83164, 83202, 83313, 180049

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152758)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemNFRastro() As Long
'Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iLinha As Integer
Dim iAchou As Integer
Dim iItem As Integer
Dim iItem_Atual As Integer
Dim sProdutoFormatado As String
Dim objNFiscal As New ClassNFiscal
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colRastreamentoSerie As New Collection
Dim iIndice2 As Integer
Dim bTemProdutoRastreado As Boolean
Dim sProdutoMascarado As String
Dim bSemGrade As Boolean, sProd1 As String, sProd2 As String

On Error GoTo Erro_Saida_Celula_ItemNFRastro

    Set objGridInt.objControle = ItemNFRastro

    'Verifica se o Produto esta preenchido
    If Len(Trim(ItemNFRastro.Text)) = 0 And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then gError 83225
    
    If Len(Trim(ItemNFRastro.Text)) > 0 Then

        iItem = CInt(ItemNFRastro.Text)
        
        If iItem < 1 Or iItem > objTela.objGridItens.iLinhasExistentes Then gError 83201

        'Se o produto não está preenchido ==> erro
        If Len(Trim(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col))) = 0 Then gError 83249
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83250
            
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83251
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83252
                
        bSemGrade = True
        If Not (objTela.gobjNFiscal Is Nothing) Then
            If objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Count <> 0 Then bSemGrade = False
        End If
                
        'se não for um produto rastreavel ==> erro
        If objProduto.iRastro = PRODUTO_RASTRO_NENHUM And bSemGrade Then gError 83253

        'descobre qual o item atual
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
            iItem_Atual = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))
        End If
        
        'se o item que se está preenchendo é diferente do item atual, ==> limpa os campos de lote
        If iItem <> iItem_Atual Then
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""
'mario distribuicao alterar para contemplar as nf de entrada com o tab de distribuicao
            If bEscaninhoVisivel = True Then objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ""
            'coloca nos campos produto e unidade de medida os valores do grid de itens
            
            If bSemGrade Then
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col)
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_UnidadeMed_Col)
            Else
                            
                bTemProdutoRastreado = False
                
                ProdutoRastro.Clear
                    
                For iIndice2 = 1 To objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Count
                
                    'Lê os demais atributos do Produto
                    objProduto.sCodigo = objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Item(iIndice2).sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 177973
                
                    'Se o produto for rastreado = > Adiciona na Combo
                    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
                    
                        lErro = Mascara_RetornaProdutoTela(objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Item(iIndice2).sProduto, sProdutoMascarado)
                        If lErro <> SUCESSO Then gError 177977
                    
                        If Not bTemProdutoRastreado Then
                        
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) = sProdutoMascarado
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objProduto.sSiglaUMEstoque
                        
                            sProduto = objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Item(iIndice2).sProduto
                        
                        End If
                                                
                        ProdutoRastro.AddItem sProdutoMascarado
                    
                        bTemProdutoRastreado = True
                    End If
                
                Next
                           
                If Not bTemProdutoRastreado Then gError 177974
            
            End If
            
            If iCodigo <> 0 Then
            
                objTipoDocInfo.iCodigo = iCodigo
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 89515
            
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 89516
            
            Else
                objTipoDocInfo.iTipo = -1
            End If
            
            If bEscaninhoVisivel = False And iCodigo <> DOCINFO_NFISFBF And iCodigo <> DOCINFO_NFISFBFPV And iCodigo <> DOCINFO_NFISPC And iCodigo <> DOCINFO_NFFISPC And objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFIE And objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFEXT And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalRemPedido" And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalRem" And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalDev" Then
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Almoxarifado_Col)
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)
            Else
            
                'mario distribuicao. Se for uma nota fiscal de entrada pegar o almoxarifado e a quantidade do grid de distribuicao
                If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRem" Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Then
        
                    'Formata o Produto para o BD
                    lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProd1, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 83250
        
                    'verifica se o item está no grid de alocacoes
                    For iLinha = 1 To objTela.gobjDistribuicao.objGridDist.iLinhasExistentes
                    
                        'Formata o Produto para o BD
                        lErro = CF("Produto_Formata", objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProd2, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 83250
                    
                        If iItem = CInt(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_ItemNFDist_Col)) And (bSemGrade Or sProd1 = sProd2) Then
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_QuantDist_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_UMDist_Col)
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    'se não tem alocação para este item
                    If iAchou = 0 Then gError 89517
                    
                Else
            
                    'se o campo escaninho estiver visivel
                
                    If bEditavel = False Then
                        If Len(Trim(Serie.Text)) = 0 Then gError 83353
                        
                        If bTelaManutencao = False Then
                            If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83354
                            objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
                            'If Len(Trim(NFiscal.Caption)) = 0 Then gError 83355
                            objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
                        Else
                            If Len(Trim(LblDataEmissao.Caption)) = 0 Then gError 83354
                            objNFiscal.dtDataEmissao = CDate(LblDataEmissao.Caption)
                            If Len(Trim(Numero.Text)) = 0 Then gError 83355
                            objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
                        End If
                        
                        objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value) 'Serie.Text
                    End If
                    
                    iAchou = 0
                    
                    'Formata o Produto para o BD
                    lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProd1, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 83250
                    
                    'verifica se o item está no grid de alocacoes
                    For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
                        
                         'Formata o Produto para o BD
                        lErro = CF("Produto_Formata", objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_ProdutoAloc_Col), sProd2, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 83250
                            
                        If iItem = CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_Item_Col)) And (bSemGrade Or sProd1 = sProd2) Then
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_AlmoxAloc_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_UMAloc_Col)
                            If iCodigo = DOCINFO_NFISFBF Or iCodigo = DOCINFO_NFISFBFPV Or iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRemPedido" Then objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_QuantAloc_Col)
                            iAchou = 1
                            Exit For
                        End If
                    Next
                    
                    'se não tem alocação para este item
                    If iAchou = 0 Then gError 83348
                    
                    'Lê o Almoxarifado
                    objAlmoxarifado.sNomeReduzido = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col)
                    
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 83349
            
                    'Se não encontrou Almoxarifado --> Erro
                    If lErro = 25060 Then gError 83350
                    
                    If bEditavel = False Then
                    
                        'Se é uma NF gravada lê o movimento de estoque, senão pega do grid de alocação
                        If objNFiscal.lNumNotaFiscal > 0 Then
                        
                            'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                            lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo)
                            If lErro <> SUCESSO And lErro <> 83345 Then gError 83346
                            
                            If lErro <> SUCESSO Then
                                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                                lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo, sProduto)
                                If lErro <> SUCESSO And lErro <> 83345 Then gError 180050
                            End If
                            
                            If lErro = SUCESSO Then
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                            Else
                                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                                lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo)
                                If lErro <> SUCESSO And lErro <> 83345 Then gError 83351
                                
                                If lErro <> SUCESSO Then
                                    'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                                    lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo, sProduto)
                                    If lErro <> SUCESSO And lErro <> 83345 Then gError 180051
                                End If
                                
                                'se não houver movimento de estoque para este item/almoxarifado ==> erro
                                If lErro = 83345 Then gError 83352
                        
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_CONSIGNADO
                                
                            End If
                            
                        Else
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                            'verifica se o item está no grid de alocacoes
                            For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
                                If iItem = CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_Item_Col)) And (bSemGrade Or objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_ProdutoAloc_Col)) Then
                                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_QuantAloc_Col)
                                    Exit For
                                End If
                            Next
                        End If
            
                    End If
                    
                End If
            
            End If

        End If
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 10/03/2006
        End If
        
    Else

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""

        'coloca nos campos produto, almoxaridado e quantrastro os valores do grid de itens
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = ""
        If bEscaninhoVisivel = True Then objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83169

    Saida_Celula_ItemNFRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFRastro:

    Saida_Celula_ItemNFRastro = gErr

    Select Case gErr
        
        Case 83169, 83207, 83250, 83251, 83346, 83349, 83351, 89515, 177973, 177977, 180050, 180051
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83201, 83225
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMRASTRO_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_ITENS", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83253, 177974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, objProduto.sCodigo, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83348, 89517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMRASTRO_NAO_ALOCACAO", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_MOVESTOQUE_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152759)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoRastro() As Long
'Rastreamento
'Função alterada por Wagner para permitir rastreamento de Grade e Kit de Venda

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iLinha As Integer
Dim iAchou As Integer
Dim sProdutoAux As String

On Error GoTo Erro_Saida_Celula_ProdutoRastro

    Set objGridInt.objControle = ProdutoRastro

    '###############################################
    'Inserido por Wagner 02/06/2006
    If TypeName(ProdutoRastro) <> "ComboBox" Then
        sProdutoAux = ProdutoRastro.ClipText
    Else
        sProdutoAux = ProdutoRastro.Text
    End If
    '###############################################
    
    'Verifica se o Produto esta preenchida
    If Len(Trim(sProdutoAux)) > 0 Then

        lErro = CF("Produto_Critica", ProdutoRastro.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 83165
        
        If lErro = 25041 Then gError 83166
                
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 83167
        
        If TypeName(ProdutoRastro) <> "ComboBox" Then
            ProdutoRastro.PromptInclude = False
            ProdutoRastro.Text = sProduto
            ProdutoRastro.PromptInclude = True
        End If

        If ProdutoRastro.Text <> objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) Then
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = ""
        End If
    
    Else
        
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = ""

    End If

    '###############################################
    'Inserido por Wagner 29/05/2006
    lErro = ProdutoRastro_Saida_Celula(objProduto)
    If lErro <> SUCESSO Then gError 177961
    '###############################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83169

    Saida_Celula_ProdutoRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoRastro:

    Saida_Celula_ProdutoRastro = gErr

    Select Case gErr
        
        Case 83165, 83169, 177961
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83166
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoRastro.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 83167
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83170
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_AUSENTE_GRID_ITENS", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83171
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ALMOX_AUSENTE_GRID_ITENS", gErr, objProduto.sCodigo, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152760)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AlmoxRastro() As Long
'Rastreamento

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim iLinha As Integer
Dim iLinhaRastro As Integer
Dim iAchou As Integer
Dim iItem As Integer
Dim objNFiscal As New ClassNFiscal
Dim dQuantidade As Double
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Saida_Celula_AlmoxRastro

    Set objGridInt.objControle = AlmoxRastro

    'Se o Almoxarifado está preenchido
    If Len(Trim(AlmoxRastro.Text)) > 0 Then

        If Trim(AlmoxRastro.Text) <> objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) Then

            iItem = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))
    
            lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83314
    
            'Valida o ALmoxarifado
            lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, AlmoxRastro, objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 83315
    
            'Se não for encontrado --> Erro
            If lErro = 25157 Then gError 83316
            
            If lErro = 25162 Then gError 83317
    
            If bEditavel = False Then
    
                If Len(Trim(Serie.Text)) = 0 Then gError 83356
            
                If bTelaManutencao = False Then
                    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83357
                    objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
                    If Len(Trim(NFiscal.Caption)) = 0 Then gError 83358
                    objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
                Else
                    If Len(Trim(LblDataEmissao.Caption)) = 0 Then gError 83357
                    objNFiscal.dtDataEmissao = CDate(LblDataEmissao.Caption)
                    If Len(Trim(Numero.Text)) = 0 Then gError 83358
                    objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
                End If
            
                objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value) 'Serie.Text
            
            End If
            
            If iCodigo <> 0 Then
            
                objTipoDocInfo.iCodigo = iCodigo
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 89520
            
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 89521
            
            Else
                objTipoDocInfo.iTipo = -1
            End If
            
            iAchou = 0
            
            'mario distribuicao. Se for uma nota fiscal de entrada pegar o almoxarifado e a quantidade do grid de distribuicao
            If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
        
                'verifica se o item está no grid de alocacoes
                For iLinha = 1 To objTela.gobjDistribuicao.objGridDist.iLinhasExistentes
                    If iItem = CInt(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_ItemNFDist_Col)) And AlmoxRastro.Text = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col) Then
                        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_QuantDist_Col)
                        iAchou = 1
                        Exit For
                    End If
                Next
                
                'se não tem alocação para este item
                If iAchou = 0 Then gError 89526
                
            Else
            
            
                'verifica se o item está no grid de alocacoes
                For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
                    If iItem = CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_Item_Col)) And _
                    AlmoxRastro.Text = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_AlmoxAloc_Col) Then
                        If iCodigo = DOCINFO_NFISFBF Or iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Then objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_QuantAloc_Col)
                        iAchou = 1
                        Exit For
                    End If
                Next
            
                'se não tem alocação para este item
                If iAchou = 0 Then gError 83318
            
            End If
            
            If bEditavel = False Then
            
                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo)
                If lErro <> SUCESSO And lErro <> 83345 Then gError 83359
                
                If lErro = SUCESSO Then
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                Else
                    'Lê a quantidade do movimento de estoque de material consignado referente ao item e escaninho da nota fiscal passados como parametros
                    lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo)
                    If lErro <> SUCESSO And lErro <> 83345 Then gError 83360
                
                    'se não houver movimento de estoque para este item/almoxarifado ==> erro
                    If lErro = 83345 Then gError 83361
            
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_CONSIGNADO
                    
                End If
        
            End If
            
        End If
        
    Else
    
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83174

    Saida_Celula_AlmoxRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxRastro:

    Saida_Celula_AlmoxRastro = gErr

    Select Case gErr

        Case 83174, 83314, 83315, 83359, 83360, 89520
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83316, 83317
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, AlmoxRastro.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83318, 89526
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_AUSENTE_GRID_ALOCACAO", gErr, AlmoxRastro.Text, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            
        Case 83358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 83361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_MOVESTOQUE_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152761)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantLoteRastro() As Long

Dim lErro As Long
Dim dQuantidadeAnterior As Double 'Inserido por Wagner 10/03/2006

On Error GoTo Erro_Saida_Celula_QuantLoteRastro

    Set objGridInt.objControle = QuantLoteRastro

    'Se a quantidade alocada do lote foi preenchida
    If Len(Trim(QuantLoteRastro.ClipText)) > 0 Then
    
        dQuantidadeAnterior = StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)) 'Inserido por Wagner 10/03/2006

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantLoteRastro.Text)
        If lErro <> SUCESSO Then gError 83175

        'Se a quantidade alocada do lote for maior que a quantidade disponível, erro
        If StrParaDbl(QuantLoteRastro.Text) > StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col)) Then gError 83176
        
        QuantLoteRastro.Text = Formata_Estoque(QuantLoteRastro.Text)

    End If

    '###############################################################
    'Inserido por Wagner 10/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(objGridInt.objGrid.Row), StrParaDbl(QuantLoteRastro.Text), objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), dQuantidadeAnterior, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), objGridInt.objGrid.Row)
    If lErro <> SUCESSO Then gError 141836
    '###############################################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83177

    Saida_Celula_QuantLoteRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantLoteRastro:

    Saida_Celula_QuantLoteRastro = gErr

    Select Case gErr

        Case 83175, 83177, 141836
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTLOTE_MAIOR_QUANTALM", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_EscaninhoRastro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iLinhaRastro As Integer
Dim iAchou As Integer
Dim iItem As Integer
Dim objNFiscal As New ClassNFiscal
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_EscaninhoRastro

    Set objGridInt.objControle = EscaninhoRastro

    If Len(Trim(EscaninhoRastro.Text)) > 0 Then

        iItem = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))

        If Len(Trim(Serie.Text)) = 0 Then gError 83362
        
        If bTelaManutencao = False Then
            If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83363
            objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
            'If Len(Trim(NFiscal.Caption)) = 0 Then gError 83364
            objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
        Else
            If Len(Trim(LblDataEmissao.Caption)) = 0 Then gError 83363
            objNFiscal.dtDataEmissao = CDate(LblDataEmissao.Caption)
            If Len(Trim(Numero.Text)) = 0 Then gError 83364
            objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
        End If
        
        objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value) 'Serie.Text
        
        iAchou = 0
                
        'verifica se o item está no grid de alocacoes
        For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
            If iItem = CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_Item_Col)) And _
               objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_AlmoxAloc_Col) Then
                iAchou = 1
                Exit For
            End If
        Next
                
        'se não tem alocação para este item
        If iAchou = 0 Then gError 83365
        
        'Lê o Almoxarifado
        objAlmoxarifado.sNomeReduzido = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col)
        
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 83366

        'Se não encontrou Almoxarifado --> Erro
        If lErro = 25060 Then gError 83367
        
        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 180168
        
        If objNFiscal.lNumNotaFiscal > 0 Then
            
            If EscaninhoRastro.Text = ESCANINHO_RASTRO_DISPONIVEL Then
            
                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo)
                If lErro <> SUCESSO And lErro <> 83345 Then gError 83368
            
                If lErro <> SUCESSO Then
                    'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                    lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo, sProduto)
                    If lErro <> SUCESSO And lErro <> 180061 Then gError 180170
                End If
                        
                If lErro <> SUCESSO Then gError 83369
            
            
            Else
                
                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo)
                If lErro <> SUCESSO And lErro <> 83345 Then gError 83370
            
                If lErro <> SUCESSO Then
                    'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                    lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo, sProduto)
                    If lErro <> SUCESSO And lErro <> 180061 Then gError 180171
                End If
                            
                If lErro <> SUCESSO Then gError 83371
            
            End If
            
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
       
        End If

    Else

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83319

    Saida_Celula_EscaninhoRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_EscaninhoRastro:

    Saida_Celula_EscaninhoRastro = gErr

    Select Case gErr

        Case 83319, 83366, 83368, 83370, 180168, 180170, 180171
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMRASTRO_NAO_ALOCACAO", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_MOVESTOQUE_DISP_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_MOVESTOQUE_CONSIG_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152763)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_LoteRastro() As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iLinha As Integer
Dim sLoteAnterior As String 'Inserido por Wagner 10/03/2006

On Error GoTo Erro_Saida_Celula_LoteRastro

    Set objGridInt.objControle = LoteRastro
        
    'Se o lote foi preenchido
    If Len(Trim(LoteRastro.Text)) > 0 Then
    
        sLoteAnterior = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) 'Inserido por Wagner 10/03/2006
        
        'Se o produto não está preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col))) = 0 Then gError 83178
        
        'Se o almoxarifado não está preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))) = 0 Then gError 83179
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83180
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83181
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83182
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objRastroLote.dtDataEntrada = DATA_NULA
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    'se se tratar do mesmo par Produto/almoxarifado/lote
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text Then
                            If bEscaninhoVisivel = False Then
                                gError 83183
                            Else
                                If objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) Then gError 83372
                            End If
                        End If
                    End If
                Next
                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 83184
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 83185
                
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
                
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
                    objRastroLote.iFilialOP = giFilialEmpresa
                Else
                    objRastroLote.iFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col))
                End If
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        'se o par itemnf/almoxarifado/lote/op já está no grid
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objRastroLote.iFilialOP Then
                            'se a OP ainda não estava preenchida
                            If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
                                objRastroLote.iFilialOP = 0
                                objRastroLote.dtDataEntrada = DATA_NULA
                            Else
                                If bEscaninhoVisivel = False Then
                                    gError 83186
                                Else
                                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) Then gError 83373
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next
                
                If objRastroLote.iFilialOP <> 0 Then
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 83187
                
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then
                    
                        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
                            objRastroLote.iFilialOP = 0
                        Else
                            gError 83188
                        End If
                    End If
                
'                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
        
                End If
                
            '##############################################################
            'Inserido por Wagner 10/03/2006
            'Se o produto possuir rastro por série
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
               For iLinha = 1 To objGridInt.iLinhasExistentes
                    'se se tratar do mesmo par itemnf/almoxarifado/lote
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text Then
                            If bEscaninhoVisivel = False Then
                                gError 83183
                            Else
                                If objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) Then gError 83372
                            End If
                        End If
                    End If
                Next
                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 83184
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 83185
                
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
                
                'Carrega as séries na coleção global
                lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(objGridInt.objGrid.Row), StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)), LoteRastro.Text, StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)), sLoteAnterior, objGridInt.objGrid.Row)
                If lErro <> SUCESSO Then gError 141835
            '###############################################################
        
            End If
    
            'Preenche campos do lote
            lErro = Lote_Saida_Celula(objRastroLote)
            If lErro <> SUCESSO Then gError 83189
    
        End If
    
    Else
    
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83190

    Saida_Celula_LoteRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_LoteRastro:

    Saida_Celula_LoteRastro = gErr

    Select Case gErr

        Case 83180, 83181, 83184, 83187, 83189, 83190, 141835
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, objGridInt.objGrid.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO_GRID", gErr, objGridInt.objGrid.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PROD_ALMOX_JA_UTILIZADO_GRID", gErr, LoteRastro.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83185
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 83186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEOP_PROD_ALMOX_JA_UTILIZADO_GRID", gErr, LoteRastro.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)), objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83188
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_FILIALOP_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 83372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PROD_ALMOX_JA_UTILIZADO_GRID", gErr, LoteRastro.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEOP_PROD_ALMOX_JA_UTILIZADO_GRID", gErr, LoteRastro.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)), objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152764)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOPRastro() As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objFilialEmpresa As New AdmFiliais
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_FilialOPRastro

    Set objGridInt.objControle = FilialOPRastro
        
    'Se a filial foi preenchida
    If Len(Trim(FilialOPRastro.Text)) > 0 Then
        
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 83030

        'Se não for encontrado --> Erro
        If lErro = 71971 Then gError 83031
        If lErro = 71972 Then gError 83032
        
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)) <> 0 Then
        
            For iLinha = 1 To objGridInt.iLinhasExistentes
                If iLinha <> objGridInt.objGrid.Row Then
                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
                       objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objFilialEmpresa.iCodFilial Then
                        If bEscaninhoVisivel = False Then
                            gError 83033
                        Else
                            If objGridInt.objGrid.TextMatrix(iLinha, iGrid_EscaninhoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) Then gError 83374
                        End If
                    End If
                End If
            Next
        
            'Formata o produto
            lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83034
        
            objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)
            objRastroLote.sProduto = sProdutoFormatado
            objRastroLote.iFilialOP = objFilialEmpresa.iCodFilial
                
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le", objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 83035
                
            If lErro = 75710 Then gError 83264
                
            If objRastroLote.dtDataEntrada <> DATA_NULA And lErro = SUCESSO Then
                GridRastro.TextMatrix(GridRastro.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
            End If
            
            FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        Else
            
            FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        End If
        
    Else
    
        GridRastro.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83037

    Saida_Celula_FilialOPRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOPRastro:

    Saida_Celula_FilialOPRastro = gErr

    Select Case gErr

        Case 83030, 83034, 83035, 83037
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83031, 83032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, FilialOPRastro.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83264
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_FILIALOP_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
            
        Case 83374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152765)

    End Select

    Exit Function

End Function

Private Function Lote_Saida_Celula(objRastroLote As ClassRastreamentoLote) As Long
'Executa a saida de celula do campo lote, o tratamento dos erros do Grid é feita na rotina chamadora

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Lote_Saida_Celula

    If objRastroLote.dtDataEntrada <> DATA_NULA Then
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
    End If
    
    'Se a filial empresa foi preenchida
    If objRastroLote.iFilialOP <> 0 Then
        
        objFilialEmpresa.iCodFilial = objRastroLote.iFilialOP
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 83152

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 83153

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
    
    End If
    
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Lote_Saida_Celula = SUCESSO
    
    Exit Function
        
Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
        
        Case 83152
        
        Case 83153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152766)
    
    End Select
    
    Exit Function
    
End Function

Public Function Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name
    
        Case ItemNFRastro.Name
    
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83334
    
            If lErro = SUCESSO Then ItemNFRastro.Enabled = True
    
        Case AlmoxRastro.Name
        
            If iCodigo <> 0 Then
            
                objTipoDocInfo.iCodigo = iCodigo
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 89518
            
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 89519
            
            Else
                objTipoDocInfo.iTipo = -1
            End If
            
            If bEscaninhoVisivel = True Or iCodigo = DOCINFO_NFISFBF Or iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRemPedido" Then
        
                lErro = Testa_Control_Enable(objControl)
                If lErro <> SUCESSO And lErro <> 1 Then gError 83333
            
                If lErro = SUCESSO Then
            
                    If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
                        AlmoxRastro.Enabled = True
                    Else
                        AlmoxRastro.Enabled = False
                    End If
                
                End If
    
            End If
    
        Case EscaninhoRastro.Name
        
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83335
        
            If lErro = SUCESSO Then
        
                If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 And _
                   Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col))) > 0 Then
                    EscaninhoRastro.Enabled = True
                Else
                    EscaninhoRastro.Enabled = False
                End If
    
            End If
    
        Case LoteRastro.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
                    LoteRastro.Enabled = True
                Else
                    LoteRastro.Enabled = False
                End If
            
            End If
            
        '##########################################################
        'Inserido por Wagner 29/05/2006
        Case ProdutoRastro.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
                    
                    If objTela.gobjNFiscal Is Nothing Then
                        ProdutoRastro.Enabled = False
                    Else
                        If objTela.gobjNFiscal.colItensNF.Item(StrParaInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))).colItensRomaneioGrade.Count = 0 Then
                            ProdutoRastro.Enabled = False
                        Else
                            ProdutoRastro.Enabled = True
                        End If
                    End If
                Else
                    ProdutoRastro.Enabled = False
                End If
            
            End If
        '##########################################################
            
        Case QuantLoteRastro.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83337
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 And _
                   Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col))) > 0 Then
                    QuantLoteRastro.Enabled = True
                Else
                    QuantLoteRastro.Enabled = False
                End If
            
            End If
            
        Case FilialOPRastro.Name

            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83338
        
            If lErro = SUCESSO Then

                lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 83193
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                    objProduto.sCodigo = sProdutoFormatado
            
                    'Lê o Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 83194
            
                    'Não achou o Produto
                    If lErro = 28030 Then gError 83195
            
                    If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                        FilialOPRastro.Enabled = True
                    Else
                        FilialOPRastro.Enabled = False
                    End If
            
                End If

            End If

    End Select
    
    Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable:

    Rotina_Grid_Enable = gErr

    Select Case gErr

        Case 83193, 83194, 83333, 83334, 83335, 83336, 83337, 89518

        Case 83195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 89519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152767)

    End Select

    Exit Function

End Function

Private Function Testa_Control_Enable(objControl As Object) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Testa_Control_Enable

    Testa_Control_Enable = SUCESSO

'    'se se tratar de uma nota fiscal de saida
'    If bSaidaMaterial = True And bDevolucaoMaterial = False And bTelaManutencao = False Then
'        If bEditavel = False Then
'            'se o numero da nota não estiver preenchida ==> nota nova ==> não permitir digitar o rastreamento, só para alteracoes
'            If Len(Trim(NFiscal.Caption)) = 0 Then
'                objControl.Enabled = False
'                Testa_Control_Enable = 1
'            End If
'        End If
'    End If
    
    If Testa_Control_Enable = SUCESSO Then
    
        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92248
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            objProduto.sCodigo = sProdutoFormatado
    
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 92246
    
            'Não achou o Produto
            If lErro = 28030 Then gError 92247
    
            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
                Testa_Control_Enable = 1
            End If
            
        End If
    
    End If
    
    
    Exit Function

Erro_Testa_Control_Enable:

    Testa_Control_Enable = gErr

    Select Case gErr

        Case 92246, 92248

        Case 92247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152768)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_Rastreamento() As Long
'Inicializa o Grid de Rastreamento

Dim iIndice As Integer
Dim iEscaninho As Integer
Dim lErro As Long

    Set objGridInt = New AdmGrid

    Set objGridInt.objForm = objTela

    If bTelaManutencao = True Then
        Set objUserControl = objTela
    Else
        Set objUserControl = objTela.objUserControl
    End If
    
    If TypeName(ProdutoRastro) <> "ComboBox" Then
    
        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoRastro)
        If lErro <> SUCESSO Then gError 83844
        
    End If
    
    If EscaninhoRastro.Enabled = True Then
        AlmoxRastro.Enabled = True
        bEscaninhoVisivel = True
    Else
        AlmoxRastro.Enabled = False
        bEscaninhoVisivel = False
    End If
   

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("ItemNF")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Almoxarifado")
    If bEscaninhoVisivel = True Then objGridInt.colColuna.Add ("Escaninho")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Qtd. Almoxarifado")
    objGridInt.colColuna.Add ("Lote\Série Ini.") 'Alterado por Wagner
    objGridInt.colColuna.Add ("FilialOP do Lote")
    objGridInt.colColuna.Add ("Data do Lote")
    objGridInt.colColuna.Add ("Qtd. Alocada Lote")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemNFRastro.Name)
    objGridInt.colCampo.Add (ProdutoRastro.Name)
    objGridInt.colCampo.Add (AlmoxRastro.Name)
    If bEscaninhoVisivel = True Then objGridInt.colCampo.Add (EscaninhoRastro.Name)
    objGridInt.colCampo.Add (UMRastro.Name)
    objGridInt.colCampo.Add (QuantRastro.Name)
    objGridInt.colCampo.Add (LoteRastro.Name)
    objGridInt.colCampo.Add (FilialOPRastro.Name)
    objGridInt.colCampo.Add (LoteDataRastro.Name)
    objGridInt.colCampo.Add (QuantLoteRastro.Name)

    'Colunas da Grid
    
    iGrid_ItemNFRastro_Col = 1
    iGrid_ProdutoRastro_Col = 2
    iGrid_AlmoxRastro_Col = 3
    If bEscaninhoVisivel = True Then
        iGrid_EscaninhoRastro_Col = 4
        iEscaninho = 1
    End If
    iGrid_UMRastro_Col = 4 + iEscaninho
    iGrid_QuantRastro_Col = 5 + iEscaninho
    iGrid_LoteRastro_Col = 6 + iEscaninho
    iGrid_FilialOPRastro_Col = 7 + iEscaninho
    iGrid_LoteDataRastro_Col = 8 + iEscaninho
    iGrid_QuantLoteRastro_Col = 9 + iEscaninho

    'Grid do GridInterno
    objGridInt.objGrid = GridRastro

    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE


    If bEscaninhoVisivel = False Then

        'Linhas visíveis
        objGridInt.iLinhasVisiveis = 7

    Else
    
        'Linhas visíveis
        objGridInt.iLinhasVisiveis = 7
        
    End If
    
    If bTelaManutencao = False Then objGridInt.iLinhasVisiveis = objGridInt.iLinhasVisiveis + 2

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Rastreamento = SUCESSO

    Exit Function

End Function

Private Sub Class_Initialize()

        Set objEventoLoteRastro = New AdmEvento
        Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 10/03/2006
        bSaidaMaterial = False
        bDevolucaoMaterial = False
        bEditavel = True
        
End Sub

Private Sub Class_Terminate()
    Set objEventoLoteRastro = Nothing
    Set objGridInt = Nothing
    Set objTela = Nothing
    Set objUserControl = Nothing
    Set gcolcolRastreamentoSerie = Nothing 'Inserido por Wagner 10/03/2006
End Sub

Public Sub BotaoLotes_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo

On Error GoTo Erro_BotaoLotes_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If objGridInt.objGrid.Row = 0 Then gError 83146
        
    'Se o produto não foi preenchido, erro
    If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col))) = 0 Then gError 83147
        
    'Formata o produto
    lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83148
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 83149
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 83150
        
    'Verifica o tipo de rastreamento do produto
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        sSelecao = " FilialOP <> ? AND Produto = ?"
        
    '#########################################################
    'Inserido por Wagner 10/03/2006
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    '#########################################################
    
    End If
    
    'Adiciona filtros
    colSelecao.Add 0
    colSelecao.Add sProdutoFormatado
    
    If bSaidaMaterial = True Then
         
        objRastroLoteSaldo.sProduto = sProdutoFormatado
        objRastroLoteSaldo.iFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col))
        objRastroLoteSaldo.sLote = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)
        
        'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
        Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoLoteRastro, sSelecao)
                    
    Else
    
        objRastroLote.sProduto = sProdutoFormatado
        objRastroLote.iFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col))
        objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)
        
        'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
        Call Chama_Tela("RastroLoteLista", colSelecao, objRastroLote, objEventoLoteRastro, sSelecao)
    
    End If
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case 83146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 83147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 83148, 83149
        
        Case 83150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152769)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoLoteRastro_evSelecao(obj1 As Object)
'Rastreamento

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLoteAnterior As String 'Inserido por Wagner 10/03/2006

On Error GoTo Erro_objEventoLoteRastro_evSelecao

    If bSaidaMaterial = True Then
        Set objRastroLoteSaldo = obj1
        objRastroLote.sProduto = objRastroLoteSaldo.sProduto
        objRastroLote.sCodigo = objRastroLoteSaldo.sLote
        objRastroLote.iFilialOP = objRastroLoteSaldo.iFilialOP
        objRastroLote.dtDataEntrada = objRastroLoteSaldo.dtDataEntrada
    Else
        Set objRastroLote = obj1
    End If
    
    'Se a Linha corrente for diferente da Linha fixa
    If objGridInt.objGrid.Row <> 0 Then
    
        sLoteAnterior = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) 'Inserido por Wagner 10/03/2006
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83154
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83155
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83156
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col) Then gError 83157
                    End If
                Next
        
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And _
                           Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objRastroLote.iFilialOP And objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) Then gError 83158
                    End If
                Next
                
            '#####################################################
            'Inserido por Wagner 10/03/2006
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) And objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) Then gError 83157
                    End If
                Next
            
                'Carrega as séries na coleção global
                lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(objGridInt.objGrid.Row), StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)), objRastroLote.sCodigo, StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)), sLoteAnterior, objGridInt.objGrid.Row)
                If lErro <> SUCESSO Then gError 141861
            '#####################################################
            
            End If


        End If

        'Coloca o Lote na tela
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = objRastroLote.sCodigo
        LoteRastro.Text = objRastroLote.sCodigo
        
        'Lê lote e preenche dados
        lErro = Lote_Saida_Celula(objRastroLote)
        If lErro <> SUCESSO Then gError 83159
        
    End If

    objTela.Show

    Exit Sub

Erro_objEventoLoteRastro_evSelecao:

    Select Case gErr

        Case 83154, 83155, 83159, 141861
        
        Case 83156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 83157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo)
            
        Case 83158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo, objRastroLote.iFilialOP)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152770)

    End Select

    Exit Sub

End Sub

Public Sub ItemNFRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub ItemNFRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub ItemNFRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = ItemNFRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdutoRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub ProdutoRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub ProdutoRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = ProdutoRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AlmoxRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AlmoxRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub AlmoxRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = AlmoxRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub EscaninhoRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EscaninhoRastro_Click()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EscaninhoRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub EscaninhoRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = EscaninhoRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub LoteRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LoteRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub LoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub LoteRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = LoteRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialOPRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOPRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub FilialOPRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = FilialOPRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantLoteRastro_Change()
'Rastreamento

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantLoteRastro_GotFocus()
'Rastreamento

    Call Grid_Campo_Recebe_Foco(objGridInt)

End Sub

Public Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridInt)

End Sub

Public Sub QuantLoteRastro_Validate(Cancel As Boolean)
'Rastreamento

Dim lErro As Long

    Set objGridInt.objControle = QuantLoteRastro
    lErro = Grid_Campo_Libera_Foco(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridRastro_Click()
'Rastreamento

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridInt, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)
    End If

End Sub

Public Sub GridRastro_EnterCell()
'Rastreamento

    Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)

End Sub

Public Sub GridRastro_GotFocus()
'Rastreamento

    Call Grid_Recebe_Foco(objGridInt)

End Sub

Public Sub GridRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridInt, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)
    End If

End Sub

Public Sub GridRastro_LeaveCell()
'Rastreamento

    Call objTela.Saida_Celula(objGridInt)

End Sub

Public Sub GridRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call Grid_Libera_Foco(objGridInt)

End Sub

Public Sub GridRastro_RowColChange()
'Rastreamento

    Call Grid_RowColChange(objGridInt)

End Sub

Public Sub GridRastro_Scroll()
'Rastreamento

    Call Grid_Scroll(objGridInt)

End Sub

Public Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
'Rastreamento

Dim lErro As Long

'########################################
'Inserido por Wagner 10/03/2006
Dim iLinhasExistentesAnterior As Integer
Dim iItemAnterior As Integer
'########################################

On Error GoTo Erro_GridRastro_KeyDown

    '#######################################################
    'Inserido por Wagner 10/03/2006
    iLinhasExistentesAnterior = objGridInt.iLinhasExistentes
    iItemAnterior = GridRastro.Row
    '#######################################################

    Call Grid_Trata_Tecla1(KeyCode, objGridInt)

    '#######################################################
    'Inserido por Wagner 10/03/2006
    If objGridInt.iLinhasExistentes < iLinhasExistentesAnterior Then
        gcolcolRastreamentoSerie.Remove iItemAnterior
    End If
    '#######################################################

    Exit Sub

Erro_GridRastro_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152771)

    End Select
    
    Exit Sub

End Sub

'Public Sub Rastreamento_Valida_Tab()
''Rastreamento
'
'Dim iLinhaRastro As Integer
'Dim iLinha As Integer
'Dim iAchou As Integer
'
'    For iLinhaRastro = objGridInt.iLinhasExistentes To 1 Step -1
'        iAchou = 0
'        For iLinha = 1 To objGridInt.iLinhasExistentes
'            If objGridRastro.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col) = objGridItens.objGrid.TextMatrix(iLinha, iGrid_Almoxarifado_Col) And _
'               objGridRastro.objGrid.TextMatrix(iLinhaRastro, iGrid_ProdutoRastro_Col) = objGridItens.objGrid.TextMatrix(iLinha, iGrid_Produto_Col) Then
'                objGridRastro.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantRastro_Col) = objGridItens.objGrid.TextMatrix(iLinha, iGrid_Quantidade_Col)
'                iAchou = 1
'                Exit For
'            End If
'        Next
'        If iAchou = 0 Then
'            objGridRastro.objGrid.Row = iLinhaRastro
'            Call GridRastro_KeyDown(vbKeyDelete, 0)
'        End If
'    Next
'
'End Sub

Public Function Exclusao_Item_Grid(iItemAtual As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim sProduto1 As String
Dim sProduto2 As String
Dim iProdutoPreenchido As Integer
Dim bExcluiu As Boolean

On Error GoTo Erro_Exclusao_Item_Grid

    bExcluiu = False

    lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iItemAtual, objTela.iGrid_Produto_Col), sProduto1, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 180070
    
    For iLinha = objGridInt.iLinhasExistentes To 1 Step -1
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
        
            lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col), sProduto2, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 180071
    
            If sProduto1 = sProduto2 Then
                'exclui a linha em questão
                Call Grid_Exclui_Linha(objGridInt, iLinha)
                bExcluiu = True
            End If
        End If
    Next
    
    If bExcluiu Then
        For iLinha = 1 To objGridInt.iLinhasExistentes
            If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
                If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) > iItemAtual Then
                    objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) - 1
                End If
            End If
        Next
    End If
    
    Exclusao_Item_Grid = SUCESSO
    
    Exit Function
    
Erro_Exclusao_Item_Grid:

    Exclusao_Item_Grid = gErr

    Select Case gErr
    
        Case 180070, 180071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152772)

    End Select
    
    Exit Function
    
End Function

Public Function Atualiza_Troca_Almoxarifado(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_Troca_Almoxarifado
    
    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
            If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) = iItemAtual Then
                objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItemAtual, objTela.iGrid_Almoxarifado_Col)
            End If
        End If
    Next

    Atualiza_Troca_Almoxarifado = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_Almoxarifado:

    Atualiza_Troca_Almoxarifado = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152773)

    End Select
    
    Exit Function

End Function

Public Function Atualiza_Troca_Almoxarifado1(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_Troca_Almoxarifado1

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
            If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) = iItemAtual Then
                objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iItemAtual, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col)
            End If
        End If
    Next

    Atualiza_Troca_Almoxarifado1 = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_Almoxarifado1:

    Atualiza_Troca_Almoxarifado1 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152774)

    End Select
    
    Exit Function

End Function

Public Function Atualiza_Troca_Quantidade(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_Troca_Quantidade

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
            If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) = iItemAtual Then
                objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItemAtual, objTela.iGrid_Quantidade_Col)
            End If
        End If
    Next

    Atualiza_Troca_Quantidade = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_Quantidade:

    Atualiza_Troca_Quantidade = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152775)

    End Select
    
    Exit Function

End Function

Public Function Atualiza_Troca_UM(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_Troca_UM

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
            If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) = iItemAtual Then
                objGridInt.objGrid.TextMatrix(iLinha, iGrid_UMRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItemAtual, objTela.iGrid_UnidadeMed_Col)
            End If
        End If
    Next

    Atualiza_Troca_UM = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_UM:

    Atualiza_Troca_UM = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152776)

    End Select
    
    Exit Function

End Function

Public Function Valida_Rastreamento(objTipoDocInfo As ClassTipoDocInfo) As Long

Dim iLinhaItens As Integer
Dim iLinhaAlocacao As Integer
Dim iLinhaRastro1 As Integer
Dim iLinhaRastro As Integer
Dim dQuantidade As Double
Dim sProduto As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim bItemPossuiRastro As Boolean
Dim objProdutoItem As New ClassProduto
Dim sProdutoFormatadoItem As String
Dim iProdutoPreenchidoItem As Integer
Dim dFatorConversao As Double

On Error GoTo Erro_Valida_Rastreamento

    'se o documento não pode ter rastreamento associado e na tela foram informados rastreamento ==> erro
    If objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_NAO And objGridInt.iLinhasExistentes > 0 Then gError 83448

    For iLinhaRastro = 1 To objGridInt.iLinhasExistentes
    
        'se o item da nota fiscal que está associado ao rastreamento não estiver preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ItemNFRastro_Col))) = 0 Then gError 83208

        'se o lote não estiver preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_LoteRastro_Col))) = 0 Then gError 83209

        'se o produto não estiver preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ProdutoRastro_Col))) = 0 Then gError 83210

        'se o almoxarifado não estiver preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col))) = 0 Then gError 83216

        If bEscaninhoVisivel = True Then

            'se o escaninho não estiver preenchido ==> erro
            If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_EscaninhoRastro_Col))) = 0 Then gError 83320

        End If

        'se a unidade de medida não estiver preenchida ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_UMRastro_Col))) = 0 Then gError 83217

        sProduto = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ProdutoRastro_Col)

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83211
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83212
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83213

        'Se o produto possuir rastro por op e a filial op não estiver preenchida ==> erro
        If objProduto.iRastro = PRODUTO_RASTRO_OP And Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_FilialOPRastro_Col))) = 0 Then gError 83219

        'Se a quantidade não foi preenchida,erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col))) = 0 Then gError 83214
        
        'Se a quantidade está zerada, erro
        If StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col)) = 0 Then gError 83215
        
    Next

    If bEscaninhoVisivel = False And iCodigo <> DOCINFO_NFISFBF And iCodigo <> DOCINFO_NFISPC And iCodigo <> DOCINFO_NFFISPC And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalRemPedido" Then
    
        'mario distribuicao se forem as notas fiscais de entrada ==> fazer este tratamento pois as notas de entrada ganharam o grid de distribuicao que não havia antes do compras
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
        
            For iLinhaAlocacao = 1 To objTela.gobjDistribuicao.objGridDist.iLinhasExistentes
                
                '##########################################################################
                'Inserido por Wagner 02/06/2006
                'Formata o Produto para o BD
                Set objProdutoItem = New ClassProduto
                
                lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(CInt(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_ItemNFDist_Col)), objTela.iGrid_Produto_Col), sProdutoFormatadoItem, iProdutoPreenchidoItem)
                If lErro <> SUCESSO Then gError 180052
                    
                objProdutoItem.sCodigo = sProdutoFormatadoItem
                        
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProdutoItem)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 180053
                '##########################################################################
                
                'Só valida os que não forem Kit de Venda
                If objProdutoItem.iKitVendaComp = DESMARCADO And Len(Trim(objProdutoItem.sGrade)) = 0 Then
                
                    dQuantidade = 0
                    For iLinhaRastro = 1 To objGridInt.iLinhasExistentes
                        
                        'soma as quantidade dos ratros para o mesmo par item/almoxarifado
                        If CInt(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_ItemNFDist_Col)) = CInt(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ItemNFRastro_Col)) And _
                           objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col) = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col) Then
                            dQuantidade = dQuantidade + StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col))
                        End If
                    Next
                                       
                    If (dQuantidade - StrParaDbl(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_QuantDist_Col))) > 0.0001 Then gError 89498
            
                End If
            
            Next
        
        Else
                            
            For iLinhaItens = 1 To objTela.objGridItens.iLinhasExistentes
                
                '##########################################################################
                'Inserido por Wagner 02/06/2006
                'Formata o Produto para o BD
                Set objProdutoItem = New ClassProduto
                
                lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iLinhaItens, objTela.iGrid_Produto_Col), sProdutoFormatadoItem, iProdutoPreenchidoItem)
                If lErro <> SUCESSO Then gError 180054
                    
                objProdutoItem.sCodigo = sProdutoFormatadoItem
                        
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProdutoItem)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 180055
                '##########################################################################
                
                'Só valida os que não forem Kit de Venda
                If objProdutoItem.iKitVendaComp = DESMARCADO And Len(Trim(objProdutoItem.sGrade)) = 0 Then
                
                    dQuantidade = 0
                    For iLinhaRastro = 1 To objGridInt.iLinhasExistentes
                        'soma as quantidade dos ratros para o mesmo item
                        If CInt(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ItemNFRastro_Col)) = iLinhaItens Then
                            
                            lErro = CF("UM_Conversao_Trans", objProdutoItem.iClasseUM, objTela.objGridItens.objGrid.TextMatrix(iLinhaItens, objTela.iGrid_UnidadeMed_Col), objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_UMRastro_Col), dFatorConversao)
                            If lErro <> SUCESSO Then gError 180053
                            
                            dQuantidade = dQuantidade + StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col)) / dFatorConversao
                        End If
                    Next
                    

                    If (dQuantidade - StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iLinhaItens, objTela.iGrid_Quantidade_Col))) > 0.0001 Then gError 83218
                
                End If
                
            Next
            
        End If
    Else
        For iLinhaAlocacao = 1 To objTela.objGridAlocacoes.iLinhasExistentes
            dQuantidade = 0
            For iLinhaRastro = 1 To objGridInt.iLinhasExistentes
                
                'soma as quantidade dos ratros para o mesmo par item/almoxarifado
                If CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinhaAlocacao, objTela.iGrid_Item_Col)) = CInt(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ItemNFRastro_Col)) And _
                   objTela.objGridAlocacoes.objGrid.TextMatrix(iLinhaAlocacao, objTela.iGrid_AlmoxAloc_Col) = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col) Then
                    dQuantidade = dQuantidade + StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col))
                End If
            Next
            
            If (dQuantidade > StrParaDbl(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinhaAlocacao, objTela.iGrid_QuantAloc_Col))) > QTDE_ESTOQUE_DELTA Then gError 83321
            
        Next
    End If

    Valida_Rastreamento = SUCESSO
    
    Exit Function
    
Erro_Valida_Rastreamento:

    Valida_Rastreamento = gErr

    Select Case gErr

        Case 83208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_ITEM_NAO_PREENCHIDO", gErr, iLinhaRastro)

        Case 83209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_LOTE_NAO_PREENCHIDO", gErr, iLinhaRastro)
        
        Case 83210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_PRODUTO_NAO_PREENCHIDO", gErr, iLinhaRastro)
        
        Case 83211, 83212, 180052, 180053, 180054, 180055
        
        Case 83213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo, iLinhaRastro)
        
        Case 83214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_QUANT_NAO_PREENCHIDA", gErr, iLinhaRastro)
        
        Case 83215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_QUANT_ZERADA", gErr, iLinhaRastro)
        
        Case 83216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_ALMOX_NAO_PREENCHIDO", gErr, iLinhaRastro)
        
        Case 83217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_UM_NAO_PREENCHIDO", gErr, iLinhaRastro)
        
        Case 83218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_QUANT_MAIOR_ITEM", gErr, iLinhaItens, dQuantidade, StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iLinhaItens, objTela.iGrid_Quantidade_Col)))
        
        Case 83219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_FILIALOP_NAO_PREENCHIDA", gErr, iLinhaRastro)
            
        Case 83320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_ESCANINHO_NAO_PREENCHIDO", gErr, iLinhaRastro)
            
        Case 83321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_QUANT_MAIOR_ALOC", gErr, objProduto.sCodigo, objTela.objGridAlocacoes.objGrid.TextMatrix(iLinhaAlocacao, objTela.iGrid_AlmoxAloc_Col), dQuantidade, StrParaDbl(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinhaAlocacao, objTela.iGrid_QuantAloc_Col)))
            
        Case 83448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_RASTREAVEL", gErr)
            
        Case 89498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRIDRASTRO_QUANT_MAIOR_ALOC", gErr, objProduto.sCodigo, objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col), dQuantidade, StrParaDbl(objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinhaAlocacao, objTela.gobjDistribuicao.iGrid_QuantDist_Col)))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152777)

    End Select
    
    Exit Function

End Function

Public Function Move_Rastro_Memoria(objNFiscal As ClassNFiscal) As Long

Dim objItemNF As ClassItemNF
Dim iLinhaRastro As Integer
Dim objRastroItemNF As ClassRastroItemNF
Dim lErro As Long
Dim objItemAloc As ClassItemNFAlocacao
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Move_Rastro_Memoria
        
    If iCodigo <> 0 Then
    
        objTipoDocInfo.iCodigo = iCodigo
    
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 89522
    
        'Se não achou o Tipo de Documento --> erro
        If lErro = 31415 Then gError 89523
    
    Else
        objTipoDocInfo.iTipo = -1
    End If
            
    For Each objItemNF In objNFiscal.colItensNF
    
        Set objItemNF.colRastreamento = New Collection
    
        For iLinhaRastro = 1 To objGridInt.iLinhasExistentes
        
            If StrParaInt(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ItemNFRastro_Col)) = objItemNF.iItem Then
        
                'Cria novo item de rastreamento
                Set objRastroItemNF = New ClassRastroItemNF
                
                If (bEscaninhoVisivel = True And bSaidaMaterial = True) Or iCodigo = DOCINFO_NFISFBF Or iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRemPedido" Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
                    For Each objItemAloc In objItemNF.ColAlocacoes
                        If objItemAloc.sAlmoxarifado = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col) Then
                            objRastroItemNF.iAlmoxCodigo = objItemAloc.iAlmoxarifado
                            Exit For
                        End If
                    Next
                Else
                    objRastroItemNF.iAlmoxCodigo = objItemNF.iAlmoxarifado
                End If
                
                objRastroItemNF.dAlmoxQtdAlocada = StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantRastro_Col))
                objRastroItemNF.dLoteQdtAlocada = StrParaDbl(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_QuantLoteRastro_Col))
                objRastroItemNF.dtLoteData = StrParaDate(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_LoteDataRastro_Col))
                objRastroItemNF.sAlmoxNomeRed = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_AlmoxRastro_Col)
                objRastroItemNF.sLote = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_LoteRastro_Col)
                objRastroItemNF.iLoteFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_FilialOPRastro_Col))
                objRastroItemNF.sSiglaUM = objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_UMRastro_Col)
                
                '#################################################################
                'Inserido por Wagner 02/06/2006
                If Len(Trim(objRastroItemNF.sAlmoxNomeRed)) > 0 Then
                    Set objAlmoxarifado = New ClassAlmoxarifado
                    
                    objAlmoxarifado.sNomeReduzido = objRastroItemNF.sAlmoxNomeRed
            
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError 180045
                    
                    objRastroItemNF.iAlmoxCodigo = objAlmoxarifado.iCodigo
                End If
                
                'Formata o produto
                lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 180023
                
                objRastroItemNF.sProduto = sProdutoFormatado
                '#################################################################
                
                objRastroItemNF.iItem = iLinhaRastro
                
                If bEscaninhoVisivel = True Then
                    If objGridInt.objGrid.TextMatrix(iLinhaRastro, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_CONSIGNADO Then
                        objRastroItemNF.iEscaninho = ESCANINHO_CONSIG_3
                    Else
                        objRastroItemNF.iEscaninho = ESCANINHO_DISPONIVEL
                    End If
                End If
                            
                'Adiciona na coleção
                objItemNF.colRastreamento.Add objRastroItemNF
        
            End If
        
        Next
        
    Next
    
    Move_Rastro_Memoria = SUCESSO
       
    Exit Function
    
Erro_Move_Rastro_Memoria:

    Move_Rastro_Memoria = gErr
    
    Select Case gErr
    
        Case 89522, 180023, 180045
        
        Case 89523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152778)

    End Select
    
    Exit Function
    
End Function

Public Sub Limpa_Tela_Rastreamento()
'limpa a parte de Rastreamento

    'Limpa o Grid de Rastreamento
    Call Grid_Limpa(objGridInt)
    Set gcolcolRastreamentoSerie = New Collection
    
End Sub

Public Sub objUserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If objTela.ActiveControl Is LoteRastro Then
            Call BotaoLotes_Click
    End If

End Sub

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objItemNFTela As ClassItemNF
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim colMovEstoque As Collection
Dim colRastreamentoMovto As Collection
Dim objNFiscalGravada As New ClassNFiscal
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer 'Inserido por Wagner 10/03/2006
Dim dQuantidade As Double
Dim objProduto As ClassProduto
Dim dQtdeAux As Double

On Error GoTo Erro_Move_Tela_Memoria

    objNFiscalGravada.lNumIntDoc = objNFiscal.lNumIntDoc

    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFiscalGravada)
    If lErro <> SUCESSO Then gError 75812

    'se o numero de itens da nota gravada for diferente do numero de itens da tela ==> erro
    If objNFiscalGravada.colItensNF.Count <> objNFiscal.colItensNF.Count Then gError 83259

    'Para cada item ==> le os movimentos e coloca os rastreamentos associados aos movimentos
    For iIndice = 1 To objNFiscalGravada.colItensNF.Count

        Set objItemNF = objNFiscalGravada.colItensNF.Item(iIndice)
        Set objItemNFTela = objNFiscal.colItensNF.Item(iIndice)

        'se o produto gravado for diferente do produto na tela ==> erro
        If objItemNF.sProduto <> objItemNFTela.sProduto Then gError 83260

'        If bEscaninhoVisivel = False And iCodigo <> DOCINFO_NFISPC And iCodigo <> DOCINFO_NFFISPC Then
'            'se o almoxarifado gravado for diferente do almoxarifado na tela ==> erro
'            If objItemNF.iAlmoxarifado <> objItemNFTela.iAlmoxarifado Then gError 83261
'        End If

        'se a unidade de medida gravada for diferente da UM na tela ==> erro
        If objItemNF.sUnidadeMed <> objItemNFTela.sUnidadeMed Then gError 83262

        'se a quantidade gravada for diferente da quantidade na tela ==> erro
        If objItemNF.dQuantidade <> objItemNFTela.dQuantidade Then gError 83263

        Set objItemMovEstoque = New ClassItemMovEstoque
    
        'Lê Numintdoc e código do movimento estoque a partir de um item da nota fiscal
        objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
        objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
        objItemMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        
        Set colMovEstoque = New Collection
        
        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
        If lErro <> SUCESSO Then gError 83236
        
        '##################################################
        'Inserido por Wagner 02/06/2006
        lErro = CF("MovEstoque_Le_ItemNF_Grade", objItemMovEstoque, colMovEstoque)
        If lErro <> SUCESSO Then gError 180056
        '##################################################
        
        'primeiro vai colocar os movimentos na sua ordem inversa de cadastramento para que os rastreamentos possam ser excluidos nesta ordem, depois o rastreamento é colocado normalmente.
        'Desta forma não haverá possibilidade de ocorrer o aviso de falta de estoque indevidamente
        'como este item não está com o rastreamento acoplado ==> o seu rastreamento será excluido pela rotina de alteração e como não tem rastreamento acoplado, não será inserido por esta coleção o rastreamento novo.
        For iIndice1 = colMovEstoque.Count To 1 Step -1
            objMovEstoque.colItens.Add1 colMovEstoque.Item(iIndice1)
        Next
        
        Set colMovEstoque = New Collection
        
        'esta nova leitura dos movimentos de estoque, permitirá que sejam carregados os novos rastreamentos na ordem em que foram originalmente processados
        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
        If lErro <> SUCESSO Then gError 83371
        
        '##################################################
        'Inserido por Wagner 02/06/2006
        lErro = CF("MovEstoque_Le_ItemNF_Grade", objItemMovEstoque, colMovEstoque)
        If lErro <> SUCESSO Then gError 180057
        '##################################################
        
        For iIndice1 = 1 To colMovEstoque.Count
        
            Set objItemMovEstoque = colMovEstoque.Item(iIndice1)
        
'            If (iIndice1 = 1 And bEscaninhoVisivel = False) Or bEscaninhoVisivel = True Then
        
                'Guarda código e data do Movimento de estoque
                objMovEstoque.lCodigo = objItemMovEstoque.lCodigo
                objMovEstoque.dtData = objItemMovEstoque.dtData
                objMovEstoque.dtHora = objItemMovEstoque.dtHora
                objMovEstoque.iFilialEmpresa = objItemMovEstoque.iFilialEmpresa
                objMovEstoque.iTipoMov = objItemMovEstoque.iTipoMov
                objMovEstoque.lCliente = objItemMovEstoque.lCliente
                objMovEstoque.lFornecedor = objItemMovEstoque.lFornecedor
                objMovEstoque.sDocOrigem = objItemMovEstoque.sDocOrigem
                        
                Set objItemMovEstoque.colRastreamentoMovto = New Collection
            
                dQuantidade = 0
                           
                'Guarda o Rastreamento dos ItensNF
                For Each objRastroItemNF In objItemNFTela.colRastreamento
                
                    'se o almoxarifado do rastreamento for diferente do movimento de estoque ==> pega o proximo rastreamento
                    If objRastroItemNF.iAlmoxCodigo <> objItemMovEstoque.iAlmoxarifado Then GoTo Label_Proximo_Rastreamento
                    
                    'se o movimento de estoque for de material consignado e o rastreamento for diferente de consignação ==> pega o proximo rastreamento
                    If (objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG Or objItemMovEstoque.iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_TERC Or objItemMovEstoque.iTipoMov = MOV_EST_ENTRADA_TRANSF_DISP1) And objRastroItemNF.iEscaninho <> ESCANINHO_CONSIG_3 Then GoTo Label_Proximo_Rastreamento
                
                    'se o movimento de estoque for de material disponivel e o rastreamento for diferente de material disponivel ==> pega o proximo rastreamento
                    If objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA And objRastroItemNF.iEscaninho <> ESCANINHO_DISPONIVEL Then GoTo Label_Proximo_Rastreamento
                    
                    Set objRastreamentoMovto = New ClassRastreamentoMovto
                    
                    objRastreamentoMovto.sProduto = objRastroItemNF.sProduto
                    objRastreamentoMovto.sLote = objRastroItemNF.sLote
                    objRastreamentoMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
                    objRastreamentoMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
                    objRastreamentoMovto.sSiglaUM = objRastroItemNF.sSiglaUM
                    
                    'se o produto do rastreamento na tela for diferente do item da nota fiscal gravado ==> erro
                    If UCase(objRastroItemNF.sProduto) <> UCase(objItemMovEstoque.sProduto) And objItemNF.colItensRomaneioGrade.Count = 0 Then gError 83254

                    'se o almoxarifado do rastreamento na tela for diferente do item da nota fiscal gravado ==> erro
                    If objRastroItemNF.iAlmoxCodigo <> objItemMovEstoque.iAlmoxarifado And objItemNF.colItensRomaneioGrade.Count = 0 Then gError 83255
                                        
                    'se a unidade de medida do rastreamento na tela for diferente do item da nota fiscal gravado ==> erro
                    If UCase(objRastroItemNF.sSiglaUM) <> UCase(objItemMovEstoque.sSiglaUM) And objItemNF.colItensRomaneioGrade.Count = 0 Then gError 83256
                                        
                    If objRastroItemNF.dAlmoxQtdAlocada <> objItemMovEstoque.dQuantidade And objItemNF.colItensRomaneioGrade.Count = 0 Then gError 83257
                    
                    dQuantidade = dQuantidade + objRastroItemNF.dLoteQdtAlocada
                    
                    If UCase(objRastroItemNF.sProduto) = UCase(objItemMovEstoque.sProduto) And objRastroItemNF.iAlmoxCodigo = objItemMovEstoque.iAlmoxarifado And objRastroItemNF.dAlmoxQtdAlocada = objItemMovEstoque.dQuantidade And UCase(objRastroItemNF.sSiglaUM) = UCase(objItemMovEstoque.sSiglaUM) Then
                    
                        '##########################################################
                        'Alterado por Wagner 10/03/2006
                        Set objProduto = New ClassProduto
                        
                        objProduto.sCodigo = objItemMovEstoque.sProduto
                        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 83459
    
                        If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                        
'                            'Acha as séries na coleção global
'                            For iIndice2 = 1 To objGridInt.iLinhasExistentes
'                                If StrParaInt(objGridInt.objGrid.TextMatrix(iIndice2, iGrid_ItemNFRastro_Col)) = objRastroItemNF.iItem And objRastroItemNF.sLote = objGridInt.objGrid.TextMatrix(iIndice2, iGrid_LoteRastro_Col) Then
'                                    Set objItemMovEstoque.colRastreamentoMovto = gcolcolRastreamentoSerie.Item(iIndice)
'                                    Exit For
'                                End If
'                            Next

                            'Set objItemMovEstoque.colRastreamentoMovto = gcolcolRastreamentoSerie.Item(objRastroItemNF.iItem)
                            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(objRastroItemNF.iItem)
                                objRastreamentoMovto.dQuantidadeSerie = 0
                                objItemMovEstoque.colRastreamentoMovto.Add objRastreamentoMovto
                            Next
                            dQtdeAux = 0
                            For Each objRastreamentoMovto In objItemMovEstoque.colRastreamentoMovto
                                dQtdeAux = dQtdeAux + objRastreamentoMovto.dQuantidade
                            Next
                            For Each objRastreamentoMovto In objItemMovEstoque.colRastreamentoMovto
                                'objRastreamentoMovto.dQuantidadeSerie = objRastroItemNF.dLoteQdtAlocada
                                objRastreamentoMovto.dQuantidadeSerie = dQtdeAux
                            Next
                        Else
                            'Adiciona objRastreamentoMovto na coleção de Rastreamento
                            objItemMovEstoque.colRastreamentoMovto.Add objRastreamentoMovto
                        
                        End If
                        '################################################################
                        
                        objRastroItemNF.iUtilizado = 1
                    
                    End If
                    
Label_Proximo_Rastreamento:
                Next
            
                'se a soma dos rastreamentos ultrapassar a quantidade do item da nf
                If (dQuantidade - objItemMovEstoque.dQuantidade) > 0.0001 And objItemNF.colItensRomaneioGrade.Count = 0 Then gError 83258
                
'                'se a soma dos rastreamentos for menor do que a quantidade do item da nf
'                If dQuantidade < objItemMovEstoque.dQuantidade Then gError 83457
                
            'End If
            
            objMovEstoque.colItens.Add1 objItemMovEstoque
                    
        Next
        
        'Verifica se algum rastreamento deixou de ser utilizado. Se deixou ==> erro
        For Each objRastroItemNF In objItemNFTela.colRastreamento
            
            If objRastroItemNF.iUtilizado <> 1 Then gError 83322
        
        Next
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 83236, 83237, 83371, 83459, 180056, 180057
        
        Case 83254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_RASTRO_DIF", gErr, objRastroItemNF.sProduto, objRastroItemNF.iItem, objItemMovEstoque.sProduto)
        
        Case 83255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOX_RASTRO_DIF", gErr, objRastroItemNF.iAlmoxCodigo, objRastroItemNF.iItem, objItemMovEstoque.iAlmoxarifado)
        
        Case 83256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAUM_RASTRO_DIF_ITEMNF", gErr, objRastroItemNF.sSiglaUM, objRastroItemNF.iItem, objItemMovEstoque.sSiglaUM)
        
        Case 83257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTRO_DIF_ITEMNF", gErr, objRastroItemNF.dAlmoxQtdAlocada, objRastroItemNF.iItem, objItemNF.dQuantidade)
        
        Case 83258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTRO_MENOR_ALOC", gErr, objItemMovEstoque.sProduto, objItemMovEstoque.iAlmoxarifado, objItemMovEstoque.dQuantidade, dQuantidade)
        
        Case 83259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_ITEMNF_DIF_TELA", gErr, objNFiscalGravada.colItensNF.Count, objNFiscal.colItensNF.Count)
        
        Case 83260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_PRODUTO_DIF_TELA", gErr, objItemNFTela.sProduto, objItemNFTela.iItem, objItemNF.sProduto)
            
        Case 83261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_ALMOX_DIF_TELA", gErr, objItemNFTela.iAlmoxarifado, objItemNFTela.iItem, objItemNF.iAlmoxarifado)
            
        Case 83262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_UM_DIF_TELA", gErr, objItemNFTela.sUnidadeMed, objItemNFTela.iItem, objItemNF.sUnidadeMed)
            
        Case 83263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_QUANT_DIF_TELA", gErr, objItemNFTela.dQuantidade, objItemNFTela.iItem, objItemNF.dQuantidade)
            
        Case 83322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTRO_NAO_UTILIZADO", gErr, objRastroItemNF.iItem)
            
        Case 83457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTRO_MENOR_ALOC", gErr, objItemMovEstoque.sProduto, objItemMovEstoque.iAlmoxarifado, objItemMovEstoque.dQuantidade, dQuantidade)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152779)

    End Select

    Exit Function

End Function

Public Function Altera_Rastreamento(objNFiscal As ClassNFiscal) As Long
'Altera o rastreamento da Nota Fiscal no BD

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Altera_Rastreamento

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Le o tipo do produto
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 89984

    lErro = Move_Tela_Memoria(objNFiscal, objMovEstoque)
    If lErro <> SUCESSO Then gError 83239
    
    For Each objItemMovEstoque In objMovEstoque.colItens
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        
            '##############################################
            'Alterado por Wagner
            If objNFiscal.dtDataSaida = DATA_NULA Then
                objItemMovEstoque.dtData = objNFiscal.dtDataEmissao
            Else
                objItemMovEstoque.dtData = objNFiscal.dtDataSaida
            End If
            objItemMovEstoque.dtHora = objNFiscal.dtHoraSaida
            '##############################################
        Else
            objItemMovEstoque.dtData = objNFiscal.dtDataEntrada
            objItemMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        End If
    Next
    
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        '##############################################
        'Alterado por Wagner
        If objNFiscal.dtDataSaida = DATA_NULA Then
            objMovEstoque.dtData = objNFiscal.dtDataEmissao
        Else
            objMovEstoque.dtData = objNFiscal.dtDataSaida
        End If
        objMovEstoque.dtHora = objNFiscal.dtHoraSaida
        '##############################################
    Else
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
    End If
    
    'Altera o Rastreamento no BD
    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then gError 83240

    Altera_Rastreamento = SUCESSO
    
    Exit Function

Erro_Altera_Rastreamento:

    Altera_Rastreamento = gErr
    
    Select Case gErr
        
        Case 83239, 83240
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152780)

    End Select

    Exit Function

End Function

Public Function Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque As ClassItemMovEstoque, colRastreamento As Collection, objTipoDocInfo As ClassTipoDocInfo)
'Guarda o Rastreamento dos ItensNF

Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastroItemNF As ClassRastroItemNF
Dim lErro As Long
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim iIndice As Integer 'Inserido por Wagner 10/03/2006
Dim dQtdeAux As Double

On Error GoTo Erro_Move_RastroItemNF_ItemMovEstoque
                
    If objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_SIM And Not (colRastreamento Is Nothing) Then
                        
        objProduto.sCodigo = objItemMovEstoque.sProduto
                    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83459
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83460
                
        'se não for um produto rastreavel ==> erro
        If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
                    
            'Guarda o Rastreamento dos ItensNF
            For Each objRastroItemNF In colRastreamento
                
                If objRastroItemNF.iAlmoxCodigo = objItemMovEstoque.iAlmoxarifado And objRastroItemNF.sProduto = objItemMovEstoque.sProduto Then
                                   
                    '##########################################################
                    'Alterado por Wagner 10/03/2006
                    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                          'Acha as séries na coleção global
                        For iIndice = 1 To objGridInt.iLinhasExistentes
'                            If StrParaInt(objGridInt.objGrid.TextMatrix(iIndice, iGrid_ItemNFRastro_Col)) = objRastroItemNF.iItem And objRastroItemNF.sLote = objGridInt.objGrid.TextMatrix(iIndice, iGrid_LoteRastro_Col) Then
                             If objRastroItemNF.sLote = objGridInt.objGrid.TextMatrix(iIndice, iGrid_LoteRastro_Col) Then
                                'Set objItemMovEstoque.colRastreamentoMovto = gcolcolRastreamentoSerie.Item(iIndice)
                                For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iIndice)
                                    objItemMovEstoque.colRastreamentoMovto.Add objRastreamentoMovto
                                Next
                            End If
                        Next
                        dQtdeAux = 0
                        For Each objRastreamentoMovto In objItemMovEstoque.colRastreamentoMovto
                            objRastreamentoMovto.dQuantidadeSerie = 0
                            dQtdeAux = dQtdeAux + objRastreamentoMovto.dQuantidade
                        Next
                        For Each objRastreamentoMovto In objItemMovEstoque.colRastreamentoMovto
                            'objRastreamentoMovto.dQuantidadeSerie = objRastroItemNF.dLoteQdtAlocada
                            objRastreamentoMovto.dQuantidadeSerie = dQtdeAux
                        Next
                        
                    Else
                    
                        Set objRastreamentoMovto = New ClassRastreamentoMovto
                        
                        objRastreamentoMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
                        objRastreamentoMovto.sLote = objRastroItemNF.sLote
                        objRastreamentoMovto.sProduto = objRastroItemNF.sProduto
                        objRastreamentoMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
                        objRastreamentoMovto.sSiglaUM = objRastroItemNF.sSiglaUM
                    
                        'Adiciona objRastreamentoMovto na coleção de Rastreamento
                        objItemMovEstoque.colRastreamentoMovto.Add objRastreamentoMovto
                    
                    End If
                    '##########################################################
                    
                    'soma as quantidades do rastreamento para verificar se ela é igual ao do movimento de estoque
                    dQuantidade = dQuantidade + objRastroItemNF.dLoteQdtAlocada
                
                End If
                
            Next
        
            'se a quantidade do rastreamento for diferente da quantidade do movimento de estoque ==> erro
            If (dQuantidade - objItemMovEstoque.dQuantidade) > 0.0001 Then gError 83458
    
        End If

    End If

    Move_RastroItemNF_ItemMovEstoque = SUCESSO
    
    Exit Function

Erro_Move_RastroItemNF_ItemMovEstoque:

    Move_RastroItemNF_ItemMovEstoque = gErr

    Select Case gErr
        
        Case 83458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_RASTRO_DIF_ALOCADO", gErr, objItemMovEstoque.sProduto, objItemMovEstoque.sAlmoxarifadoNomeRed, objItemMovEstoque.dQuantidade, dQuantidade)
        
        Case 83459
        
        Case 83460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemMovEstoque.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152781)

    End Select

    Exit Function

End Function

'####################################################################
'Inserido por Wagner 14/03/2006
Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_BotaoSerie_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If objGridInt.objGrid.Row = 0 Then gError 141812
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col))) = 0 Then gError 141867
    If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col))) = 0 Then gError 177303
    If StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantLoteRastro_Col)) = 0 Then gError 177304
        
    'Formata o produto
    lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141813
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141814
       
    objItemMovEstoque.dQuantidade = StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col))
    objItemMovEstoque.iItemNF = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col)
    objItemMovEstoque.sAlmoxarifadoNomeRed = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col)
        
    objTipoDocInfo.iCodigo = iCodigo

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 177243
    
    objItemMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela_Modal("RastreamentoSerie", gcolcolRastreamentoSerie.Item(objGridInt.objGrid.Row), objItemMovEstoque, objTela.Name)
                    
    lErro = Acerta_Quantidade_Rastreada(objGridInt.objGrid.Row)
    If lErro <> SUCESSO Then gError 141860
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case 141812
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141813, 141814, 141860, 141867
        
        Case 177303
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, objGridInt.objGrid.Row)
        
        Case 177304
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, objGridInt.objGrid.Row)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152769)
    
    End Select
    
    Exit Sub

End Sub

Public Function Carrega_Series(colRastreamentoMovto As Collection, ByVal dQuantidade As Double, ByVal sLoteIni As String, ByVal dQuantidadeAnterior As Double, ByVal sLoteIniAnterior As String, ByVal iLinha As Integer)
'Gera o rastreamento por número de série

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoSerie As ClassRastreamentoLote
Dim objRastreamentoSerieIni As ClassRastreamentoLote
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objAlmoxarifado As ClassAlmoxarifado
Dim vbResult As VbMsgBoxResult
Dim colRastreamentoMovtoAux As New Collection
Dim objTipoDocInfo As ClassTipoDocInfo
Dim iTipoMovto As Integer

On Error GoTo Erro_Carrega_Series

    'Formata o produto
    lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoRastro_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141803
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141804
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 141805

    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
   
        Set objTipoDocInfo = New ClassTipoDocInfo

        objTipoDocInfo.iCodigo = iCodigo
    
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 177243
        
        iTipoMovto = objTipoDocInfo.iTipoMovtoEstoque
                
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 And iTipoMovtoAnt = iTipoMovto Then
            
            If Abs(dQuantidade - dQuantidadeAnterior) > QTDE_ESTOQUE_DELTA Or sLoteIni <> sLoteIniAnterior Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODIFICACAO_SERIES")
                If vbResult = vbNo Then gError 141866
            Else
                vbResult = vbNo
            End If
            
        Else
            vbResult = vbYes
            iTipoMovtoAnt = iTipoMovto
        End If
                    
        If vbResult = vbYes Then
            
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 141806
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.iFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col))
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141807
                
                If lErro <> SUCESSO Then gError 141808
                
                Set objItemMovEstoque = New ClassItemMovEstoque
                
                objItemMovEstoque.dQuantidade = dQuantidade 'StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col))
                objItemMovEstoque.iItemNF = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col)
                objItemMovEstoque.sAlmoxarifadoNomeRed = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col)
                objItemMovEstoque.sProduto = sProdutoFormatado
                objItemMovEstoque.sSiglaUM = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col)
                objItemMovEstoque.iTipoMov = iTipoMovto
                
                Set objAlmoxarifado = New ClassAlmoxarifado
                
                objAlmoxarifado.sNomeReduzido = objItemMovEstoque.sAlmoxarifadoNomeRed
        
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 177241
                
                objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                                
                lErro = CF("Rastreamento_Serie_Gera", objItemMovEstoque, objProduto, sLoteIni, colRastreamentoMovtoAux)
                If lErro <> SUCESSO Then gError 177244
                
            End If
            
            'Remove os dados anteriores
            For iIndice = colRastreamentoMovto.Count To 1 Step -1
                colRastreamentoMovto.Remove iIndice
            Next
            
            'Coloca os novos dados
            For Each objRastreamentoMovto In colRastreamentoMovtoAux
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
        End If
        
    End If

    Carrega_Series = SUCESSO
    
    Exit Function

Erro_Carrega_Series:

    Carrega_Series = gErr

    Select Case gErr
    
        Case 141803, 141804, 141807, 141809, 141866, 141867, 177243, 177244
        
        Case 141805
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 141806
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
        
        Case 141808
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 141810
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerie.sProduto, objRastreamentoSerie.sCodigo, objRastreamentoSerie.iFilialOP)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141811)

    End Select

    Exit Function

End Function

Public Function Acerta_Quantidade_Rastreada(ByVal iLinha As Integer)
'Acerta a quantidade do lote no grid com base na coleção global

Dim lErro As Long
Dim dQuantidade As Double
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_Acerta_Quantidade_Rastreada

    For Each objRastreamentoSerie In gcolcolRastreamentoSerie.Item(iLinha)
    
        dQuantidade = dQuantidade + objRastreamentoSerie.dQuantidade
    
    Next
                
    objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantLoteRastro_Col) = Formata_Estoque(dQuantidade)

    Acerta_Quantidade_Rastreada = SUCESSO
    
    Exit Function

Erro_Acerta_Quantidade_Rastreada:

    Acerta_Quantidade_Rastreada = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141811)

    End Select

    Exit Function

End Function
'####################################################################

'####################################################################
'Inserido por Wagner 02/06/2006
Private Function ProdutoRastro_Saida_Celula(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dQuantidade As Double
Dim objNFiscal As New ClassNFiscal
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iLinha As Integer
Dim iAchou As Integer
Dim sProduto1 As String
Dim sProduto2 As String

On Error GoTo Erro_ProdutoRastro_Saida_Celula
    
    'Pega o item da Nota
    iItem = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col)
    
    'Formata o Produto para o BD
    lErro = CF("Produto_Formata", ProdutoRastro.Text, sProduto1, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 180041
    
    If Not (objTela.gobjNFiscal Is Nothing) Then
    
        'Se tem rastreamento por Grade ou Kit de Venda
        If objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Count <> 0 Then
                                
            'Para cada item da Grade
            For iIndice = 1 To objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Count
            
                'Se encontrar
                If objProduto.sCodigo = objTela.gobjNFiscal.colItensNF.Item(iItem).colItensRomaneioGrade.Item(iIndice).sProduto Then
                    
                    'Coloca a Unidade de Medida na Tela
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objProduto.sSiglaUMEstoque
                
                    Exit For
                
                End If
            
            Next
            
            If iCodigo <> 0 Then
            
                objTipoDocInfo.iCodigo = iCodigo
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 177963
            
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 177964
            
            Else
                objTipoDocInfo.iTipo = -1
            End If
                
            If bEscaninhoVisivel = False And iCodigo <> DOCINFO_NFISFBF And iCodigo <> DOCINFO_NFISPC And iCodigo <> DOCINFO_NFFISPC And objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFIE And objTipoDocInfo.iTipo <> TIPODOCINFO_TIPO_NFEXT And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalRemPedido" And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalRem" And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalDev" Then
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Almoxarifado_Col)
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)
            Else
            
                'mario distribuicao. Se for uma nota fiscal de entrada pegar o almoxarifado e a quantidade do grid de distribuicao
                If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRem" Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Then
                            
                    'verifica se o item está no grid de alocacoes
                    For iLinha = 1 To objTela.gobjDistribuicao.objGridDist.iLinhasExistentes
                        
                        'Formata o Produto para o BD
                        lErro = CF("Produto_Formata", objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProduto2, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 180042
                        
                        If sProduto1 = sProduto2 Then
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_AlmoxDist_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_QuantDist_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objTela.gobjDistribuicao.objGridDist.objGrid.TextMatrix(iLinha, objTela.gobjDistribuicao.iGrid_UMDist_Col)
                            Exit For
                        End If
                    Next
                   
                Else
                                       
                    'verifica se o item está no grid de alocacoes
                    For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
                    
                        'Formata o Produto para o BD
                        lErro = CF("Produto_Formata", objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_ProdutoAloc_Col), sProduto2, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 180042
        
                        If sProduto1 = sProduto2 Then
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_AlmoxAloc_Col)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_UMAloc_Col)
                            If iCodigo = DOCINFO_NFISFBF Or iCodigo = DOCINFO_NFISPC Or iCodigo = DOCINFO_NFFISPC Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRemPedido" Then objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_QuantAloc_Col)
                            Exit For
                        End If
                    Next
                                            
                    If bEditavel = False Then
                     
                        If Len(Trim(Serie.Text)) = 0 Then gError 177965
                        
                        If bTelaManutencao = False Then
                            If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 177966
                            objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
                            'If Len(Trim(NFiscal.Caption)) = 0 Then gError 177967
                            objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
                        Else
                            If Len(Trim(LblDataEmissao.Caption)) = 0 Then gError 177966
                            objNFiscal.dtDataEmissao = CDate(LblDataEmissao.Caption)
                            'If Len(Trim(Numero.Text)) = 0 Then gError 177967
                            objNFiscal.lNumNotaFiscal = StrParaLong(Numero.Text)
                        End If
                        
                        objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value) 'Serie.Text
                    
                         'Lê o Almoxarifado
                        objAlmoxarifado.sNomeReduzido = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxRastro_Col)
                        
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 177971
                
                        'Se não encontrou Almoxarifado --> Erro
                        If lErro = 25060 Then gError 177972
                        
                        'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                        lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo)
                        If lErro <> SUCESSO And lErro <> 83345 Then gError 177968
                        
                        '#####################################################
                        'Inserido por Wagner 25/05/2006
                        If lErro <> SUCESSO Then
                            'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                            lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA, iItem, objAlmoxarifado.iCodigo, objProduto.sCodigo)
                            If lErro <> SUCESSO And lErro <> 180061 Then gError 177968
                        End If
                        '#####################################################
                        
                        If lErro = SUCESSO Then
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = Formata_Estoque(dQuantidade)
                            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                        Else
                            'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                            lErro = CF("MovEstoque_Le_QuantItemNF", objNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo)
                            If lErro <> SUCESSO And lErro <> 83345 Then gError 177969
                            
                            '#####################################################
                            'Inserido por Wagner 25/05/2006
                            If lErro <> SUCESSO Then
                                'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros
                                lErro = CF("MovEstoque_Le_QuantItemNF_Grade", objTela.gobjNFiscal, dQuantidade, MOV_EST_NF_VENDA_MAT_CONSIG, iItem, objAlmoxarifado.iCodigo, objProduto.sCodigo)
                                If lErro <> SUCESSO And lErro <> 180061 Then gError 177969
                            End If
                            '#####################################################
                            
                            'se não houver movimento de estoque para este item/almoxarifado ==> erro
                            'If lErro = 83345 Then gError 177970
                    
                            If lErro <> SUCESSO Then
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)
                                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_EscaninhoRastro_Col) = ESCANINHO_RASTRO_DISPONIVEL
                                'verifica se o item está no grid de alocacoes
                                For iLinha = 1 To objTela.objGridAlocacoes.iLinhasExistentes
                                    If iItem = CInt(objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_Item_Col)) And ProdutoRastro.Text = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_ProdutoAloc_Col) Then
                                        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantRastro_Col) = objTela.objGridAlocacoes.objGrid.TextMatrix(iLinha, objTela.iGrid_QuantAloc_Col)
                                        Exit For
                                    End If
                                Next
                            End If
                            
                        End If
                    End If
                End If
            End If
        End If
    End If

    ProdutoRastro_Saida_Celula = SUCESSO
    
    Exit Function

Erro_ProdutoRastro_Saida_Celula:

    ProdutoRastro_Saida_Celula = gErr

    Select Case gErr
    
        Case 177963, 177968, 177969, 177971, 180041, 180042, 180043
    
        Case 177964
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
    
        Case 177965
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 177966
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            
        Case 177967
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 177970
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_MOVESTOQUE_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)

        Case 177972
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177960)

    End Select

    Exit Function
    
End Function
'####################################################################

Public Function Atualiza_Troca_Quantidade_Grade(ByVal iItemAtual As Integer, ByVal dFator As Double) As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Atualiza_Troca_Quantidade_Grade

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(Trim(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))) > 0 Then
            If CInt(objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col)) = iItemAtual Then
                objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantRastro_Col) = Format(StrParaDbl(objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantRastro_Col)) * dFator, "STANDARD")
            End If
        End If
    Next

    Atualiza_Troca_Quantidade_Grade = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_Quantidade_Grade:

    Atualiza_Troca_Quantidade_Grade = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180065)

    End Select
    
    Exit Function

End Function

Public Function Atualiza_Troca_UM_Grade(ByVal iItemAtual As Integer, ByVal sUMAnt As String, ByVal sUMAtual As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dFator As Double

On Error GoTo Erro_Atualiza_Troca_UM_Grade

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUMAtual, sUMAnt, dFator)
    If lErro <> SUCESSO Then gError 180066

    lErro = Atualiza_Troca_Quantidade_Grade(iItemAtual, dFator)
    If lErro <> SUCESSO Then gError 180067

    Atualiza_Troca_UM_Grade = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Troca_UM_Grade:

    Atualiza_Troca_UM_Grade = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180068)

    End Select
    
    Exit Function

End Function

Public Function Carrega_Rastro(ByVal objItemMovEst As ClassItemMovEstoque, ByVal iItemNF As Integer) As Long
'Carrega rastreamento de um item da nota

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLote As New ClassRastreamentoLote
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim iLinha As Integer
Dim sProdutoMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim colRastreamentoMovto As Collection
Dim colMovEstoque As Collection
Dim sProdutoEnxuto As String
Dim objProduto As ClassProduto
Dim iIndice2 As Integer 'Inserido por Wagner
Dim colRastreamentoSerie As Collection


On Error GoTo Erro_Carrega_Rastro

    If gcolcolRastreamentoSerie Is Nothing Then
        Set gcolcolRastreamentoSerie = New Collection
    End If

'    'Para cada item da nota fiscal
'    For Each objItemNF In objNFiscal.colItensNF
'
'        '##############################################################
'        'Se ProdutoRastro for uma Combo => Limpa
        If TypeName(ProdutoRastro) = "ComboBox" Then
            ProdutoRastro.Clear

            If objItemMovEst.colRastreamentoMovto.Count > 0 Then
        
                lErro = Mascara_RetornaProdutoTela(objItemMovEst.colRastreamentoMovto(1).sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 180047
                                    
                ProdutoRastro.AddItem sProdutoMascarado
     
            End If
        
        End If
'
'        'Para cada Item da Grade ou componente do Kit de Venda
'        For iIndice2 = 1 To objItemNF.colItensRomaneioGrade.Count
'
'            Set objProduto = New ClassProduto
'
'            'Lê o Produto
'            objProduto.sCodigo = objItemNF.colItensRomaneioGrade.Item(iIndice2).sProduto
'
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 180046
'
'            'Se ele for rastreado adiciona na combo
'            If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
'
'                lErro = Mascara_RetornaProdutoTela(objItemNF.colItensRomaneioGrade.Item(iIndice2).sProduto, sProdutoMascarado)
'                If lErro <> SUCESSO Then gError 180047
'
'                ProdutoRastro.AddItem sProdutoMascarado
'
'            End If
'
'        Next
        '################################################################


'        If objItemNF.iAlmoxarifado <> 0 Then

            'Lê item de movimento de estoque
'            objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
'            objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
'            objItemMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
'
'            Set colMovEstoque = New Collection
'
'            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
'            If lErro <> SUCESSO Then gError 75937
'
'            '############################################################
'            'Inserido por Wagner 02/06/2006
'            'Le os movimento gerados por grade ou componentes de um kit de venda
'            lErro = CF("MovEstoque_Le_ItemNF_Grade", objItemMovEstoque, colMovEstoque)
'            If lErro <> SUCESSO Then gError 180048
'            '############################################################
'
'            For Each objItemMovEstoque In colMovEstoque
            
            
                    'Lê o Almoxarifado
                    objAlmoxarifado.iCodigo = objItemMovEst.iAlmoxarifado
                    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25056 Then gError 75935
            
                    'Se não encontrou Almoxarifado --> Erro
                    If lErro = 25056 Then gError 75938
            
'                    'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
'                    Set colRastreamentoMovto = New Collection
'
'                    lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
'                    If lErro <> SUCESSO Then gError 75934
'
'                    Set objItemNF.colRastreamento = New Collection
''                    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 10/03/2006
             
                    iLinha = objGridInt.iLinhasExistentes
                
             
                    If objItemMovEst.colRastreamentoMovto.Count > 0 Then
             
'                    'Guarda as quantidades alocadas dos lotes
'                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
            
                        iLinha = iLinha + 1
                        
                        gcolcolRastreamentoSerie.Add objItemMovEst.colRastreamentoMovto
            
                        Set objRastreamentoMovto = objItemMovEst.colRastreamentoMovto(1)
                        
                        
                        'procura RastreamentoLote
'                         objRastreamentoLote.sProduto = objItemNF.sProduto 'O produto Rastreado pode não ser o item da Nota associado (Grade e Kit de Venda)
                        objRastreamentoLote.sProduto = objRastreamentoMovto.sProduto
                        objRastreamentoLote.sCodigo = objRastreamentoMovto.sLote
                        
                        lErro = CF("RastreamentoLote_Le", objRastreamentoLote)
                        If lErro <> SUCESSO And lErro <> 75710 Then gError 75936
            
                        'Se não encontrou, erro
                        If lErro = 75710 Then gError 75939
            
            
                        'Formata o Produto
'                        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                        lErro = Mascara_RetornaProdutoEnxuto(objRastreamentoMovto.sProduto, sProdutoEnxuto)
                        If lErro <> SUCESSO Then gError 83198
                
                        If TypeName(ProdutoRastro) <> "ComboBox" Then
                            ProdutoRastro.PromptInclude = False
                            ProdutoRastro.Text = sProdutoEnxuto
                            ProdutoRastro.PromptInclude = True
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = ProdutoRastro.Text
                        Else
                            lErro = Mascara_RetornaProdutoTela(objRastreamentoMovto.sProduto, sProdutoEnxuto)
                            If lErro <> SUCESSO Then gError 188057
                        
                            Call CF("sCombo_Seleciona2", ProdutoRastro, sProdutoEnxuto)
                            objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoRastro_Col) = sProdutoEnxuto
                        End If
                
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = iItemNF
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxRastro_Col) = objAlmoxarifado.sNomeReduzido
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantRastro_Col) = Formata_Estoque(objItemMovEst.dQuantidade)
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastreamentoMovto.sLote
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteDataRastro_Col) = Format(objRastreamentoLote.dtDataEntrada, "dd/mm/yyyy")
                        
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantLoteRastro_Col) = Formata_Estoque(objItemMovEst.dQuantidade)
                        
                        '###########################################################
                        
                        objGridInt.objGrid.TextMatrix(iLinha, iGrid_UMRastro_Col) = objRastreamentoMovto.sSiglaUM
                        
                        objGridInt.objGrid.Row = iLinha
                        
                        
                        
                        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
                    
                    End If
'                    Next
                    
'                End If
    
 '           Next
            
 '       End If
        
 '   Next
    
    Carrega_Rastro = SUCESSO

    Exit Function

Erro_Carrega_Rastro:

    Carrega_Rastro = gErr

    Select Case gErr

        Case 75934, 75935, 75936, 75937, 83199, 141952, 180046, 180047, 180048, 188057, 199181

        Case 75938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)

        Case 75939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO1", gErr, objRastreamentoLote.sProduto, objRastreamentoLote.sCodigo, objRastreamentoLote.iFilialOP)

        Case 83198
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)

        Case 83200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152757)
    
    End Select

    Exit Function

End Function


