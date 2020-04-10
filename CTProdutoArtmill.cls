VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdutoArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const CONVERSAO_METRO_PARA_MILIMETRO = 1000 ' Inserido por Wagner

Public Property Get DetalheCor(ByVal objCT As Object) As Object
     Set DetalheCor = objCT.objUserControl.Controls("DetalheCor")
End Property

Public Property Get DimEmbalagem(ByVal objCT As Object) As Object
     Set DimEmbalagem = objCT.objUserControl.Controls("DimEmbalagem")
End Property

Public Property Get CodAnterior(ByVal objCT As Object) As Object
     Set CodAnterior = objCT.objUserControl.Controls("CodAnterior")
End Property

'colocar em customizacao de form_load:
Public Function Prod_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Form_Load

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_PRODUTO_COR, objCT.Cor, False, False)
    If lErro <> SUCESSO Then gError 124146
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_PRODUTO_DETALHE_COR, DetalheCor(objCT), False, False)
    If lErro <> SUCESSO Then gError 124147
    
    Prod_Form_Load = SUCESSO
    
    Exit Function
    
Erro_Prod_Form_Load:

    Prod_Form_Load = gErr

    Select Case gErr

        Case 124146, 124147
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184268)

    End Select
    
    Exit Function

End Function

Function Prod_Traz_Produto_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Traz_Produto_Tela

    lErro = CF("Produto_Le_InfoUsu", objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 124047
    If lErro <> SUCESSO Then gError 124048

    Prod_Traz_Produto_Tela = SUCESSO
    
    Exit Function
    
Erro_Prod_Traz_Produto_Tela:

    Prod_Traz_Produto_Tela = gErr

    Select Case gErr

        Case 124047, 124048
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184269)

    End Select
    
    Exit Function

End Function

Function Prod_TrazCarac_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'em Traz_TabCaracteristicas_Tela

Dim lErro As Long

On Error GoTo Erro_Prod_TrazCarac_Tela

    If objProduto.dComprimento > 0 Then objCT.Comprimento.Text = objProduto.dComprimento * CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner
    If objProduto.dLargura > 0 Then objCT.Largura.Text = objProduto.dLargura * CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner
    If objProduto.dEspessura > 0 Then objCT.Espessura.Text = objProduto.dEspessura * CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner
    
    DetalheCor(objCT).Text = objProduto.objInfoUsu.sDetalheCor
    DimEmbalagem(objCT).Text = objProduto.objInfoUsu.sDimEmbalagem
    CodAnterior(objCT).Text = objProduto.objInfoUsu.sCodAnterior

    Prod_TrazCarac_Tela = SUCESSO
    
    Exit Function
    
Erro_Prod_TrazCarac_Tela:

    Prod_TrazCarac_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184271)

    End Select
    
    Exit Function

End Function

    
Function Prod_Traz_TabPrecos_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'traz os dados do tab tabela de preços do BD para a tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colTabelaPrecoItem As New Collection
Dim colTabPrecoEmpresa As New Collection
Dim iLinha As Integer
Dim objTabelaPrecoItemEmp As New ClassTabelaPrecoItem
Dim iAchou As Integer
Dim iTotalLinhas As Integer
Dim iIndiceEmp As Integer
Dim iIndiceItem As Integer

On Error GoTo Erro_Prod_Traz_TabPrecos_Tela

    'Limpa o Grid de Tabelas de Preço
    Call Grid_Limpa(objCT.objGridTabelaPreco)

    'Lê os preços do Produtos nas Tabela
    lErro = CF("Produto_Le_TabelaPrecoItem", objProduto, colTabelaPrecoItem, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 64437
    
    'No caso de não ser empresa toda mostra na tela o preço e empresa toda
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'Lê os preços para empresa toda
        lErro = CF("Produto_Le_TabelaPrecoItem", objProduto, colTabPrecoEmpresa, EMPRESA_TODA)
        If lErro <> SUCESSO Then Error 64438
        
        For Each objTabelaPrecoItemEmp In colTabPrecoEmpresa
        
            iAchou = 0
        
            For Each objTabelaPrecoItem In colTabelaPrecoItem
        
                If objTabelaPrecoItemEmp.iCodTabela = objTabelaPrecoItem.iCodTabela Then
                    iAchou = 1
                    Exit For
                End If
        
            Next
            
            If iAchou = 0 Then iTotalLinhas = iTotalLinhas + 1
            
        Next
    
    End If
    
    'Redimenciona o Grid
    If colTabelaPrecoItem.Count + iTotalLinhas >= objCT.objGridTabelaPreco.objGrid.Rows Then
        objCT.objGridTabelaPreco.objGrid.Rows = colTabelaPrecoItem.Count + iTotalLinhas + 1
        Call Grid_Inicializa(objCT.objGridTabelaPreco)
'    Else
'        objCT.objGridTabelaPreco.objGrid.Rows = 11
'        Call Grid_Inicializa(objCT.objGridTabelaPreco)
    End If
    
    iIndiceItem = 1
    iIndiceEmp = 1
    
    Do While iIndiceItem <= colTabelaPrecoItem.Count And iIndiceEmp <= colTabPrecoEmpresa.Count
    
        Set objTabelaPrecoItem = colTabelaPrecoItem.Item(iIndiceItem)
        Set objTabelaPrecoItemEmp = colTabPrecoEmpresa.Item(iIndiceEmp)
        
        iIndice = iIndice + 1
    
        If objTabelaPrecoItem.iCodTabela < objTabelaPrecoItemEmp.iCodTabela Then
            
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_Tabela_Col) = objTabelaPrecoItem.iCodTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DescricaoTabela_Col) = objTabelaPrecoItem.sDescricaoTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorFilial_Col) = Format(objTabelaPrecoItem.dPreco, "Standard")
            If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DataPreco_Col) = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
            
            iIndiceItem = iIndiceItem + 1
            
        ElseIf objTabelaPrecoItemEmp.iCodTabela < objTabelaPrecoItem.iCodTabela Then
    
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_Tabela_Col) = objTabelaPrecoItemEmp.iCodTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DescricaoTabela_Col) = objTabelaPrecoItemEmp.sDescricaoTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorEmpresa_Col) = Format(objTabelaPrecoItemEmp.dPreco, "Standard")
            If objTabelaPrecoItemEmp.dtDataVigencia <> DATA_NULA Then objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DataPreco_Col) = Format(objTabelaPrecoItemEmp.dtDataVigencia, "dd/mm/yyyy")
            
            iIndiceEmp = iIndiceEmp + 1

        Else
    
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_Tabela_Col) = objTabelaPrecoItem.iCodTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DescricaoTabela_Col) = objTabelaPrecoItem.sDescricaoTabela
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorFilial_Col) = Format(objTabelaPrecoItem.dPreco, "Standard")
            objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorEmpresa_Col) = Format(objTabelaPrecoItemEmp.dPreco, "Standard")
            If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DataPreco_Col) = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
            
            iIndiceItem = iIndiceItem + 1
            
            iIndiceEmp = iIndiceEmp + 1
    
        End If
    
        objCT.objGridTabelaPreco.iLinhasExistentes = objCT.objGridTabelaPreco.iLinhasExistentes + 1
    
    Loop
    
    
    Do While iIndiceItem <= colTabelaPrecoItem.Count
    
        Set objTabelaPrecoItem = colTabelaPrecoItem.Item(iIndiceItem)
    
        iIndice = iIndice + 1
    
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_Tabela_Col) = objTabelaPrecoItem.iCodTabela
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DescricaoTabela_Col) = objTabelaPrecoItem.sDescricaoTabela
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorFilial_Col) = Format(objTabelaPrecoItem.dPreco, "Standard")
        If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DataPreco_Col) = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
        
        objCT.objGridTabelaPreco.iLinhasExistentes = objCT.objGridTabelaPreco.iLinhasExistentes + 1

        iIndiceItem = iIndiceItem + 1
    
    Loop
    
    
    Do While iIndiceEmp <= colTabPrecoEmpresa.Count
    
        Set objTabelaPrecoItemEmp = colTabPrecoEmpresa.Item(iIndiceEmp)
    
        iIndice = iIndice + 1
    
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_Tabela_Col) = objTabelaPrecoItemEmp.iCodTabela
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DescricaoTabela_Col) = objTabelaPrecoItemEmp.sDescricaoTabela
        objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_ValorEmpresa_Col) = Format(objTabelaPrecoItemEmp.dPreco, "Standard")
        If objTabelaPrecoItemEmp.dtDataVigencia <> DATA_NULA Then objCT.GridTabelaPreco.TextMatrix(iIndice, objCT.iGrid_DataPreco_Col) = Format(objTabelaPrecoItemEmp.dtDataVigencia, "dd/mm/yyyy")
        
        objCT.objGridTabelaPreco.iLinhasExistentes = objCT.objGridTabelaPreco.iLinhasExistentes + 1

        iIndiceEmp = iIndiceEmp + 1
   
    Loop
    
    objCT.DescrUM.Caption = objProduto.sSiglaUMVenda
    
    Prod_Traz_TabPrecos_Tela = SUCESSO

    Exit Function

Erro_Prod_Traz_TabPrecos_Tela:

    Prod_Traz_TabPrecos_Tela = Err

    Select Case Err

        Case 64437, 64438

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 177516)

    End Select

    Exit Function

End Function

Function Prod_Gravar_Registro(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Gravar_Registro

    '############################################
    'Inserido por Wagner 17/01/2006
    lErro = Verifica_Troca_UMEstoque(objProduto)
    If lErro <> SUCESSO Then gError 141531
    '############################################
    
    Prod_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Prod_Gravar_Registro:

    Prod_Gravar_Registro = gErr

    Select Case gErr

        Case 141531
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184273)

    End Select
    
    Exit Function

End Function

Function Prod_MoveCaracMemoria(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'em Move_TabCaracteristicas_Memoria

Dim lErro As Long

On Error GoTo Erro_Prod_MoveCaracMemoria

    If Len(Trim(objCT.Comprimento.Text)) > 0 Then objProduto.dComprimento = CDbl(objCT.Comprimento.Text) / CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner
    If Len(Trim(objCT.Largura.Text)) > 0 Then objProduto.dLargura = CDbl(objCT.Largura.Text) / CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner
    If Len(Trim(objCT.Espessura.Text)) > 0 Then objProduto.dEspessura = CDbl(objCT.Espessura.Text) / CONVERSAO_METRO_PARA_MILIMETRO 'Alterado por Wagner

    Set objProduto.objInfoUsu = New ClassProdutoInfoUsu
    
    objProduto.objInfoUsu.sCodigo = objProduto.sCodigo
    objProduto.objInfoUsu.sDetalheCor = DetalheCor(objCT).Text
    objProduto.objInfoUsu.sDimEmbalagem = DimEmbalagem(objCT).Text
    objProduto.objInfoUsu.sCodAnterior = CodAnterior(objCT).Text
    
    Prod_MoveCaracMemoria = SUCESSO
    
    Exit Function
    
Erro_Prod_MoveCaracMemoria:

    Prod_MoveCaracMemoria = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184272)

    End Select
    
    Exit Function

End Function

'################################################
'Inserido por Wagner 17/01/2006
Private Function Verifica_Troca_UMEstoque(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoBD As New ClassProduto
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Verifica_Troca_UMEstoque

    objProdutoBD.sCodigo = objProduto.sCodigo

    'Lê o Produto no BD
    lErro = CF("Produto_Le", objProdutoBD)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141532
    
    'Se já está cadastrado
    If lErro = SUCESSO Then
        
        'Se houve troca na UM de estoque
        If objProduto.sSiglaUMEstoque <> objProdutoBD.sSiglaUMEstoque Then
        
            'Avisa da possível demora na gravação
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_TROCA_UMESTOQUE")
            If vbResult = vbNo Then gError 141533
        
        End If
    
    End If
    
    Verifica_Troca_UMEstoque = SUCESSO

    Exit Function

Erro_Verifica_Troca_UMEstoque:

    Verifica_Troca_UMEstoque = gErr

    Select Case gErr
    
        Case 141532, 141533

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 177554)

    End Select

    Exit Function

End Function
'################################################

Sub Cor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Sub Cor_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Sub Cor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cor_Validate

    lErro = CF("CamposGenericos_Validate2", CAMPOSGENERICOS_PRODUTO_COR, objCT.Cor, "AVISO_CRIAR_COR")
    If lErro <> SUCESSO Then gError 102417
    
    Exit Sub

Erro_Cor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177458)

    End Select

End Sub

Sub DetalheCor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Sub DetalheCor_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Sub DetalheCor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DetalheCor_Validate

    lErro = CF("CamposGenericos_Validate2", CAMPOSGENERICOS_PRODUTO_DETALHE_COR, DetalheCor(objCT), "AVISO_CRIAR_DETALHE_COR")
    If lErro <> SUCESSO Then gError 102417
    
    Exit Sub

Erro_DetalheCor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177553)

    End Select

End Sub

Function Prod_Limpa_Tela(ByVal objCT As Object) As Long
        
    objCT.Cor.ListIndex = -1
    DetalheCor(objCT).ListIndex = -1
    
    Prod_Limpa_Tela = SUCESSO
    
End Function
