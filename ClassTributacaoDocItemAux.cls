VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoDocItemAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'area de trabalho amarrada a um doc item (item de um pedido, NF,...)

Private colVariavelCalculo As New Collection 'colecao de objVariavelCalculo

'+ colecao de categorias a que pertence o produto com os respectivos valores (serve p/pesquisar excecoes de IPI e ICMS)
'+ dados especificos de cada driver: posso ter um obj p/cada driver
'
'Interface:
'==========
'
'na criacao:
'
'Qdo trocar produto, tipo trib, valor item, qtde item, tipo trib ou nat op

Private mvarobjDocItem As ClassTributoDocItem
Private mvarobjVarAux As ClassVarAux

Private mvariJaCalculado As Integer

'Interface:
'==========
'
Function Ativar(objDocItem As ClassTributoDocItem) As Long

    'guarda objDocItem
    Set mvarobjDocItem = objDocItem
    
    'cria objVarAux
    Set mvarobjVarAux = New ClassVarAux
    
    mvariJaCalculado = 0
    
End Function

Function Alterado() As Long

    mvariJaCalculado = 0

End Function

Function Calculado() As Long

    mvariJaCalculado = 1

End Function

Function iJaCalculado() As Integer

    iJaCalculado = mvariJaCalculado
    
End Function
    
Function Desativar() As Long

    'apaga referencia a objDocItem
    Set mvarobjDocItem = Nothing
    
    'apaga referencia a objVarAux
    Set mvarobjVarAux = Nothing
    
End Function
        
Function ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'vai ser chamado pelo objDocItem quando este nao conseguir "tratar"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long, colCateg As New Collection
On Error GoTo Erro_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    Select Case sIdentificador
        
        Case "I_PRODUTO_COL_CATEG"
        
            lErro = ObterColCategProd(colCateg)
            If lErro <> SUCESSO Then Error 27358
            
            objVariavelCalculo.sIdentificador = "I_PRODUTO_COL_CATEG"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = colCateg
            
        Case "OBJ_ITEM_NF_ORIG"
            objVariavelCalculo.sIdentificador = "OBJ_ITEM_NF_ORIG"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = Nothing
            
        Case Else
 '           erro ?
            
    End Select

    ObterVar = SUCESSO

    Exit Function

Erro_ObterVar:

    ObterVar = Err

    Select Case Err

        Case 27358
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153802)

    End Select

    Exit Function

End Function

Function GuardarVar(objVariavelCalculo As ClassVariavelCalculo)
'vai ser chamado pelo objDocItem quando este nao conseguir "tratar"
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_GuardarVar
        
    lErro = mvarobjVarAux.GuardarVar(objVariavelCalculo)
    If lErro <> SUCESSO Then Error 27359
    
    GuardarVar = SUCESSO

    Exit Function

Erro_GuardarVar:

    GuardarVar = Err

    Select Case Err

        Case 27359
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153803)

    End Select

    Exit Function

End Function

Function ObterColCategProd(colCateg As Collection) As Long
'faz com que colCateg aponte p/uma colecao com as (categoria,item) associadas a um produto
Dim lErro As Long
Dim objVarProduto As ClassVariavelCalculo, objVarProdutoAtual As ClassVariavelCalculo
Dim objVarColCateg As New ClassVariavelCalculo
On Error GoTo Erro_ObterColCategProd

    lErro = mvarobjDocItem.ObterVar("PRODUTO_CODIGO", objVarProduto)
    If lErro <> SUCESSO Then Error 27360
        
    lErro = mvarobjVarAux.ObterVar("I_CCP_PRODUTO", objVarProdutoAtual)
    If lErro <> SUCESSO Then Error 27361
        
    'se nao houve modificacao
    If (objVarProdutoAtual.iStatus <> VAR_PREENCH_VAZIO And objVarProdutoAtual.vValor = objVarProduto.vValor) Then
        
        lErro = mvarobjVarAux.ObterVar("I_PRODUTO_COL_CATEG", objVarColCateg)
        If lErro <> SUCESSO Then Error 27362
        
        Set colCateg = objVarColCateg.vValor
        
    Else
    
        lErro = CF("ProdutoCategorias_Le_Produto", objVarProduto.vValor, colCateg)
        If lErro <> SUCESSO Then Error 27363

        Set objVarColCateg.vValor = colCateg
        objVarColCateg.sIdentificador = "I_PRODUTO_COL_CATEG"
        objVarColCateg.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "I_PRODUTO_COL_CATEG"
        lErro = mvarobjVarAux.GuardarVar(objVarColCateg)
        If lErro <> SUCESSO Then Error 27364

        'atualizar "I_CCP_PRODUTO"
        objVarProdutoAtual.vValor = objVarProduto.vValor
        objVarProdutoAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarProdutoAtual)
        If lErro <> SUCESSO Then Error 27365
    
    End If

    ObterColCategProd = SUCESSO

    Exit Function

Erro_ObterColCategProd:

    ObterColCategProd = Err

    Select Case Err

        Case 27360 To 27365
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153804)

    End Select

    Exit Function

End Function

