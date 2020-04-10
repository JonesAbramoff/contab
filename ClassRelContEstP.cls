VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelContEstP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gsCodProdutoAnt As String
Private giFilialEmpresaAnt As Integer
Private gcolItemPC As New Collection

Public Function Obter_PedidoCompra_Codigo(dCodigoPC As Double, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_Obter_PedidoCompra_Codigo

    lErro = ObterDadosContEstP(iFilialEmpresa, sCodProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iColuna > gcolItemPC.Count Or iColuna < 1 Then
        dCodigoPC = 0
    Else
        Set objItemPC = gcolItemPC(iColuna)
        dCodigoPC = objItemPC.lCodigoPC
    End If

    Obter_PedidoCompra_Codigo = SUCESSO

    Exit Function

Erro_Obter_PedidoCompra_Codigo:

    Obter_PedidoCompra_Codigo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210767)

    End Select

    Exit Function

End Function

Public Function Obter_PedidoCompra_Qtde(dQtde As Double, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objRelContEstP As ClassRelContEstP
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_Obter_PedidoCompra_Qtde

    lErro = ObterDadosContEstP(iFilialEmpresa, sCodProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iColuna > gcolItemPC.Count Or iColuna < 1 Then
        dQtde = 0
    Else
        Set objItemPC = gcolItemPC(iColuna)
        dQtde = objItemPC.dQuantidade
    End If

    Obter_PedidoCompra_Qtde = SUCESSO

    Exit Function

Erro_Obter_PedidoCompra_Qtde:

    Obter_PedidoCompra_Qtde = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210768)

    End Select

    Exit Function

End Function


Private Function ObterDadosContEstP(ByVal iFilialEmpresa As Integer, ByVal sCodProduto As String) As Long
'obtem colecao com dados de todos os pedidos de compra de um produto numa filial

Dim lErro As Long

On Error GoTo Erro_ObterDadosContEstP
 
    'Verifica se troucou parametro
    If giFilialEmpresaAnt <> iFilialEmpresa Or sCodProduto <> gsCodProdutoAnt Then
               
        Set gcolItemPC = New Collection
               
        'obtem colecao com dados de todos os pedidos de compra de um produto numa filial
        lErro = CF("ItensPC_Le_Produto", iFilialEmpresa, sCodProduto, gcolItemPC)
        If lErro <> SUCESSO Then gError 210769
                        
        giFilialEmpresaAnt = iFilialEmpresa
        gsCodProdutoAnt = sCodProduto
        
    End If
    
    ObterDadosContEstP = SUCESSO

    Exit Function

Erro_ObterDadosContEstP:

    ObterDadosContEstP = gErr

    Select Case gErr
        
        Case 210769
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210770)

    End Select

    Exit Function

End Function


Public Function Obter_PedidoCompra_Data(dtDataLimite As Date, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objRelContEstP As ClassRelContEstP
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_Obter_PedidoCompra_Data

    lErro = ObterDadosContEstP(iFilialEmpresa, sCodProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iColuna > gcolItemPC.Count Or iColuna < 1 Then
        dtDataLimite = DATA_NULA
    Else
        Set objItemPC = gcolItemPC(iColuna)
        dtDataLimite = objItemPC.dtDataLimite
    End If

    Obter_PedidoCompra_Data = SUCESSO

    Exit Function

Erro_Obter_PedidoCompra_Data:

    Obter_PedidoCompra_Data = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210768)

    End Select

    Exit Function

End Function


