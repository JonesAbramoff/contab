VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PedidoCompra_Valida_Campos_Cust(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long
Dim sUM As String
Dim dQuantidade As Double
Dim sProduto As String
Dim iMoeda As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo

On Error GoTo Erro_PedidoCompra_Valida_Campos_Cust

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 210944
                
    If Len(objPedidoCompra.sObsEmbalagem) > 0 Then
                
        'Lê os Pedidos de Compra associados as Itens Da Concorrência
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM PedidoCompraTodos WHERE ObsEmbalagem = ? AND Codigo <> ? AND FilialEmpresa = ?", lCodigo, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 210945
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210946
        
        If lErro = SUCESSO Then gError 210947
    
    End If
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    PedidoCompra_Valida_Campos_Cust = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Valida_Campos_Cust:

    PedidoCompra_Valida_Campos_Cust = gErr
    
    Select Case gErr
    
        Case 210944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 210945, 210946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRATODOS", gErr)
            
        Case 210947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEM_COMPRA_CADASTRADA", gErr, objPedidoCompra.sObsEmbalagem, lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210948)
        
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

    PedidoCompra_Valida_Campos_Cust = SUCESSO
End Function

