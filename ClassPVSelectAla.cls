VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PV_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim sTabela As String
Dim dtDataEntrega As Date
Dim objPVAla As New ClassPedidoDeVendaAla

On Error GoTo Erro_PV_Le_Customizado

     'Abre Comando
     For iIndice = LBound(alComando) To UBound(alComando)
         alComando(iIndice) = Comando_Abrir()
         If alComando(iIndice) = 0 Then gError 140306
     Next
     
     sTabela = "PedidosDeVenda"
    
     'Tenta ler Pedido de Venda e Pedido de Venda Baixado
     For iIndice = 1 To 2
        
        'Le o Pedido de Venda
        lErro = Comando_Executar(alComando(1), "SELECT DataEntregaReal FROM " & sTabela & " WHERE NumIntDoc= ? ", dtDataEntrega, objPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 140307
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140308
    
        If lErro = AD_SQL_SUCESSO Then Exit For
                 
        sTabela = "PedidosDeVendaBaixados"
    
    Next
    
    If lErro = AD_SQL_SUCESSO Then
        objPVAla.dtDataEntregaReal = dtDataEntrega
    Else
        objPVAla.dtDataEntregaReal = DATA_NULA
    End If
    
    Set objPV.objInfoUsu = objPVAla
          
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PV_Le_Customizado = SUCESSO

    Exit Function

Erro_PV_Le_Customizado:

    PV_Le_Customizado = gErr

        Select Case gErr

        Case 140306
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140307, 140308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 140309
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPV.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
