VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PV_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objPVMotivo As ClassPVMotivo
Dim objItemPV As ClassItemPedido
Dim iFilialEmpresa As Long
Dim sMotivo As String
Dim dPrecoTabela As Double

On Error GoTo Erro_PV_Traz_Pedido_Tela

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197746
    Next
    
    Set objCT.gobjInfoUsu.gcolMotivos = New Collection

    For Each objItemPV In objPedidoVenda.colItensPedido
        
        Set objPVMotivo = New ClassPVMotivo
    
        sMotivo = String(STRING_PVMOTIVO_MOTIVO, 0)
    
        'Pesquisa os ítens vinculados ao item
        lErro = Comando_Executar(alComando(0), "SELECT Motivo, PrecoTabela FROM ItensPVMotivo WHERE FilialEmpresa = ? And CodPedido = ? AND Produto = ? ", sMotivo, dPrecoTabela, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 197747

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197748

        If lErro = AD_SQL_SUCESSO Then
        
            objPVMotivo.sProduto = objItemPV.sProduto
            objPVMotivo.sMotivo = sMotivo
            objPVMotivo.dPrecoTabela = dPrecoTabela
        
        End If
    
        objCT.gobjInfoUsu.gcolMotivos.Add objPVMotivo

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PV_Traz_Pedido_Tela = SUCESSO

    Exit Function

Erro_PV_Traz_Pedido_Tela:

    PV_Traz_Pedido_Tela = gErr

    Select Case gErr

        Case 197746
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197747, 197748
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVMOTIVO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objPVMotivo.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197749)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function PV_Preco_Menor_Tabela(ByVal objPedidoVenda As ClassPedidoDeVenda, colPVMotivo As Collection) As Long
    
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objPVMotivo As ClassPVMotivo
Dim objItemPV As ClassItemPedido
Dim iFilialEmpresa As Long
Dim sMotivo As String
Dim dPrecoTabela As Double

On Error GoTo Erro_PV_Preco_Menor_Tabela

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197757
    Next
    
    For Each objItemPV In objPedidoVenda.colItensPedido
    
        sMotivo = String(STRING_PVMOTIVO_MOTIVO, 0)
    
        'Pesquisa os ítens vinculados ao item
        lErro = Comando_Executar(alComando(0), "SELECT Motivo, PrecoTabela FROM ItensPVMotivo WHERE FilialEmpresa = ? And CodPedido = ? AND Produto = ? ", sMotivo, dPrecoTabela, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 197758

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197759

        If lErro = AD_SQL_SUCESSO Then
        
            Set objPVMotivo = New ClassPVMotivo
        
            objPVMotivo.sProduto = objItemPV.sProduto
            objPVMotivo.sMotivo = sMotivo
            objPVMotivo.dPrecoTabela = dPrecoTabela
            objPVMotivo.dPrecoInformado = objItemPV.dPrecoUnitario
        
            colPVMotivo.Add objPVMotivo
        
        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PV_Preco_Menor_Tabela = SUCESSO

    Exit Function

Erro_PV_Preco_Menor_Tabela:

    PV_Preco_Menor_Tabela = gErr

    Select Case gErr

        Case 197757
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197758, 197759
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVMOTIVO", gErr, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objPVMotivo.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197760)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function PV_Preco_Menor_Tabela_Texto(ByVal objPedidoVenda As ClassPedidoDeVenda, sTexto As String) As Long
    
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objPVMotivo As ClassPVMotivo
Dim objItemPV As ClassItemPedido
Dim iFilialEmpresa As Long
Dim sMotivo As String
Dim dPrecoTabela As Double
Dim colPVMotivo As New Collection

On Error GoTo Erro_PV_Preco_Menor_Tabela_Texto


    lErro = CF("PV_Preco_Menor_Tabela", objPedidoVenda, colPVMotivo)
    If lErro <> SUCESSO Then gError 197761
    
    For Each objPVMotivo In colPVMotivo
    
        sTexto = sTexto & "O produto " & objPVMotivo.sProduto & " possui preço de tabela = R$ " & Format(objPVMotivo.dPrecoTabela, gobjFAT.sFormatoPrecoUnitario) & " e foi informado o preço R$ " & Format(objPVMotivo.dPrecoInformado, gobjFAT.sFormatoPrecoUnitario) & ". e o motivo foi " & objPVMotivo.sMotivo & ". " & Chr(10)
        
    Next

    PV_Preco_Menor_Tabela_Texto = SUCESSO

    Exit Function

Erro_PV_Preco_Menor_Tabela_Texto:

    PV_Preco_Menor_Tabela_Texto = gErr

    Select Case gErr

        Case 197761
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197762)

    End Select

    Exit Function
    
End Function

Public Function Importa_Dados_Cliente_Arq_Aux(sNomeFunc As String) As Long
    sNomeFunc = "Importa_Dados_Cliente_Arq_Aux2"
    Importa_Dados_Cliente_Arq_Aux = SUCESSO
End Function

Public Function Integracao_Obtem_NomeID_Cust(ByVal sRegistro As String, sNomeArqID As String) As Long
    If Len(sRegistro) >= 25 Then
        sNomeArqID = Left(sRegistro, 8) & "20" & Mid(sRegistro, 14, 12) & ".txt"
    End If
    Integracao_Obtem_NomeID_Cust = SUCESSO
End Function

Public Function Importa_Dados_Cliente_ProcuraCGC(bProcuraPeloCGC As Boolean) As Long
    bProcuraPeloCGC = False
    Importa_Dados_Cliente_ProcuraCGC = SUCESSO
End Function
