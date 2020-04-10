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

Public Function ItensOV_Le_Customizado(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim lComando As Long
Dim dPrecoUnitImp As Double, iTabelaPrecoImp As Integer
Dim objItensOVWShip As ClassItemOVWShip
Dim objItensOV As ClassItemOV

On Error GoTo Erro_ItensOV_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 210453
    
    For Each objItensOV In objOrcamentoVenda.colItens
    
        If objOrcamentoVenda.iHistorico <> MARCADO Then
            lErro = Comando_Executar(lComando, "SELECT PrecoUnitImp, TabelaPrecoImp FROM ItensOrcamentoVenda WHERE NumIntDoc = ? ", dPrecoUnitImp, iTabelaPrecoImp, objItensOV.lNumIntDoc)
        Else
            lErro = Comando_Executar(lComando, "SELECT PrecoUnitImp, TabelaPrecoImp FROM ItensOrcamentoVendaHist WHERE NumIntDoc = ? ", dPrecoUnitImp, iTabelaPrecoImp, objItensOV.lNumIntDoc)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 210454

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 210455
    
        If lErro = AD_SQL_SEM_DADOS Then gError 210456
        
        Set objItensOVWShip = New ClassItemOVWShip

        objItensOVWShip.dPrecoUnitImp = dPrecoUnitImp
        objItensOVWShip.iTabelaPrecoImp = iTabelaPrecoImp

        Set objItensOV.objInfoUsu = objItensOVWShip

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensOV_Le_Customizado = SUCESSO

    Exit Function

Erro_ItensOV_Le_Customizado:

    ItensOV_Le_Customizado = gErr

    Select Case gErr

        Case 210453
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 210454, 210455
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 210456
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210457)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

