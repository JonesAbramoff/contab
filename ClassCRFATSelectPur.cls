VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelectPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Vendedor_Le_Customizado(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim objVendedorPur As New ClassVendedorPur
Dim objVendPurExc As ClassVendedorPurExc
Dim dtDataInicio As Date, sObs As String
Dim sProduto As String, dPercComissao As Double, dMetaComissao As Double
Dim alComando(1 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Le_Customizado

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213784
    Next

    'Alocação de espaço no buffer
    sObs = String(STRING_MAXIMO, 0)

    'Le a tabela VendedorTRV
    lErro = Comando_Executar(alComando(1), "SELECT DataInicio, Observacao FROM VendedoresPur WHERE Vendedor= ? ", dtDataInicio, sObs, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213785

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213786
    
    objVendedorPur.dtDataInicio = DATA_NULA

    If lErro = AD_SQL_SUCESSO Then
    
        objVendedorPur.dtDataInicio = dtDataInicio
        objVendedorPur.sObservacao = sObs
        
        sProduto = String(STRING_PRODUTO, 0)
        
        'Le a tabela VendedorTRVComissao
        lErro = Comando_Executar(alComando(2), "SELECT Produto, PercComissao, MetaComissao FROM VendedoresPurExc WHERE Vendedor = ? ORDER BY Seq", sProduto, dPercComissao, dMetaComissao, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 213787
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213788
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objVendPurExc = New ClassVendedorPurExc
    
            objVendPurExc.sProduto = sProduto
            objVendPurExc.dPercComissao = dPercComissao
            objVendPurExc.dMetaComissao = dMetaComissao
        
            objVendedorPur.colExcecoes.Add objVendPurExc
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213789
        
        
        Loop
        
    End If
    
    Set objVendedor.objInfoUsu = objVendedorPur

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedor_Le_Customizado = SUCESSO

    Exit Function

Erro_Vendedor_Le_Customizado:

    Vendedor_Le_Customizado = gErr

    Select Case gErr

        Case 213784
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213785 To 213789
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213790)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
