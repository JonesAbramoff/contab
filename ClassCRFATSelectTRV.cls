VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelectTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Vendedor_Le_Customizado(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim objVendedorTRV As New ClassVendedorTRV
Dim tVendedorTRV As typeVendedorTRV
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim tVNDComissaoTRV As typeVNDComissaoTRV
Dim tVNDReducaoTRV As typeVNDReducaoTRV
Dim tVNDRegiaoTRV As typeVNDRegiaoTRV
Dim objVNDComissaoTRV As ClassVNDComissaoTRV
Dim objVNDReducaoTRV As ClassVNDReducaoTRV
Dim objVNDRegiaoTRV As ClassVNDRegiaoTRV

On Error GoTo Erro_Vendedor_Le_Customizado

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195880
    Next

    'Alocação de espaço no buffer

    'Le a tabela VendedorTRV
    lErro = Comando_Executar(alComando(1), "SELECT Cargo, Superior, PercCallCenter FROM VendedorTRV WHERE Vendedor= ? ", tVendedorTRV.iCargo, tVendedorTRV.iSuperior, tVendedorTRV.dPercCallCenter, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195881

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195882

    objVendedorTRV.iCargo = tVendedorTRV.iCargo
    objVendedorTRV.iSuperior = tVendedorTRV.iSuperior
    objVendedorTRV.dPercCallCenter = tVendedorTRV.dPercCallCenter
    
    'Le a tabela VendedorTRVComissao
    lErro = Comando_Executar(alComando(2), "SELECT ValorDe, ValorAte, Moeda, PercComissao FROM VendedorTRVComissao WHERE Vendedor = ? ORDER BY Seq", tVNDComissaoTRV.dValorDe, tVNDComissaoTRV.dValorAte, tVNDComissaoTRV.iMoeda, tVNDComissaoTRV.dPercComissao, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195967

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195968

    Do While lErro = AD_SQL_SUCESSO

        Set objVNDComissaoTRV = New ClassVNDComissaoTRV

        objVNDComissaoTRV.dValorDe = tVNDComissaoTRV.dValorDe
        objVNDComissaoTRV.dValorAte = tVNDComissaoTRV.dValorAte
        objVNDComissaoTRV.iMoeda = tVNDComissaoTRV.iMoeda
        objVNDComissaoTRV.dPercComissao = tVNDComissaoTRV.dPercComissao
    
        objVendedorTRV.colVNDComissaoTRV.Add objVNDComissaoTRV
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195969
    
    
    Loop
    
    'Le a tabela VendedorTRVComiRed
    lErro = Comando_Executar(alComando(3), "SELECT ValorDe, ValorAte, Moeda, PercComissaoMax FROM VendedorTRVComiRed WHERE Vendedor = ? ORDER BY Seq", tVNDReducaoTRV.dValorDe, tVNDReducaoTRV.dValorAte, tVNDReducaoTRV.iMoeda, tVNDReducaoTRV.dPercComissaoMax, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195970

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195971

    Do While lErro = AD_SQL_SUCESSO

        Set objVNDReducaoTRV = New ClassVNDReducaoTRV

        objVNDReducaoTRV.dValorDe = tVNDReducaoTRV.dValorDe
        objVNDReducaoTRV.dValorAte = tVNDReducaoTRV.dValorAte
        objVNDReducaoTRV.iMoeda = tVNDReducaoTRV.iMoeda
        objVNDReducaoTRV.dPercComissaoMax = tVNDReducaoTRV.dPercComissaoMax
    
        objVendedorTRV.colVNDReducaoTRV.Add objVNDReducaoTRV
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195972
    
    Loop
    
    'Le a tabela VendedorTRVRegVenda
    lErro = Comando_Executar(alComando(4), "SELECT RegiaoVenda, PercComissao FROM VendedorTRVRegVenda WHERE Vendedor = ? ORDER BY Seq", tVNDRegiaoTRV.iRegiaoVenda, tVNDRegiaoTRV.dPercComissao, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195973

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195974

    Do While lErro = AD_SQL_SUCESSO

        Set objVNDRegiaoTRV = New ClassVNDRegiaoTRV

        objVNDRegiaoTRV.iRegiaoVenda = tVNDRegiaoTRV.iRegiaoVenda
        objVNDRegiaoTRV.dPercComissao = tVNDRegiaoTRV.dPercComissao
    
        objVendedorTRV.colVNDRegiaoTRV.Add objVNDRegiaoTRV
    
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195975
    
    Loop
    
    Set objVendedor.objInfoUsu = objVendedorTRV

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedor_Le_Customizado = SUCESSO

    Exit Function

Erro_Vendedor_Le_Customizado:

    Vendedor_Le_Customizado = gErr

    Select Case gErr

        Case 195880
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195881, 195882
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRV", gErr)

        Case 195967 To 195969
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVCOMISSAO", gErr)

        Case 195970 To 195972
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVCOMIRED", gErr)

        Case 195973 To 195975
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVREGVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195883)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CamposGenericos_ValidaExclusao_CargoVendedor() As Long

Dim lComando As Long
Dim lErro As Long
Dim iVendedor As Integer
Dim iCargo As Integer

On Error GoTo Erro_CamposGenericos_ValidaExclusao_CargoVendedor
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195987
    
    'Busca em relacionamentos com clientes, tipos de relacionamentos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Cargo, Vendedor FROM VendedorTRV WHERE Cargo NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Cargo <> 0", iCargo, iVendedor, CAMPOSGENERICOS_CARGO_VENDEDOR)
    If lErro <> AD_SQL_SUCESSO Then gError 195988
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195989
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 195990
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_CargoVendedor = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_CargoVendedor:

    CamposGenericos_ValidaExclusao_CargoVendedor = gErr
    
    Select Case gErr
    
        Case 195987
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 195988, 195989
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRV_CAMPOSGENERICOS", gErr)
        
        Case 195990
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CARGO", gErr, iCargo, iVendedor)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195991)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function

Public Function CamposGenericos_ValidaExclusao_DestinoViagem() As Long

Dim lComando As Long
Dim lErro As Long
Dim iDestino As Integer
Dim lNumero As Long
On Error GoTo Erro_CamposGenericos_ValidaExclusao_DestinoViagem
    
    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197039
    
    'Busca em relacionamentos com clientes, tipos de relacionamentos que não existam na tabela CamposGenericos Valores
    lErro = Comando_Executar(lComando, "SELECT Destino, Numero FROM TRVAcordoComissao, TRVAcordos WHERE TRVAcordos.NumIntDoc = NumIntAcordo AND Destino NOT IN (SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = ?) AND Destino <> 0", iDestino, lNumero, CAMPOSGENERICOS_DESTINO_VIAGEM)
    If lErro <> AD_SQL_SUCESSO Then gError 197040
    
    'Posiciona no primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197041
    
    'Se encontrou algum registro => erro de inconsistência
    If lErro = AD_SQL_SUCESSO Then gError 197042
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    CamposGenericos_ValidaExclusao_DestinoViagem = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_ValidaExclusao_DestinoViagem:

    CamposGenericos_ValidaExclusao_DestinoViagem = gErr
    
    Select Case gErr
    
        Case 197039
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 197040, 197041
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVACORDOCOMISSAO_CGENERICO", gErr)
        
        Case 197042
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DESTINOVIAGEM", gErr, iDestino, lNumero)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197043)
    
    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

End Function


Function Vendedor_Le_Customizado1(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim objVendedorTRV As New ClassVendedorTRV
Dim tVendedorTRV As typeVendedorTRV
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim tVNDComissaoTRV As typeVNDComissaoTRV
Dim tVNDReducaoTRV As typeVNDReducaoTRV
Dim tVNDRegiaoTRV As typeVNDRegiaoTRV
Dim objVNDComissaoTRV As ClassVNDComissaoTRV
Dim objVNDReducaoTRV As ClassVNDReducaoTRV
Dim objVNDRegiaoTRV As ClassVNDRegiaoTRV

On Error GoTo Erro_Vendedor_Le_Customizado1

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197347
    Next

    'Alocação de espaço no buffer

    'Le a tabela VendedorTRV
    lErro = Comando_Executar(alComando(1), "SELECT Cargo, Superior, PercCallCenter FROM VendedorTRV WHERE Vendedor= ? ", tVendedorTRV.iCargo, tVendedorTRV.iSuperior, tVendedorTRV.dPercCallCenter, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197348

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197349

    If lErro = AD_SQL_SUCESSO Then

        objVendedorTRV.iCargo = tVendedorTRV.iCargo
        objVendedorTRV.iSuperior = tVendedorTRV.iSuperior
        objVendedorTRV.dPercCallCenter = tVendedorTRV.dPercCallCenter
        
        'Le a tabela VendedorTRVComissao
        lErro = Comando_Executar(alComando(2), "SELECT ValorDe, ValorAte, Moeda, PercComissao FROM VendedorTRVComissao WHERE Vendedor = ? ORDER BY ValorDe DESC", tVNDComissaoTRV.dValorDe, tVNDComissaoTRV.dValorAte, tVNDComissaoTRV.iMoeda, tVNDComissaoTRV.dPercComissao, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 197350
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197351
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objVNDComissaoTRV = New ClassVNDComissaoTRV
    
            objVNDComissaoTRV.dValorDe = tVNDComissaoTRV.dValorDe
            objVNDComissaoTRV.dValorAte = tVNDComissaoTRV.dValorAte
            objVNDComissaoTRV.iMoeda = tVNDComissaoTRV.iMoeda
            objVNDComissaoTRV.dPercComissao = tVNDComissaoTRV.dPercComissao
        
            objVendedorTRV.colVNDComissaoTRV.Add objVNDComissaoTRV
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197352
        
        
        Loop
        
        'Le a tabela VendedorTRVComiRed
        lErro = Comando_Executar(alComando(3), "SELECT ValorDe, ValorAte, Moeda, PercComissaoMax FROM VendedorTRVComiRed WHERE Vendedor = ? ORDER BY Seq", tVNDReducaoTRV.dValorDe, tVNDReducaoTRV.dValorAte, tVNDReducaoTRV.iMoeda, tVNDReducaoTRV.dPercComissaoMax, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 197353
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197354
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objVNDReducaoTRV = New ClassVNDReducaoTRV
    
            objVNDReducaoTRV.dValorDe = tVNDReducaoTRV.dValorDe
            objVNDReducaoTRV.dValorAte = tVNDReducaoTRV.dValorAte
            objVNDReducaoTRV.iMoeda = tVNDReducaoTRV.iMoeda
            objVNDReducaoTRV.dPercComissaoMax = tVNDReducaoTRV.dPercComissaoMax
        
            objVendedorTRV.colVNDReducaoTRV.Add objVNDReducaoTRV
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197355
        
        Loop
        
        'Le a tabela VendedorTRVRegVenda
        lErro = Comando_Executar(alComando(4), "SELECT RegiaoVenda, PercComissao FROM VendedorTRVRegVenda WHERE Vendedor = ? ORDER BY Seq", tVNDRegiaoTRV.iRegiaoVenda, tVNDRegiaoTRV.dPercComissao, objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 197356
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197357
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objVNDRegiaoTRV = New ClassVNDRegiaoTRV
    
            objVNDRegiaoTRV.iRegiaoVenda = tVNDRegiaoTRV.iRegiaoVenda
            objVNDRegiaoTRV.dPercComissao = tVNDRegiaoTRV.dPercComissao
            
            objVendedorTRV.colVNDRegiaoTRV.Add objVNDRegiaoTRV
        
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197358
        
        Loop
        
        Set objVendedor.objInfoUsu = objVendedorTRV

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedor_Le_Customizado1 = SUCESSO

    Exit Function

Erro_Vendedor_Le_Customizado1:

    Vendedor_Le_Customizado1 = gErr

    Select Case gErr

        Case 197347
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197348, 197349
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRV", gErr)

        Case 197350 To 197352
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVCOMISSAO", gErr)

        Case 197353 To 197355
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVCOMIRED", gErr)

        Case 197356 To 197358
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRVREGVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197359)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

