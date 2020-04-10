VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesFIS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Repreenche_Linha_LivroFiscal(objNFiscal As ClassNFiscal, objLivRegES As ClassLivRegES, objTela As Object) As Long
'Preenche o Grid de Lancamentos dos Livros de Reg de Entrada e de Saída

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Repreenche_Linha_LivroFiscal
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REPREENCHER_LINHA_LIVROFISCAL", X)
    If lErro <> SUCESSO Then gError 69971
    
    Repreenche_Linha_LivroFiscal = X.Repreenche_Linha_LivroFiscal(objNFiscal, objLivRegES, objTela)
    
    Exit Function
    
Erro_Repreenche_Linha_LivroFiscal:

    Repreenche_Linha_LivroFiscal = gErr
    
    Select Case gErr

        Case 69971
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145810)
        
    End Select
    
    Exit Function
    
End Function


Function FISConfig_Le(objFis As ClassFIS) As Long
'Lê a alcada a partir do código do usuário.
'Devolve os dados em objLoja.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FisConfig_Le
    
    
    lErro = Rotina_ObterObjetoFuncao("FIS_FISCONFIG_LE", X)
    If lErro <> SUCESSO Then gError 80168
    
    FISConfig_Le = X.FISConfig_Le(objFis)

    Exit Function
    
Erro_FisConfig_Le:
    
    FISConfig_Le = gErr
    
    Select Case gErr

        Case 80168
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145811)
        
    End Select
    
    Exit Function

End Function

Function Declan_Obter_Estoque(dtData As Date, iFilialEmpresa As Integer, dValorEstoque As Double, lComando As Long) As Long
'Le o valor de estoque na data passada para a Filial Passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Declan_Obter_Estoque
    
    lErro = Rotina_ObterObjetoFuncao("FIS_DECLAN_OBTER_ESTOQUE", X)
    If lErro <> SUCESSO Then gError 71389
    
    Declan_Obter_Estoque = X.Declan_Obter_Estoque(dtData, iFilialEmpresa, dValorEstoque, lComando)

    Exit Function
    
Erro_Declan_Obter_Estoque:
    
    Declan_Obter_Estoque = gErr
    
    Select Case gErr

        Case 71389
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145812)
        
    End Select
    
    Exit Function

End Function

Function LivRegES_Le_Saidas(dValorSaidasICMS As Double, dValorSaidasIPI As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer, lComando As Long) As Long
'Lê o valor das Saidas de ICMS e IpI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivRegES_Le_Saidas
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVREGES_LE_SAIDAS", X)
    If lErro <> SUCESSO Then gError 71391
    
    LivRegES_Le_Saidas = X.LivRegES_Le_Saidas(dValorSaidasICMS, dValorSaidasIPI, dtDataDe, dtDataAte, iFilialEmpresa, lComando)

    Exit Function
    
Erro_LivRegES_Le_Saidas:
    
    LivRegES_Le_Saidas = gErr
    
    Select Case gErr

        Case 71391
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145813)
        
    End Select
    
    Exit Function

End Function

Function LivRegES_Le_Entradas(dValorEntradasICMS As Double, dValorEntradasIPI As Double, dtDataDe As Date, dtDataAte As Date, iFilialEmpresa As Integer, lComando As Long) As Long
'Lê o valor das Entradas de ICMS e IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivRegES_Le_Entradas
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVREGES_LE_ENTRADA", X)
    If lErro <> SUCESSO Then gError 71390
    
    LivRegES_Le_Entradas = X.LivRegES_Le_Entradas(dValorEntradasICMS, dValorEntradasIPI, dtDataDe, dtDataAte, iFilialEmpresa, lComando)

    Exit Function
    
Erro_LivRegES_Le_Entradas:
    
    LivRegES_Le_Entradas = gErr
    
    Select Case gErr

        Case 71390
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145814)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoICMS_Le_SaldoCredor(objApuracao As ClassRegApuracao, lComando As Long) As Long
'Le o Saldo credor do periodo anterior
'função de Relatorio

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoICMS_Le_SaldoCredor
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOICMS_LE_SALDOCREDOR", X)
    If lErro <> SUCESSO Then gError 71392
    
    ApuracaoICMS_Le_SaldoCredor = X.ApuracaoICMS_Le_SaldoCredor(objApuracao, lComando)

    Exit Function
    
Erro_ApuracaoICMS_Le_SaldoCredor:
    
    ApuracaoICMS_Le_SaldoCredor = gErr
    
    Select Case gErr

        Case 71392
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145815)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoIPI_Le_SaldoCredor(objApuracao As ClassRegApuracao, lComando As Long) As Long
'Le o Saldo credor do periodo anterior
'função de Relatorio

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoIPI_Le_SaldoCredor
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOIPI_LE_SALDOCREDOR", X)
    If lErro <> SUCESSO Then gError 71393
    
    ApuracaoIPI_Le_SaldoCredor = X.ApuracaoIPI_Le_SaldoCredor(objApuracao, lComando)

    Exit Function
    
Erro_ApuracaoIPI_Le_SaldoCredor:
    
    ApuracaoIPI_Le_SaldoCredor = gErr
    
    Select Case gErr

        Case 71393
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145816)
        
    End Select
    
    Exit Function

End Function

Function LivrosFilial_Le(objLivroFilial As ClassLivrosFilial) As Long
'Lê dados do Livro do Tipo passado da Filial Empresa também passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFilial_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFILIAL_LE", X)
    If lErro <> SUCESSO Then gError 71425
    
    LivrosFilial_Le = X.LivrosFilial_Le(objLivroFilial)

    Exit Function
    
Erro_LivrosFilial_Le:
    
    LivrosFilial_Le = gErr
    
    Select Case gErr

        Case 71425
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145817)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoICMS_Le_UltimaFechada(objApuracao As ClassRegApuracao) As Long
'Lê dados da última Apuração ICMS Fechada (ordem de data final)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoICMS_Le_UltimaFechada
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOICMS_LE_ULTIMAFECHADA", X)
    If lErro <> SUCESSO Then gError 71426
    
    ApuracaoICMS_Le_UltimaFechada = X.ApuracaoICMS_Le_UltimaFechada(objApuracao)

    Exit Function
    
Erro_ApuracaoICMS_Le_UltimaFechada:
    
    ApuracaoICMS_Le_UltimaFechada = gErr
    
    Select Case gErr

        Case 71426
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145818)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoICMS_Le(objApuracao As ClassRegApuracao) As Long
'Lê dados da Apuração ICMS a partir da data Inicial, data final e FilialEmpresa passados em objApuracao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoICMS_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOICMS_LE", X)
    If lErro <> SUCESSO Then gError 71427
    
    ApuracaoICMS_Le = X.ApuracaoICMS_Le(objApuracao)

    Exit Function
    
Erro_ApuracaoICMS_Le:
    
    ApuracaoICMS_Le = gErr
    
    Select Case gErr

        Case 71427
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145819)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoICMS_Exclui(objApuracao As ClassRegApuracao) As Long
'Exclui um Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoICMS_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOICMS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71428
    
    RegApuracaoICMS_Exclui = X.RegApuracaoICMS_Exclui(objApuracao)

    Exit Function
    
Erro_RegApuracaoICMS_Exclui:
    
    RegApuracaoICMS_Exclui = gErr
    
    Select Case gErr

        Case 71428
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145820)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoICMS_Grava(objApuracao As ClassRegApuracao) As Long
'Grava um Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoICMS_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOICMS_GRAVA", X)
    If lErro <> SUCESSO Then gError 71429
    
    RegApuracaoICMS_Grava = X.RegApuracaoICMS_Grava(objApuracao)

    Exit Function
    
Erro_RegApuracaoICMS_Grava:
    
    RegApuracaoICMS_Grava = gErr
    
    Select Case gErr

        Case 71429
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145821)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoIPIItens_Le(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Lê o Item de Registro de Apuração IPI a partir do Tipo, da descrição e da data passados como parametros no objRegApuracaoItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoIPIItens_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOIPIITENS_LE", X)
    If lErro <> SUCESSO Then gError 71430
    
    RegApuracaoIPIItens_Le = X.RegApuracaoIPIItens_Le(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoIPIItens_Le:
    
    RegApuracaoIPIItens_Le = gErr
    
    Select Case gErr

        Case 71430
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145822)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoIPIItem_Exclui(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Exclui um Tipo de Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoIPIItem_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOIPIITEM_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71431
    
    RegApuracaoIPIItem_Exclui = X.RegApuracaoIPIItem_Exclui(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoIPIItem_Exclui:
    
    RegApuracaoIPIItem_Exclui = gErr
    
    Select Case gErr

        Case 71431
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145823)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoIPIItem_Grava(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Grava um Item de Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoIPIItem_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOIPIITEM_GRAVA", X)
    If lErro <> SUCESSO Then gError 71432
    
    RegApuracaoIPIItem_Grava = X.RegApuracaoIPIItem_Grava(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoIPIItem_Grava:
    
    RegApuracaoIPIItem_Grava = gErr
    
    Select Case gErr

        Case 71432
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145824)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoIPI_Le(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Lê o tipo de apuração a partir do código passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoIPI_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOIPI_LE", X)
    If lErro <> SUCESSO Then gError 71433
    
    TipoRegApuracaoIPI_Le = X.TipoRegApuracaoIPI_Le(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoIPI_Le:
    
    TipoRegApuracaoIPI_Le = gErr
    
    Select Case gErr

        Case 71433
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145825)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoICMSItens_Le(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Lê o Item de Registro de Apuração ICMS a partir do Tipo, da descrição e da data passados como parametros no objRegApuracaoItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoICMSItens_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOICMSITENS_LE", X)
    If lErro <> SUCESSO Then gError 71434
    
    RegApuracaoICMSItens_Le = X.RegApuracaoICMSItens_Le(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoICMSItens_Le:
    
    RegApuracaoICMSItens_Le = gErr
    
    Select Case gErr

        Case 71434
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145826)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoICMSItem_Exclui(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Exclui um Tipo de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoICMSItem_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOICMSITENS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71435
    
    RegApuracaoICMSItem_Exclui = X.RegApuracaoICMSItem_Exclui(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoICMSItem_Exclui:
    
    RegApuracaoICMSItem_Exclui = gErr
    
    Select Case gErr

        Case 71435
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145827)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoICMSItem_Grava(objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Grava um Item de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoICMSItem_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOICMSITEM_GRAVA", X)
    If lErro <> SUCESSO Then gError 71436
    
    RegApuracaoICMSItem_Grava = X.RegApuracaoICMSItem_Grava(objRegApuracaoItem)

    Exit Function
    
Erro_RegApuracaoICMSItem_Grava:
    
    RegApuracaoICMSItem_Grava = gErr
    
    Select Case gErr

        Case 71436
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145828)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoICMS_Le(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Lê o tipo de apuração a partir do código passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoICMS_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOICMS_LE", X)
    If lErro <> SUCESSO Then gError 71437
    
    TipoRegApuracaoICMS_Le = X.TipoRegApuracaoICMS_Le(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoICMS_Le:
    
    TipoRegApuracaoICMS_Le = gErr
    
    Select Case gErr

        Case 71437
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145829)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoIPI_Le_UltimaFechada(objApuracao As ClassRegApuracao) As Long
'Lê dados da última Apuração IPI Fechada (ordem de data final)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoIPI_Le_UltimaFechada
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOIPI_LE_ULTIMAFECHADA", X)
    If lErro <> SUCESSO Then gError 71438
    
    ApuracaoIPI_Le_UltimaFechada = X.ApuracaoIPI_Le_UltimaFechada(objApuracao)

    Exit Function
    
Erro_ApuracaoIPI_Le_UltimaFechada:
    
    ApuracaoIPI_Le_UltimaFechada = gErr
    
    Select Case gErr

        Case 71438
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145830)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoIPI_Le(objApuracao As ClassRegApuracao) As Long
'Lê dados da Apuração IPI a partir da data Inicial, data final e FilialEmpresa passados em objApuracao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoIPI_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOIPI_LE", X)
    If lErro <> SUCESSO Then gError 71439
    
    ApuracaoIPI_Le = X.ApuracaoIPI_Le(objApuracao)

    Exit Function
    
Erro_ApuracaoIPI_Le:
    
    ApuracaoIPI_Le = gErr
    
    Select Case gErr

        Case 71439
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145831)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoIPI_Exclui(objApuracao As ClassRegApuracao) As Long
'Exclui um Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoIPI_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOIPI_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71440
    
    RegApuracaoIPI_Exclui = X.RegApuracaoIPI_Exclui(objApuracao)

    Exit Function
    
Erro_RegApuracaoIPI_Exclui:
    
    RegApuracaoIPI_Exclui = gErr
    
    Select Case gErr

        Case 71440
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145832)
        
    End Select
    
    Exit Function

End Function

Function RegApuracaoIPI_Grava(objApuracao As ClassRegApuracao) As Long
'Grava um Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegApuracaoIPI_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGAPURACAOIPI_GRAVA", X)
    If lErro <> SUCESSO Then gError 71441
    
    RegApuracaoIPI_Grava = X.RegApuracaoIPI_Grava(objApuracao)

    Exit Function
    
Erro_RegApuracaoIPI_Grava:
    
    RegApuracaoIPI_Grava = gErr
    
    Select Case gErr

        Case 71441
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145833)
        
    End Select
    
    Exit Function

End Function

Function GNRICMS_Le(objGNRICMS As ClassGNRICMS) As Long
'Lê dados da Guia de Recolhimento de ICMS a partir do Código passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GNRICMS_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GNRICMS_LE", X)
    If lErro <> SUCESSO Then gError 71442
    
    GNRICMS_Le = X.GNRICMS_Le(objGNRICMS)

    Exit Function
    
Erro_GNRICMS_Le:
    
    GNRICMS_Le = gErr
    
    Select Case gErr

        Case 71442
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145834)
        
    End Select
    
    Exit Function

End Function

Function GNRICMS_Exclui(objGNRICMS As ClassGNRICMS) As Long
'Exclui uma Guia de Recolhimento de ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GNRICMS_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GNRICMS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71443
    
    GNRICMS_Exclui = X.GNRICMS_Exclui(objGNRICMS)

    Exit Function
    
Erro_GNRICMS_Exclui:
    
    GNRICMS_Exclui = gErr
    
    Select Case gErr

        Case 71443
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145835)
        
    End Select
    
    Exit Function

End Function

Function GNRICMS_Grava(objGNRICMS As ClassGNRICMS) As Long
'Grava uma Guia de ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GNRICMS_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GNRICMS_GRAVA", X)
    If lErro <> SUCESSO Then gError 71444
    
    GNRICMS_Grava = X.GNRICMS_Grava(objGNRICMS)

    Exit Function
    
Erro_GNRICMS_Grava:
    
    GNRICMS_Grava = gErr
    
    Select Case gErr

        Case 71444
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145836)
        
    End Select
    
    Exit Function

End Function

Function GNRICMS_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GNRICMS_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GNRICMS_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71445
    
    GNRICMS_Codigo_Automatico = X.GNRICMS_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_GNRICMS_Codigo_Automatico:
    
    GNRICMS_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71445
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145837)
        
    End Select
    
    Exit Function

End Function

Function GeracaoArqICMS_Limpar_TabelasAuxiliares() As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoArqICMS_Limpar_TabelasAuxiliares
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GERACAOARQICMS_LIMPAR_TABELASAUXILIA", X)
    If lErro <> SUCESSO Then gError 71446
    
    GeracaoArqICMS_Limpar_TabelasAuxiliares = X.GeracaoArqICMS_Limpar_TabelasAuxiliares()

    Exit Function
    
Erro_GeracaoArqICMS_Limpar_TabelasAuxiliares:
    
    GeracaoArqICMS_Limpar_TabelasAuxiliares = gErr
    
    Select Case gErr

        Case 71446
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145838)
        
    End Select
    
    Exit Function

End Function

Function GeracaoArqICMS_Grava_FIS(objInfoArqICMS As ClassInfoArqICMS, iTipo As Integer, lComando As Long) As Long
'Grava o Tipo no BD
'Tem que ser chamado  em transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoArqICMS_Grava_FIS
    
    lErro = Rotina_ObterObjetoFuncao("FIS_GERACAOARQICMS_GRAVA_FIS", X)
    If lErro <> SUCESSO Then gError 71447
    
    GeracaoArqICMS_Grava_FIS = X.GeracaoArqICMS_Grava_FIS(objInfoArqICMS, iTipo, lComando)

    Exit Function
    
Erro_GeracaoArqICMS_Grava_FIS:
    
    GeracaoArqICMS_Grava_FIS = gErr
    
    Select Case gErr

        Case 71447
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145839)
        
    End Select
    
    Exit Function

End Function

Function InfoArqICMS_Grava(objInfoArqICMS As ClassInfoArqICMS) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InfoArqICMS_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_INFOARQICMS_GRAVA", X)
    If lErro <> SUCESSO Then gError 71448
    
    InfoArqICMS_Grava = X.InfoArqICMS_Grava(objInfoArqICMS)

    Exit Function
    
Erro_InfoArqICMS_Grava:
    
    InfoArqICMS_Grava = gErr
    
    Select Case gErr

        Case 71448
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145840)
        
    End Select
    
    Exit Function

End Function

Function InfoArqICMS_Le(objInfoArqICMS As ClassInfoArqICMS) As Long
'Le o Arquivo se para as datas Iniciais e Finais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InfoArqICMS_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_INFOARQICMS_LE", X)
    If lErro <> SUCESSO Then gError 71449
    
    InfoArqICMS_Le = X.InfoArqICMS_Le(objInfoArqICMS)

    Exit Function
    
Erro_InfoArqICMS_Le:
    
    InfoArqICMS_Le = gErr
    
    Select Case gErr

        Case 71449
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145841)
        
    End Select
    
    Exit Function

End Function

Function InfoArqICMS_Exclui(objInfoArqICMS As ClassInfoArqICMS) As Long
'Exclui o Arquivo de ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InfoArqICMS_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_INFOARQICMS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71450
    
    InfoArqICMS_Exclui = X.InfoArqICMS_Exclui(objInfoArqICMS)

    Exit Function
    
Erro_InfoArqICMS_Exclui:
    
    InfoArqICMS_Exclui = gErr
    
    Select Case gErr

        Case 71450
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145842)
        
    End Select
    
    Exit Function

End Function

Function ConfiguraFIS_GravarTrans(objFis As ClassFIS) As Long
'Efetua a gravação dos dados com a transação já aberta em função _
 anterior

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ConfiguraFIS_GravarTrans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_CONFIGURAFIS_GRAVARTRANS", X)
    If lErro <> SUCESSO Then gError 71451
    
    ConfiguraFIS_GravarTrans = X.ConfiguraFIS_GravarTrans(objFis)

    Exit Function
    
Erro_ConfiguraFIS_GravarTrans:
    
    ConfiguraFIS_GravarTrans = gErr
    
    Select Case gErr

        Case 71451
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145843)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoICMS_Exclui(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Exclui um Tipo de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoICMS_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOICMS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71452
    
    TipoRegApuracaoICMS_Exclui = X.TipoRegApuracaoICMS_Exclui(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoICMS_Exclui:
    
    TipoRegApuracaoICMS_Exclui = gErr
    
    Select Case gErr

        Case 71452
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145844)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoICMS_Grava(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Grava um tipo de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoICMS_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOICMS_GRAVA", X)
    If lErro <> SUCESSO Then gError 71453
    
    TipoRegApuracaoICMS_Grava = X.TipoRegApuracaoICMS_Grava(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoICMS_Grava:
    
    TipoRegApuracaoICMS_Grava = gErr
    
    Select Case gErr

        Case 71453
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145845)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoICMS_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoICMS_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOICMS_CODIGO_AUTOMATIC", X)
    If lErro <> SUCESSO Then gError 71454
    
    TipoRegApuracaoICMS_Codigo_Automatico = X.TipoRegApuracaoICMS_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_TipoRegApuracaoICMS_Codigo_Automatico:
    
    TipoRegApuracaoICMS_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71454
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145846)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoIPI_Exclui(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Exclui um Tipo de Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoIPI_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOIPI_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71455
    
    TipoRegApuracaoIPI_Exclui = X.TipoRegApuracaoIPI_Exclui(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoIPI_Exclui:
    
    TipoRegApuracaoIPI_Exclui = gErr
    
    Select Case gErr

        Case 71455
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145847)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoIPI_Grava(objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Grava um tipo de Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoIPI_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOIPI_GRAVA", X)
    If lErro <> SUCESSO Then gError 71456
    
    TipoRegApuracaoIPI_Grava = X.TipoRegApuracaoIPI_Grava(objTipoRegApuracao)

    Exit Function
    
Erro_TipoRegApuracaoIPI_Grava:
    
    TipoRegApuracaoIPI_Grava = gErr
    
    Select Case gErr

        Case 71456
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145848)
        
    End Select
    
    Exit Function

End Function

Function TipoRegApuracaoIPI_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoRegApuracaoIPI_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TIPOREGAPURACAOIPI_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71457
    
    TipoRegApuracaoIPI_Codigo_Automatico = X.TipoRegApuracaoIPI_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_TipoRegApuracaoIPI_Codigo_Automatico:
    
    TipoRegApuracaoIPI_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71457
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145849)
        
    End Select
    
    Exit Function

End Function

Function LivroFilial_Exclui(objLivroFilial As ClassLivrosFilial) As Long
'Exclui um Livro Fiscal da Filial Empresa passada em objLivroFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivroFilial_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROFILIAL_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71458
    
    LivroFilial_Exclui = X.LivroFilial_Exclui(objLivroFilial)

    Exit Function
    
Erro_LivroFilial_Exclui:
    
    LivroFilial_Exclui = gErr
    
    Select Case gErr

        Case 71458
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145850)
        
    End Select
    
    Exit Function

End Function

Function LivroFilial_Grava(objLivroFilial As ClassLivrosFilial) As Long
'Grava um Livro Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivroFilial_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROFILIAL_GRAVA", X)
    If lErro <> SUCESSO Then gError 71459
    
    LivroFilial_Grava = X.LivroFilial_Grava(objLivroFilial)

    Exit Function
    
Erro_LivroFilial_Grava:
    
    LivroFilial_Grava = gErr
    
    Select Case gErr

        Case 71459
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145851)
        
    End Select
    
    Exit Function

End Function

Function Tributos_Le(colTributos As Collection) As Long
'Lê Tributos que possuem Livro Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Tributos_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_TRIBUTOS_LE", X)
    If lErro <> SUCESSO Then gError 71460
    
    Tributos_Le = X.Tributos_Le(colTributos)

    Exit Function
    
Erro_Tributos_Le:
    
    Tributos_Le = gErr
    
    Select Case gErr

        Case 71460
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145852)
        
    End Select
    
    Exit Function

End Function

Function Periodicidades_Le(colPeriodo As Collection) As Long
'Lê as Periodicidades dos Livros Fiscais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Periodicidades_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_PERIODICIDADES_LE", X)
    If lErro <> SUCESSO Then gError 71461
    
    Periodicidades_Le = X.Periodicidades_Le(colPeriodo)

    Exit Function
    
Erro_Periodicidades_Le:
    
    Periodicidades_Le = gErr
    
    Select Case gErr

        Case 71461
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145853)
        
    End Select
    
    Exit Function

End Function

Function LivrosFiscais_Le(iTributo As Integer, colLivroFiscal As Collection) As Long
'Lê Livros Fiscais associados ao Tributo passado por iTributo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFiscais_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFISCAIS_LE", X)
    If lErro <> SUCESSO Then gError 71462
    
    LivrosFiscais_Le = X.LivrosFiscais_Le(iTributo, colLivroFiscal)

    Exit Function
    
Erro_LivrosFiscais_Le:
    
    LivrosFiscais_Le = gErr
    
    Select Case gErr

        Case 71462
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145854)
        
    End Select
    
    Exit Function

End Function

Function LivroFiscal_Le_Codigo(objLivroFiscal As ClassLivrosFiscais) As Long
'Lê dados do Livro Fiscal a partir do código passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivroFiscal_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROFISCAL_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 71463
    
    LivroFiscal_Le_Codigo = X.LivroFiscal_Le_Codigo(objLivroFiscal)

    Exit Function
    
Erro_LivroFiscal_Le_Codigo:
    
    LivroFiscal_Le_Codigo = gErr
    
    Select Case gErr

        Case 71463
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145855)
        
    End Select
    
    Exit Function

End Function

Function LivrosFechados_Le_UltimaData(objLivroFechado As ClassLivrosFechados) As Long
'Lê dados do último Livro Fiscal Fechado do Livro de código e Filial Empresea passados em objLivroFechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFechados_Le_UltimaData
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFECHADOS_LE_ULTIMADATA", X)
    If lErro <> SUCESSO Then gError 71464
    
    LivrosFechados_Le_UltimaData = X.LivrosFechados_Le_UltimaData(objLivroFechado)

    Exit Function
    
Erro_LivrosFechados_Le_UltimaData:
    
    LivrosFechados_Le_UltimaData = gErr
    
    Select Case gErr

        Case 71464
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145856)
        
    End Select
    
    Exit Function

End Function

Function LivRegES_Le_NotaFiscal(objLivRegES As ClassLivRegES) As Long
'Le o Livro correspondente a Nota Fiscal passada
'chave lNumIntDoc de Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivRegES_Le_NotaFiscal
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVREGES_LE_NOTAFISCAL", X)
    If lErro <> SUCESSO Then gError 71465
    
    LivRegES_Le_NotaFiscal = X.LivRegES_Le_NotaFiscal(objLivRegES)

    Exit Function
    
Erro_LivRegES_Le_NotaFiscal:
    
    LivRegES_Le_NotaFiscal = gErr
    
    Select Case gErr

        Case 71465
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145857)
        
    End Select
    
    Exit Function

End Function

Function LivRegESLinha_Le(objLivRegES As ClassLivRegES) As Long
'Le as Linhas para o Livro passado em objLivRegES.lNumIntDoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivRegESLinha_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVREGESLINHA_LE", X)
    If lErro <> SUCESSO Then gError 71466
    
    LivRegESLinha_Le = X.LivRegESLinha_Le(objLivRegES)

    Exit Function
    
Erro_LivRegESLinha_Le:
    
    LivRegESLinha_Le = gErr
    
    Select Case gErr

        Case 71466
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145858)
        
    End Select
    
    Exit Function

End Function

Function LivRegESItemNF_Le(objLivRegES As ClassLivRegES) As Long
'Le os Itens para o Livro passado em objLivRegES.lNumIntDoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivRegESItemNF_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVREGESITEMNF_LE", X)
    If lErro <> SUCESSO Then gError 71467
    
    LivRegESItemNF_Le = X.LivRegESItemNF_Le(objLivRegES)

    Exit Function
    
Erro_LivRegESItemNF_Le:
    
    LivRegESItemNF_Le = gErr
    
    Select Case gErr

        Case 71467
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145859)
        
    End Select
    
    Exit Function

End Function

Function EdicaoRegES_Grava(objLivRegES As ClassLivRegES) As Long
'Grava as Linhas do Livro e Altera os Itens da Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoRegES_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FIS_EDICAOREGES_GRAVA", X)
    If lErro <> SUCESSO Then gError 71468
    
    EdicaoRegES_Grava = X.EdicaoRegES_Grava(objLivRegES)

    Exit Function
    
Erro_EdicaoRegES_Grava:
    
    EdicaoRegES_Grava = gErr
    
    Select Case gErr

        Case 71468
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145860)
        
    End Select
    
    Exit Function

End Function

Function EdicaoRegES_Grava_Linha_Trans(objLivRegES As ClassLivRegES) As Long
'Grava as Linhas do Livro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoRegES_Grava_Linha_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_EDICAOREGES_GRAVA_LINHA_TRANS", X)
    If lErro <> SUCESSO Then gError 71469
    
    EdicaoRegES_Grava_Linha_Trans = X.EdicaoRegES_Grava_Linha_Trans(objLivRegES)

    Exit Function
    
Erro_EdicaoRegES_Grava_Linha_Trans:
    
    EdicaoRegES_Grava_Linha_Trans = gErr
    
    Select Case gErr

        Case 71469
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145861)
        
    End Select
    
    Exit Function

End Function

Function EdicaoRegES_Grava_AlteraStatus_Trans(objLivRegES As ClassLivRegES) As Long
'Altera o Status do Livro Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoRegES_Grava_AlteraStatus_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_EDICAOREGES_GRAVA_ALTERASTATUS_TRANS", X)
    If lErro <> SUCESSO Then gError 71470
    
    EdicaoRegES_Grava_AlteraStatus_Trans = X.EdicaoRegES_Grava_AlteraStatus_Trans(objLivRegES)

    Exit Function
    
Erro_EdicaoRegES_Grava_AlteraStatus_Trans:
    
    EdicaoRegES_Grava_AlteraStatus_Trans = gErr
    
    Select Case gErr

        Case 71470
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145862)
        
    End Select
    
    Exit Function

End Function

Function EdicaoRegES_Grava_Item_Trans(objLivRegES As ClassLivRegES) As Long
'Altera os Itens da Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoRegES_Grava_Item_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_EDICAOREGES_GRAVA_ITEM_TRANS", X)
    If lErro <> SUCESSO Then gError 71471
    
    EdicaoRegES_Grava_Item_Trans = X.EdicaoRegES_Grava_Item_Trans(objLivRegES)

    Exit Function
    
Erro_EdicaoRegES_Grava_Item_Trans:
    
    EdicaoRegES_Grava_Item_Trans = gErr
    
    Select Case gErr

        Case 71471
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145863)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Le_Data(objRegInventario As ClassRegInventario) As Long
'Lê o Registro de Inventário a partir da Data e Filial Empresa passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Le_Data
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_LE_DATA", X)
    If lErro <> SUCESSO Then gError 71472
    
    RegInventario_Le_Data = X.RegInventario_Le_Data(objRegInventario)

    Exit Function
    
Erro_RegInventario_Le_Data:
    
    RegInventario_Le_Data = gErr
    
    Select Case gErr

        Case 71472
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145864)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Le(objRegInventario As ClassRegInventario) As Long
'Lê o Registro de Inventário a partir da Data, Produto e Filial Empresa passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_LE", X)
    If lErro <> SUCESSO Then gError 71473
    
    RegInventario_Le = X.RegInventario_Le(objRegInventario)

    Exit Function
    
Erro_RegInventario_Le:
    
    RegInventario_Le = gErr
    
    Select Case gErr

        Case 71473
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145865)
        
    End Select
    
    Exit Function

End Function

Function RegInventarioAlmox_Le(objRegInventarioAlmox As ClassRegInventarioAlmox) As Long
'Lê o Registro de Inventário a partir da Data, Produto e Filial Empresa passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventarioAlmox_Le
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIOALMOX_LE", X)
    If lErro <> SUCESSO Then gError 71474
    
    RegInventarioAlmox_Le = X.RegInventarioAlmox_Le(objRegInventarioAlmox)

    Exit Function
    
Erro_RegInventarioAlmox_Le:
    
    RegInventarioAlmox_Le = gErr
    
    Select Case gErr

        Case 71474
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145866)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Exclui(objRegInventario As ClassRegInventario) As Long
'Exclui um Registro de Inventário

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71475
    
    RegInventario_Exclui = X.RegInventario_Exclui(objRegInventario)

    Exit Function
    
Erro_RegInventario_Exclui:
    
    RegInventario_Exclui = gErr
    
    Select Case gErr

        Case 71475
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145867)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Atualiza(objRegInventario As ClassRegInventario) As Long
'Atualiza dados de RegInventário

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71476
    
    RegInventario_Atualiza = X.RegInventario_Atualiza(objRegInventario)

    Exit Function
    
Erro_RegInventario_Atualiza:
    
    RegInventario_Atualiza = gErr
    
    Select Case gErr

        Case 71476
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145868)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_IntervaloData_Critica(objLivroFilial As ClassLivrosFilial, dtData As Date) As Long
'Verifica a data passada está no intervalo de datas do Livro Fiscal de Registro de inventário aberto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_IntervaloData_Critica
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_INTERVALODATA_CRITICA", X)
    If lErro <> SUCESSO Then gError 71477
    
    RegInventario_IntervaloData_Critica = X.RegInventario_IntervaloData_Critica(objLivroFilial, dtData)

    Exit Function
    
Erro_RegInventario_IntervaloData_Critica:
    
    RegInventario_IntervaloData_Critica = gErr
    
    Select Case gErr

        Case 71477
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145869)
        
    End Select
    
    Exit Function

End Function

Function RegInventarioTodos_Exclui(objRegInventario As ClassRegInventario) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventarioTodos_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIOTODOS_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71478
    
    RegInventarioTodos_Exclui = X.RegInventarioTodos_Exclui(objRegInventario)

    Exit Function
    
Erro_RegInventarioTodos_Exclui:
    
    RegInventarioTodos_Exclui = gErr
    
    Select Case gErr

        Case 71478
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145870)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Geracao_Trans(dtData As Date) As Long
'Gera Registro de Inventário para todos Produtos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Geracao_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_GERACAO_TRANS", X)
    If lErro <> SUCESSO Then gError 71479
    
    RegInventario_Geracao_Trans = X.RegInventario_Geracao_Trans(dtData)

    Exit Function
    
Erro_RegInventario_Geracao_Trans:
    
    RegInventario_Geracao_Trans = gErr
    
    Select Case gErr

        Case 71479
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145871)
        
    End Select
    
    Exit Function

End Function

Function RegInventarioAlmox_Geracao_Trans(dtData As Date)
'Gera Registro de Inventário para todos Produtos agrupados por Almoxarifado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventarioAlmox_Geracao_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIOALMOX_GERACAO_TRANS", X)
    If lErro <> SUCESSO Then gError 71480
    
    RegInventarioAlmox_Geracao_Trans = X.RegInventarioAlmox_Geracao_Trans(dtData)

    Exit Function
    
Erro_RegInventarioAlmox_Geracao_Trans:
    
    RegInventarioAlmox_Geracao_Trans = gErr
    
    Select Case gErr

        Case 71480
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145872)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quantidades(objRegInventario As ClassRegInventario) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quantidades
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANTIDADES", X)
    If lErro <> SUCESSO Then gError 71481
    
    RegInventario_Obter_Quantidades = X.RegInventario_Obter_Quantidades(objRegInventario)

    Exit Function
    
Erro_RegInventario_Obter_Quantidades:
    
    RegInventario_Obter_Quantidades = gErr
    
    Select Case gErr

        Case 71481
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145873)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst(objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANT_SLDMESEST", X)
    If lErro <> SUCESSO Then gError 71482
    
    RegInventario_Obter_Quantidades_SldMesEst = X.RegInventario_Obter_Quantidades_SldMesEst(objRegInventario, alComando())

    Exit Function
    
Erro_RegInventario_Obter_Quantidades_SldMesEst:
    
    RegInventario_Obter_Quantidades_SldMesEst = gErr
    
    Select Case gErr

        Case 71482
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145874)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst1(objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst1
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANT_SLDMESEST1", X)
    If lErro <> SUCESSO Then gError 71483
    
    RegInventario_Obter_Quantidades_SldMesEst1 = X.RegInventario_Obter_Quantidades_SldMesEst1(objRegInventario, alComando())

    Exit Function
    
Erro_RegInventario_Obter_Quantidades_SldMesEst1:
    
    RegInventario_Obter_Quantidades_SldMesEst1 = gErr
    
    Select Case gErr

        Case 71483
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145875)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quantidades_SldMesEst2(objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quantidades_SldMesEst2
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANT_SLDMESEST2", X)
    If lErro <> SUCESSO Then gError 71484
    
    RegInventario_Obter_Quantidades_SldMesEst2 = X.RegInventario_Obter_Quantidades_SldMesEst2(objRegInventario, alComando())

    Exit Function
    
Erro_RegInventario_Obter_Quantidades_SldMesEst2:
    
    RegInventario_Obter_Quantidades_SldMesEst2 = gErr
    
    Select Case gErr

        Case 71484
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145876)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quant_Mov_UltimoMes(objRegInventario As ClassRegInventario, alComando() As Long) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quant_Mov_UltimoMes
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANT_MOV_ULTIMO", X)
    If lErro <> SUCESSO Then gError 71485
    
    RegInventario_Obter_Quant_Mov_UltimoMes = X.RegInventario_Obter_Quant_Mov_UltimoMes(objRegInventario, alComando())

    Exit Function
    
Erro_RegInventario_Obter_Quant_Mov_UltimoMes:
    
    RegInventario_Obter_Quant_Mov_UltimoMes = gErr
    
    Select Case gErr

        Case 71485
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145877)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Custos(objRegInventario As ClassRegInventario) As Long
'Devolve o custo do produto e dos Escaninhos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Custos
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_CUSTOS", X)
    If lErro <> SUCESSO Then gError 71486
    
    RegInventario_Obter_Custos = X.RegInventario_Obter_Custos(objRegInventario)

    Exit Function
    
Erro_RegInventario_Obter_Custos:
    
    RegInventario_Obter_Custos = gErr
    
    Select Case gErr

        Case 71486
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145878)
        
    End Select
    
    Exit Function

End Function

Function RegInventario_Obter_Quantidades_Almox(objRegInventario As ClassRegInventario) As Long
'Le as Quantidade Inclusive dos escaninhos para data do Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RegInventario_Obter_Quantidades_Almox
    
    lErro = Rotina_ObterObjetoFuncao("FIS_REGINVENTARIO_OBTER_QUANTIDADE_ALMOX", X)
    If lErro <> SUCESSO Then gError 71487
    
    RegInventario_Obter_Quantidades_Almox = X.RegInventario_Obter_Quantidades_Almox(objRegInventario)

    Exit Function
    
Erro_RegInventario_Obter_Quantidades_Almox:
    
    RegInventario_Obter_Quantidades_Almox = gErr
    
    Select Case gErr

        Case 71487
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145879)
        
    End Select
    
    Exit Function

End Function

Function Livro_Fechado_Insere(objLivroFechado As ClassLivrosFechados) As Long
'Insere o registro no livro Fechado com os valores que veio da Tabela de Livro Filial passado no objLivroFechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Livro_Fechado_Insere
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVRO_FECHADO_INSERE", X)
    If lErro <> SUCESSO Then gError 71488
    
    Livro_Fechado_Insere = X.Livro_Fechado_Insere(objLivroFechado)

    Exit Function
    
Erro_Livro_Fechado_Insere:
    
    Livro_Fechado_Insere = gErr
    
    Select Case gErr

        Case 71488
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145880)
        
    End Select
    
    Exit Function

End Function

Function Livro_Filial_Atualiza_Trans(objLivroFilial As ClassLivrosFilial, objLivroFilialLido As ClassLivrosFilial) As Long
'Atualiza os Valores de DataInicial, DataFinal, NumeroProxLivro, NumeroProxFolha, ImpressoEm
'Depois retorna o Valor que está no BD para poder inserir no Livro Fechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Livro_Filial_Atualiza_Trans
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVRO_FILIAL_ATUALIZA_TRANS", X)
    If lErro <> SUCESSO Then gError 71489
    
    Livro_Filial_Atualiza_Trans = X.Livro_Filial_Atualiza_Trans(objLivroFilial, objLivroFilialLido)

    Exit Function
    
Erro_Livro_Filial_Atualiza_Trans:
    
    Livro_Filial_Atualiza_Trans = gErr
    
    Select Case gErr

        Case 71489
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145881)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegEntrada(objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro_RegEntrada
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO_REGENTRADA", X)
    If lErro <> SUCESSO Then gError 71490
    
    Rotina_Fechamento_Livro_RegEntrada = X.Rotina_Fechamento_Livro_RegEntrada(objLivroFechado)

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegEntrada:
    
    Rotina_Fechamento_Livro_RegEntrada = gErr
    
    Select Case gErr

        Case 71490
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145882)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro(objLivrosFilial As ClassLivrosFilial) As Long
'Esta Rotina faz o Fechamento do Livro passado em objLivroFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO", X)
    If lErro <> SUCESSO Then gError 71491
    
    Rotina_Fechamento_Livro = X.Rotina_Fechamento_Livro(objLivrosFilial)

    Exit Function
    
Erro_Rotina_Fechamento_Livro:
    
    Rotina_Fechamento_Livro = gErr
    
    Select Case gErr

        Case 71491
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145883)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegSaida(objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Saida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro_RegSaida
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO_REGSAIDA", X)
    If lErro <> SUCESSO Then gError 71492
    
    Rotina_Fechamento_Livro_RegSaida = X.Rotina_Fechamento_Livro_RegSaida(objLivroFechado)

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegSaida:
    
    Rotina_Fechamento_Livro_RegSaida = gErr
    
    Select Case gErr

        Case 71492
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145884)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegInventario(objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Inventario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro_RegInventario
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO_REGINVENTARI", X)
    If lErro <> SUCESSO Then gError 71493
    
    Rotina_Fechamento_Livro_RegInventario = X.Rotina_Fechamento_Livro_RegInventario(objLivroFechado)

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegInventario:
    
    Rotina_Fechamento_Livro_RegInventario = gErr
    
    Select Case gErr

        Case 71493
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145885)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegApuracaoICMS(objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Apuracao ICMS

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro_RegApuracaoICMS
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO_REGAPURICMS", X)
    If lErro <> SUCESSO Then gError 71494
    
    Rotina_Fechamento_Livro_RegApuracaoICMS = X.Rotina_Fechamento_Livro_RegApuracaoICMS(objLivroFechado)

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegApuracaoICMS:
    
    Rotina_Fechamento_Livro_RegApuracaoICMS = gErr
    
    Select Case gErr

        Case 71494
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145886)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegApuracaoIPI(objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Apuracao IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Fechamento_Livro_RegApuracaoIPI
    
    lErro = Rotina_ObterObjetoFuncao("FIS_ROTINA_FECHAMENTO_LIVRO_REGAPURIPI", X)
    If lErro <> SUCESSO Then gError 71495
    
    Rotina_Fechamento_Livro_RegApuracaoIPI = X.Rotina_Fechamento_Livro_RegApuracaoIPI(objLivroFechado)

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegApuracaoIPI:
    
    Rotina_Fechamento_Livro_RegApuracaoIPI = gErr
    
    Select Case gErr

        Case 71495
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145887)
        
    End Select
    
    Exit Function

End Function

Function LivrosFilial_Le_TodosAbertos(iTributo As Integer, colLivrosFilial As Collection) As Long
'Lê Todos os Livros Abertos para o Tributo passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFilial_Le_TodosAbertos
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFILIAL_LE_TODOSABERTOS", X)
    If lErro <> SUCESSO Then gError 71496
    
    LivrosFilial_Le_TodosAbertos = X.LivrosFilial_Le_TodosAbertos(iTributo, colLivrosFilial)

    Exit Function
    
Erro_LivrosFilial_Le_TodosAbertos:
    
    LivrosFilial_Le_TodosAbertos = gErr
    
    Select Case gErr

        Case 71496
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145888)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoICMS_Le_IntervaloData(objApuracao As ClassRegApuracao) As Long
'Procura por ApuraçãoICMS com data Inicial e Final maiores do que as datas passadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoICMS_Le_IntervaloData
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOICMS_LE_INTERVALODATA", X)
    If lErro <> SUCESSO Then gError 71497
    
    ApuracaoICMS_Le_IntervaloData = X.ApuracaoICMS_Le_IntervaloData(objApuracao)

    Exit Function
    
Erro_ApuracaoICMS_Le_IntervaloData:
    
    ApuracaoICMS_Le_IntervaloData = gErr
    
    Select Case gErr

        Case 71497
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145889)
        
    End Select
    
    Exit Function

End Function

Function ApuracaoIPI_Le_IntervaloData(objApuracao As ClassRegApuracao) As Long
'Procura por ApuraçãoIPI com data Inicial e Final maiores do que as datas passadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApuracaoIPI_Le_IntervaloData
    
    lErro = Rotina_ObterObjetoFuncao("FIS_APURACAOIPI_LE_INTERVALODATA", X)
    If lErro <> SUCESSO Then gError 71498
    
    ApuracaoIPI_Le_IntervaloData = X.ApuracaoIPI_Le_IntervaloData(objApuracao)

    Exit Function
    
Erro_ApuracaoIPI_Le_IntervaloData:
    
    ApuracaoIPI_Le_IntervaloData = gErr
    
    Select Case gErr

        Case 71498
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145890)
        
    End Select
    
    Exit Function

End Function

Function LivrosFilial_Atualiza_DataImpresao(objLivroFilial As ClassLivrosFilial) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFilial_Atualiza_DataImpresao
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFILIAL_ATUALIZA_DATAIMPRESSAO", X)
    If lErro <> SUCESSO Then gError 71499
    
    LivrosFilial_Atualiza_DataImpresao = X.LivrosFilial_Atualiza_DataImpresao(objLivroFilial)

    Exit Function
    
Erro_LivrosFilial_Atualiza_DataImpresao:
    
    LivrosFilial_Atualiza_DataImpresao = gErr
    
    Select Case gErr

        Case 71499
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145891)
        
    End Select
    
    Exit Function

End Function

Function LivrosFechados_Le_IntervaloData(objLivroFechado As ClassLivrosFechados) As Long
'Procura por LivroFechado com data Inicial e Final maiores do que as datas passadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFechados_Le_IntervaloData
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFECHADOS_LE_INTERVALODATA", X)
    If lErro <> SUCESSO Then gError 88000
    
    LivrosFechados_Le_IntervaloData = X.LivrosFechados_Le_IntervaloData(objLivroFechado)

    Exit Function
    
Erro_LivrosFechados_Le_IntervaloData:
    
    LivrosFechados_Le_IntervaloData = gErr
    
    Select Case gErr

        Case 88000
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145892)
        
    End Select
    
    Exit Function

End Function

Function LivrosFilial_Le_IntervaloData(objLivrosFilial As ClassLivrosFilial) As Long
'Procura por LivrosFilial com data Inicial e Final maiores do que as datas passadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFilial_Le_IntervaloData
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFILIAL_LE_INTERVALODATA", X)
    If lErro <> SUCESSO Then gError 88001
    
    LivrosFilial_Le_IntervaloData = X.LivrosFilial_Le_IntervaloData(objLivrosFilial)

    Exit Function
    
Erro_LivrosFilial_Le_IntervaloData:
    
    LivrosFilial_Le_IntervaloData = gErr
    
    Select Case gErr

        Case 88001
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145893)
        
    End Select
    
    Exit Function

End Function

Function LivrosFiscais_Valida_Data(dtData As Date) As Long
'Verifica se a data passada como parametro está entre a Data Inicial
'e a Data Final de algum LivroAberto ou Fechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LivrosFiscais_Valida_Data
    
    lErro = Rotina_ObterObjetoFuncao("FIS_LIVROSFISCAIS_VALIDA_DATA", X)
    If lErro <> SUCESSO Then gError 88002
    
    LivrosFiscais_Valida_Data = X.LivrosFiscais_Valida_Data(dtData)

    Exit Function
    
Erro_LivrosFiscais_Valida_Data:
    
    LivrosFiscais_Valida_Data = gErr
    
    Select Case gErr

        Case 88002
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145894)
        
    End Select
    
    Exit Function

End Function

