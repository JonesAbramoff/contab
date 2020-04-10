VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelectDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Dan_OS_Le(ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim tDan_OS As typeDan_OS

On Error GoTo Erro_Dan_OS_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184771
    Next

    'Alocação de espaço no buffer
    tDan_OS.sOS = String(STRING_DAN_OS, 0)
    tDan_OS.sModelo = String(STRING_DAN_OS_MODELO, 0)
    tDan_OS.sNumSerie = String(STRING_DAN_OS_NUMSERIE, 0)

    'Le a tabelaDan_OS
    lErro = Comando_Executar(alComando(0), "SELECT OS, Cliente, Modelo, NumSerie FROM Dan_OS WHERE OS= ? ", _
                tDan_OS.sOS, tDan_OS.lCliente, tDan_OS.sModelo, tDan_OS.sNumSerie, _
                objDan_OS.sOS)
    If lErro <> AD_SQL_SUCESSO Then gError 184772

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184773

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objDan_OS.sOS = tDan_OS.sOS
    objDan_OS.lCliente = tDan_OS.lCliente
    objDan_OS.sModelo = tDan_OS.sModelo
    objDan_OS.sNumSerie = tDan_OS.sNumSerie
    
    lErro = Dan_ItensOS_Le(alComando(1), objDan_OS)
    If lErro <> SUCESSO Then gError 184774

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Dan_OS_Le = SUCESSO

    Exit Function

Erro_Dan_OS_Le:

    Dan_OS_Le = gErr

    Select Case gErr

        Case 184771
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184772, 184773
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_OS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 184774

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184775)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function Dan_ItensOS_Le(ByVal lComando As Long, ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim tDan_ItensOS As typeDan_ItensOS
Dim objDan_ItensOS As ClassDan_ItensOS

On Error GoTo Erro_Dan_ItensOS_Le

    'Alocação de espaço no buffer
    tDan_ItensOS.sOS = String(STRING_DAN_OS, 0)
    tDan_ItensOS.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaDan_ItensOS
    lErro = Comando_Executar(lComando, "SELECT OS, Item, Produto, Quantidade FROM Dan_ItensOS WHERE OS= ? ", _
                tDan_ItensOS.sOS, tDan_ItensOS.iItem, tDan_ItensOS.sProduto, tDan_ItensOS.dQuantidade, _
                objDan_OS.sOS)
    If lErro <> AD_SQL_SUCESSO Then gError 198242

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198243

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objDan_ItensOS = New ClassDan_ItensOS
    
        objDan_ItensOS.sOS = tDan_ItensOS.sOS
        objDan_ItensOS.iItem = tDan_ItensOS.iItem
        objDan_ItensOS.sProduto = tDan_ItensOS.sProduto
        objDan_ItensOS.dQuantidade = tDan_ItensOS.dQuantidade
        
        objDan_OS.colItens.Add objDan_ItensOS
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198244
        
    Loop


    Dan_ItensOS_Le = SUCESSO

    Exit Function

Erro_Dan_ItensOS_Le:

    Dan_ItensOS_Le = gErr

    Select Case gErr
        
        Case 198242, 198243, 198244
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_ITENSOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198241)

    End Select

    Exit Function

End Function

Public Function Dan_ItemOS_Le_OSProd(ByVal objDan_ItensOS As ClassDan_ItensOS) As Long

Dim lErro As Long, lComando As Long
Dim tDan_ItensOS As typeDan_ItensOS

On Error GoTo Erro_Dan_ItemOS_Le_OSProd

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184852
    
    'Le a tabelaDan_ItensOS
    lErro = Comando_Executar(lComando, "SELECT Item, Quantidade FROM Dan_ItensOS WHERE OS= ? AND Produto = ?", _
                tDan_ItensOS.iItem, tDan_ItensOS.dQuantidade, _
                objDan_ItensOS.sOS, objDan_ItensOS.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 184853

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184854

    If lErro <> AD_SQL_SUCESSO Then gError 184855
    
    objDan_ItensOS.iItem = tDan_ItensOS.iItem
    objDan_ItensOS.dQuantidade = tDan_ItensOS.dQuantidade
        
    Call Comando_Fechar(lComando)
    
    Dan_ItemOS_Le_OSProd = SUCESSO

    Exit Function

Erro_Dan_ItemOS_Le_OSProd:

    Dan_ItemOS_Le_OSProd = gErr

    Select Case gErr
        
        Case 184852
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 184853, 184854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_ITENSOS", gErr)
        
        Case 184855
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184851)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


