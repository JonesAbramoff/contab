VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFilFornFilEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lComando As Long
Private objUltFilFornFilEmpEst As New ClassFilialFornecedorEst 'guarda ultimo objeto lido

Private lComandoSaldoTit As Long
Private objUltFilFornSaldoTit As New ClassFilialFornecedorEst 'guarda ultimo objeto lido

Private lComandoUltCompra As Long
Private objUltFilFornUltCompra As New ClassFilialFornecedorEst 'guarda ultimo objeto lido

Public Function Obter_PedCompraAberto(objFilFornFilEmp As ClassFilFornFilEmp)
'Devolve para o relatório o # de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_PedCompraAberto

    'se trocou a chave (fornecedor/filial/filialempresa)
    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84673
        
    End If

    objFilFornFilEmp.lPedidosEmAberto = objUltFilFornFilEmpEst.lPedidosEmAberto

    Obter_PedCompraAberto = SUCESSO

    Exit Function

Erro_Obter_PedCompraAberto:

    Obter_PedCompraAberto = gErr

    Select Case gErr

        Case 84673

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153126)

    End Select

    Exit Function

End Function

Public Function Obter_NumCompras(objFilFornFilEmp As ClassFilFornFilEmp)
'Devolve para o relatório o # de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_NumCompras

    'se trocou a chave (fornecedor/filial/filialempresa)
    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84674
        
    End If

    objFilFornFilEmp.lNumeroCompras = objUltFilFornFilEmpEst.lNumeroCompras

    Obter_NumCompras = SUCESSO

    Exit Function

Erro_Obter_NumCompras:

    Obter_NumCompras = gErr

    Select Case gErr

        Case 84674

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153127)

    End Select

    Exit Function

End Function

Public Function Obter_MaiorAtraso(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Maior Atraso do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_MaiorAtraso

    'se trocou a chave (fornecedor/filial/filialempresa)
    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84675
        
    End If

    objFilFornFilEmp.lMaiorAtraso = objUltFilFornFilEmpEst.lMaiorAtraso

    Obter_MaiorAtraso = SUCESSO

    Exit Function

Erro_Obter_MaiorAtraso:

    Obter_MaiorAtraso = gErr

    Select Case gErr

        Case 84675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153128)

    End Select

    Exit Function

End Function

'Public Function FilFornFilEmp_Obter_MediaCompra(dMediaCompra As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório a Media de compras do / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaCompra
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84619
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_MediaCompra = objRelFilFornFilEmp.Obter_MediaCompra(objFilFornFilEmp)
'
'    dMediaCompra = objFilFornFilEmp.dMediaCompra
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_MediaCompra:
'
'    FilFornFilEmp_Obter_MediaCompra = gErr
'
'    Select Case gErr
'
'        Case 84619
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153129)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function FilFornFilEmp_Obter_NumDevolucoes(lNumDevolucoes As Long, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório o numero de Devolucoes do / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaCompra
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84620
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_NumDevolucoes = objRelFilFornFilEmp.Obter_NumDevolucoes(objFilFornFilEmp)
'
'    lNumDevolucoes = objFilFornFilEmp.lNumDevolucoes
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_MediaCompra:
'
'    FilFornFilEmp_Obter_MediaCompra = gErr
'
'    Select Case gErr
'
'        Case 84620
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153130)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Public Function FilFornFilEmp_Obter_UltimaCompra(dtUltimaCompra As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório a Data da Última compra do / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaCompra
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84621
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_UltimaCompra = objRelFilFornFilEmp.Obter_UltimaCompra(objFilFornFilEmp)
'
'    dtUltimaCompra = objFilFornFilEmp.dtUltimaCompra
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_UltimaCompra:
'
'    FilFornFilEmp_Obter_UltimaCompra = gErr
'
'    Select Case gErr
'
'        Case 84621
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153131)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function FilFornFilEmp_Obter_MediaAtraso(lMediaAtraso As Long, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório a média de atrasos do / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaAtraso
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84622
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_MediaAtraso = objRelFilFornFilEmp.Obter_MaediaAtraso(objFilFornFilEmp)
'
'    lMediaAtraso = objFilFornFilEmp.lMediaAtraso
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_MediaAtraso:
'
'    FilFornFilEmp_Obter_MediaAtraso = gErr
'
'    Select Case gErr
'
'        Case 84622
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153132)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function FilFornFilEmp_Obter_MaiorAtraso(lMaiorAtraso As Long, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório o maior atraso do / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaAtraso
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84623
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_MaiorAtraso = objRelFilFornFilEmp.Obter_MaiorAtraso(objFilFornFilEmp)
'
'    lMaiorAtraso = objFilFornFilEmp.lMaiorAtraso
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_MaiorAtraso:
'
'    FilFornFilEmp_Obter_MaiorAtraso = gErr
'
'    Select Case gErr
'
'        Case 84623
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153133)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function FilFornFilEmp_Obter_PedCompra(lPedCompra As Long, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório os Pedidos de Compras em Aberto / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaAtraso
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84624
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_PedCompra = objRelFilFornFilEmp.Obter_PedCompra(objFilFornFilEmp)
'
'    lPedCompra = objFilFornFilEmp.lPedCompra
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_PedCompra:
'
'    FilFornFilEmp_Obter_PedCompra = gErr
'
'    Select Case gErr
'
'        Case 84624
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153134)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function FilFornFilEmp_Obter_SaldoTitulos(dSaldoTitulos As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
''Devolve para o relatório os Saldos em Títulos / fornecedor / filialforn na filialempresa
'
'Dim lErro As Long
'Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'
'On Error GoTo Erro_FilFornFilEmp_Obter_MediaAtraso
'
'    'Carrega o obj, caso ele já tenha sido preenchido
'    Set objFilFornFilEmp = ObtemObj("RelFilFornFilEmp")
'
'    'Se o obj estiver vazio
'    If objFilFornFilEmp Is Nothing Then
'
'        'Seta o obj como uma nova classe com os dados do Relatório
'        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp
'
'        'Guarda o obj em uma coleção
'        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
'        If lErro <> SUCESSO Then gError 84625
'
'    End If
'
'    'Preenche o obj que será passado como parâmetro para a função
'    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
'    objFilFornFilEmp.lFornecedor = lFornecedor
'    objFilFornFilEmp.iFilialForn = iFilialForn
'
'    FilFornFilEmp_Obter_SaldoTitulos = objRelFilFornFilEmp.Obter_SaldoTitulos(objFilFornFilEmp)
'
'    dSladoTitulos = objFilFornFilEmp.dSaldoTitulos
'
'    Exit Function
'
'Erro_FilFornFilEmp_Obter_SaldoTitulos:
'
'    FilFornFilEmp_Obter_SaldoTitulos = gErr
'
'    Select Case gErr
'
'        Case 84625
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153135)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Obter_MediaAtraso(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório a Media de Atraso de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_MediaAtraso

If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84676
        
    End If

    Obter_MediaAtraso = SUCESSO

    Exit Function

Erro_Obter_MediaAtraso:

    Obter_MediaAtraso = gErr

    Select Case gErr

        Case 84676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153136)

    End Select

    Exit Function

End Function

Public Function Obter_ValorAcumulado(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_ValorAcumulado

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84678
        
    End If

    objFilFornFilEmp.dValorAcumuladoCompras = objUltFilFornFilEmpEst.dValorAcumuladoCompras

    Obter_ValorAcumulado = SUCESSO

    Exit Function

Erro_Obter_ValorAcumulado:

    Obter_ValorAcumulado = gErr

    Select Case gErr

        Case 84678

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153137)

    End Select

    Exit Function

End Function

Public Function Obter_MediaCompra(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_MediaCompra

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84679
        
    End If

    objFilFornFilEmp.dMediaCompra = objUltFilFornFilEmpEst.dMediaCompra

    Obter_MediaCompra = SUCESSO

    Exit Function

Erro_Obter_MediaCompra:

    Obter_MediaCompra = gErr

    Select Case gErr

        Case 84679

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153138)

    End Select

    Exit Function

End Function

Public Function Obter_NumDevolucoes(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_NumDevolucoes

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84680
        
    End If

    objFilFornFilEmp.lNumeroCompras = objUltFilFornFilEmpEst.lNumTotalDevolucoes

    Obter_NumDevolucoes = SUCESSO

    Exit Function

Erro_Obter_NumDevolucoes:

    Obter_NumDevolucoes = gErr

    Select Case gErr

        Case 84680

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153139)

    End Select

    Exit Function

End Function

Public Function Obter_UltimaCompra(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_UltimaCompra

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_UltimaCompra_Rel(objFilFornFilEmp, lComando)
        If lErro <> SUCESSO Then gError 84681
        
    End If

    objFilFornFilEmp.dtDataUltimaCompra = objUltFilFornFilEmpEst.dtDataUltimaCompra

    Obter_UltimaCompra = SUCESSO

    Exit Function

Erro_Obter_UltimaCompra:

    Obter_UltimaCompra = gErr

    Select Case gErr

        Case 84681

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153140)

    End Select

    Exit Function

End Function

Public Function Obter_PedCompra(objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long

On Error GoTo Erro_Obter_PedCompra

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objUltFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objUltFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objUltFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_Rel(objUltFilFornFilEmpEst, lComando)
        If lErro <> SUCESSO Then gError 84682
        
    End If

    objFilFornFilEmp.lPedidosEmAberto = objUltFilFornFilEmpEst.lPedidosEmAberto

    Obter_PedCompra = SUCESSO

    Exit Function

Erro_Obter_PedCompra:

    Obter_PedCompra = gErr

    Select Case gErr

        Case 84682

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153141)

    End Select

    Exit Function

End Function

Public Function Obter_SaldoTitulos(objFilFornFilEmp As ClassFilFornFilEmp, dSaldoTitulos As Double) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objFilFornFilEmpEst As New ClassFilialFornecedorEst

On Error GoTo Erro_Obter_SaldoTitulos

    If objUltFilFornFilEmpEst.lCodFornecedor <> objFilFornFilEmp.lCodFornecedor Or objUltFilFornFilEmpEst.iCodFilial <> objFilFornFilEmp.iCodFilial Or objUltFilFornFilEmpEst.iFilialEmpresa <> objFilFornFilEmp.iFilialEmpresa Then
        
        objFilFornFilEmpEst.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objFilFornFilEmpEst.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objFilFornFilEmpEst.iCodFilial = objFilFornFilEmp.iCodFilial
    
        lErro = FilialFornEst_Le_SaldoTit_Rel(objFilFornFilEmp, lComando)
        If lErro <> SUCESSO Then gError 84683
        
    End If

    dSaldoTitulos = objUltFilFornFilEmpEst.dSaldoTitulos
    
    Obter_SaldoTitulos = SUCESSO

    Exit Function

Erro_Obter_SaldoTitulos:

    Obter_SaldoTitulos = gErr

    Select Case gErr

        Case 84683

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153142)

    End Select

    Exit Function

End Function

Public Function FilialFornEst_Le_Rel(objFilFornFilEmpEst As ClassFilialFornecedorEst, lComando As Long) As Long

Dim lErro As Long, objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilialFornEst_Le_Rel

    objFilFornFilEmp.lCodFornecedor = objFilFornFilEmpEst.lCodFornecedor
    objFilFornFilEmp.iFilialEmpresa = objFilFornFilEmpEst.iFilialEmpresa
    objFilFornFilEmp.iCodFilial = objFilFornFilEmpEst.iCodFilial
    
    lErro = FornFilEmp_Le_Rel(objFilFornFilEmp, lComando)
    If lErro <> SUCESSO Then gError 84684

    With objFilFornFilEmpEst
        .lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto
        .dMediaCompra = objFilFornFilEmp.dMediaCompra
        .dtDataPrimeiraCompra = objFilFornFilEmp.dtDataPrimeiraCompra
        .dtDataUltDevolucao = objFilFornFilEmp.dtDataUltDevolucao
        .dtDataUltimaCompra = objFilFornFilEmp.dtDataUltimaCompra
        .dValorAcumuladoCompras = objFilFornFilEmp.dValorAcumuladoCompras
        .dValorAcumuladoDevolucoes = objFilFornFilEmp.dValorAcumuladoDevolucoes
        .lNumeroCompras = objFilFornFilEmp.lNumeroCompras
        .lNumTotalDevolucoes = objFilFornFilEmp.lNumTotalDevolucoes
        .lMaiorAtraso = objFilFornFilEmp.lMaiorAtraso
    End With
    
    FilialFornEst_Le_Rel = SUCESSO

    Exit Function

Erro_FilialFornEst_Le_Rel:

    FilialFornEst_Le_Rel = gErr
    
    Select Case gErr
    
        Case 84684
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153143)
    
    End Select
    
    Exit Function

End Function




Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Function FornFilEmp_Le1(objFornFilEmp As ClassFilFornFilEmp, lComando As Long) As Long
'Lê os dados da tabela FilFornFilEmp correspondentes a
'Fornecedor e FilialEmpresa passados em objFornFilEmp

Dim lErro As Long
Dim tFornFilEmp As typeFilFornFilEmp
Dim lNumFiliaisForn As Long

On Error GoTo Erro_FornFilEmp_Le1

    'Lê a tabela FilFornFilEmp para trazer os dados
    lErro = Comando_Executar(lComando, "SELECT SUM(NumeroCompras), SUM(MediaCompra), COUNT(*), MIN(DataPrimeiraCompra), MAX(DataUltimaCompra), SUM(ValorAcumuladoCompras), SUM(ValorAcumuladoDevolucoes), SUM(NumTotalDevolucoes), MAX(DataUltDevolucao) , MAX(MaiorAtraso), SUM(PedidosEmAberto), SUM(AtrasoAcumulado), SUM(PedidosRecebidos), SUM(ItensPedidosRecebidos) FROM FilialFornFilEmp WHERE FilialEmpresa = ? AND CodFornecedor =? ", tFornFilEmp.lNumeroCompras, tFornFilEmp.dMediaCompra, lNumFiliaisForn, tFornFilEmp.dtDataPrimeiraCompra, tFornFilEmp.dtDataUltimaCompra, tFornFilEmp.dValorAcumuladoCompras, tFornFilEmp.dValorAcumuladoDevolucoes, tFornFilEmp.lNumTotalDevolucoes, tFornFilEmp.dtDataUltDevolucao, tFornFilEmp.lMaiorAtraso, tFornFilEmp.lPedidosEmAberto, tFornFilEmp.lAtrasoAcumulado, tFornFilEmp.lPedidosRecebidos, tFornFilEmp.lItensPedidosRecebidos, objFornFilEmp.iFilialEmpresa, objFornFilEmp.lCodFornecedor)
    If lErro <> AD_SQL_SUCESSO Then gError 84685

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84686
   
    objFornFilEmp.lNumeroCompras = tFornFilEmp.lNumeroCompras
    objFornFilEmp.dtDataPrimeiraCompra = tFornFilEmp.dtDataPrimeiraCompra
    objFornFilEmp.dtDataUltimaCompra = tFornFilEmp.dtDataUltimaCompra
    objFornFilEmp.dValorAcumuladoCompras = tFornFilEmp.dValorAcumuladoCompras
    objFornFilEmp.dValorAcumuladoDevolucoes = tFornFilEmp.dValorAcumuladoDevolucoes
    objFornFilEmp.lNumTotalDevolucoes = tFornFilEmp.lNumTotalDevolucoes
    objFornFilEmp.dtDataUltDevolucao = tFornFilEmp.dtDataUltDevolucao
    objFornFilEmp.lMaiorAtraso = tFornFilEmp.lMaiorAtraso
    objFornFilEmp.lPedidosEmAberto = tFornFilEmp.lPedidosEmAberto
    objFornFilEmp.lAtrasoAcumulado = tFornFilEmp.lAtrasoAcumulado
    objFornFilEmp.lPedidosRecebidos = tFornFilEmp.lPedidosRecebidos
    objFornFilEmp.lItensPedidosRecebidos = tFornFilEmp.lItensPedidosRecebidos
    
    If lNumFiliaisForn > 0 Then
    
        objFornFilEmp.dMediaCompra = tFornFilEmp.dMediaCompra / lNumFiliaisForn
    Else
    
        objFornFilEmp.dMediaCompra = 0
    
    End If
            
    FornFilEmp_Le1 = SUCESSO
    
    Exit Function
    
Erro_FornFilEmp_Le1:

    FornFilEmp_Le1 = gErr
    
    Select Case gErr
        
        Case 84685, 84686
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALFORNFILEMP", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153144)

    End Select
        
    Exit Function
    
End Function

Function FornFilEmp_Le_Rel(objFornFilEmp As ClassFilFornFilEmp, lComando As Long) As Long
'Lê os dados da tabela FilFornFilEmp correspondentes a
'Fornecedor e FilialEmpresa passados em objFornFilEmp

Dim lErro As Long

On Error GoTo Erro_FornFilEmp_Le_Rel

    If lComando = 0 Then
    
        'Inicializa comandos
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 84687

    End If

    lErro = FornFilEmp_Le1(objFornFilEmp, lComando)
    If lErro <> SUCESSO Then gError 84688
            
    FornFilEmp_Le_Rel = SUCESSO
    
    Exit Function
    
Erro_FornFilEmp_Le_Rel:

    FornFilEmp_Le_Rel = gErr
    
    Select Case gErr
        
        Case 84687
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 84688
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153145)

    End Select
        
    Exit Function
    
End Function


Public Function FilialFornEst_Le_SaldoTit_Rel(objFilFornSaldoTit As ClassFilialFornecedorEst, lComandoSaldoTit As Long) As Long

Dim lErro As Long
Dim dSaldoTitulos As Double
Dim objFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilialFornEst_Le_SaldoTit_Rel
    
    'Lê a soma dos saldo em TitulosPagar
    lErro = Comando_Executar(lComando, "SELECT SUM(Saldo) FROM TitulosPag WHERE Fornecedor =? AND FilialEmpresa= ? ", dSaldoTitulos, objFilFornSaldoTit.lCodFornecedor, objFilFornSaldoTit.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 84689

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84690

    objFilFornSaldoTit.dSaldoTitulos = dSaldoTitulos
           
    objUltFilFornSaldoTit.dSaldoTitulos = objFilFornSaldoTit.dSaldoTitulos
    objUltFilFornSaldoTit.lCodFornecedor = objFilFornSaldoTit.lCodFornecedor
    objUltFilFornSaldoTit.iFilialEmpresa = objFilFornSaldoTit.iFilialEmpresa
    objUltFilFornSaldoTit.iCodFilial = objFilFornSaldoTit.iCodFilial
    
    
    FilialFornEst_Le_SaldoTit_Rel = SUCESSO
    
    Exit Function
    
Erro_FilialFornEst_Le_SaldoTit_Rel:

    FilialFornEst_Le_SaldoTit_Rel = gErr
    
    Select Case gErr
        
        Case 84689, 84690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153146)

    End Select

    Exit Function


End Function

Public Function FilialFornEst_Le_UltimaCompra_Rel(objUltFilFornUltCompra As ClassFilialFornecedorEst, lComandoUltCompra As Long) As Long

Dim lErro As Long

On Error GoTo erro_FilialFornEst_Le_UltimaCompra_Rel

    lErro = CF("Pesquisa_Ultima_Compra_Forn",objUltFilFornUltCompra)
    If lErro <> SUCESSO Then gError 84691
          
    FilialFornEst_Le_UltimaCompra_Rel = SUCESSO
      
    Exit Function
    
erro_FilialFornEst_Le_UltimaCompra_Rel:

    FilialFornEst_Le_UltimaCompra_Rel = gErr
    
    Select Case gErr
    
        Case 84691

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153147)

    End Select
    
    Exit Function

End Function
