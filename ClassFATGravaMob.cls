VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PVItens_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objPVMotivo As ClassPVMotivo
Dim objItemPV As ClassItemPedido
Dim iFilialEmpresa As Long

On Error GoTo Erro_PVItens_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197740
    Next



    lErro = Comando_Executar(alComando(1), "DELETE FROM ItensPVMotivo WHERE FilialEmpresa = ? And CodPedido = ?", objPV.iFilialEmpresa, objPV.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197743


    If Not objPV.objInfoUsu Is Nothing Then

        For Each objPVMotivo In objPV.objInfoUsu.gcolMotivos
            
            If Len(Trim(objPVMotivo.sMotivo)) <> 0 Then
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO ItensPVMotivo (FilialEmpresa, CodPedido, Produto, Motivo, PrecoTabela) VALUES (?,?,?,?,?) ", objPV.iFilialEmpresa, objPV.lCodigo, objPVMotivo.sProduto, objPVMotivo.sMotivo, objPVMotivo.dPrecoTabela)
                If lErro <> AD_SQL_SUCESSO Then gError 197744
    
            End If
    
        Next

    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PVItens_Grava_Customizado = SUCESSO

    Exit Function

Erro_PVItens_Grava_Customizado:

    PVItens_Grava_Customizado = gErr

    Select Case gErr

        Case 197740
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197741, 197742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVMOTIVO", gErr, objPV.iFilialEmpresa, objPV.lCodigo, objPVMotivo.sProduto)

        Case 197743
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSPVMOTIVO", gErr, objPV.iFilialEmpresa, objPV.lCodigo, objPVMotivo.sProduto)

        Case 197744
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPVMOTIVO", gErr, objPV.iFilialEmpresa, objPV.lCodigo, objPVMotivo.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197745)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

''Function VW_ImportarArquivos() As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_VW_ImportarArquivos
''
''    'percorrer o diretorio ou permitir ao usuario selecionar os arquivos jogando-os para uma coleção
''
''
''
''    VW_ImportarArquivos = SUCESSO
''
''    Exit Function
''
''Erro_VW_ImportarArquivos:
''
''    VW_ImportarArquivos = gErr
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201000)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function VW_ImportarClientes(ByVal sNomeArq As String) As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_VW_ImportarClientes
''
''    'abrir arquivo
''
''    'ler 1o registro
''
''    'tratar header
''    'validar header
''
''    'ler proximo
''
''    bArquivoIntegro = False
''
''    'enquanto nao for fim de arquivo
''
''        'obter tipo de registro
''        Select case
''
''            case 'fim de arquivo
''                'conferir totais
''                bArquivoIntegro = True
''                Exit Do
''
''            Case 1
''                'guarda informacoes em memoria
''                'insere registro na tabela de registros recebidos
''
''            Case 2
''                'guarda informacoes em memoria
''                'insere registro na tabela de registros recebidos
''                'verificar se cliente já estava cadastrado
''                'se nao estava, mover dados para obj e incluir (registrando codigo externo)
''                'senao (já existia)
''                    '??? verificar se é igual ao que foi enviado anteriormente
''                    '??? ou verificar se algo foi alterado
''                    'se necessario, fazer a atualizacao
''
''            Case Else
''                erro
''
''        End Select
''
''        'ler proximo registro
''
''    Loop
''
''    If Not bArquivoIntegro Then gError
''
''    VW_ImportarClientes = SUCESSO
''
''    Exit Function
''
''Erro_VW_ImportarClientes:
''
''    VW_ImportarClientes = gErr
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201000)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function VW_ImportarPedidos() As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_VW_ImportarPedidos
''
''
''    VW_ImportarPedidos = SUCESSO
''
''    Exit Function
''
''Erro_VW_ImportarPedidos:
''
''    VW_ImportarPedidos = gErr
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201000)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function VW_ProtocoloReceb() As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_VW_ProtocoloReceb
''
''    VW_ProtocoloReceb = SUCESSO
''
''    Exit Function
''
''Erro_VW_ProtocoloReceb:
''
''    VW_ProtocoloReceb = gErr
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201000)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function VW_PendenciasPV() As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_VW_PendenciasPV
''
''    VW_PendenciasPV = SUCESSO
''
''    Exit Function
''
''Erro_VW_PendenciasPV:
''
''    VW_PendenciasPV = gErr
''
''    Select Case gErr
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201000)
''
''    End Select
''
''    Exit Function
''
''End Function
''
