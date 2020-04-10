VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ALTERAÇÕES INPAL
'1 - Márcio 08/2000 - Incluindo os campos PadraoTaxaFin, PadraoJuros, TaxaFinanceira, Juros na gravação do Cliente para a INPAL

'??? por que as funcoes ???_Lock estao fazendo lock exclusive ??

Option Explicit

Public Function Cliente_Le_Customizado(ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim lComando As Long
Dim objClienteInpal As New ClassClienteInpal
Dim dJuros As Double, dTaxaFinanceira As Double, iPadraoJuros As Integer, iPadraoTaxaFin As Integer

On Error GoTo Erro_Cliente_Le_Customizado

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 192936

    'Le a tabelaClientesAF
    lErro = Comando_Executar(lComando, "SELECT Juros, TaxaFinanceira, PadraoJuros, PadraoTaxaFin FROM Clientes WHERE Codigo= ? ", _
                dJuros, dTaxaFinanceira, iPadraoJuros, iPadraoTaxaFin, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 192937

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192938
    
    If lErro <> AD_SQL_SEM_DADOS Then

        objClienteInpal.dJuros = dJuros
        objClienteInpal.dTaxaFinanceira = dTaxaFinanceira
        objClienteInpal.iPadraoJuros = iPadraoJuros
        objClienteInpal.iPadraoTaxaFin = iPadraoTaxaFin
        
    End If
    
    Set objCliente.objInfoUsu = objClienteInpal

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Cliente_Le_Customizado = SUCESSO

    Exit Function

Erro_Cliente_Le_Customizado:

    Cliente_Le_Customizado = gErr

    Select Case gErr

        Case 192936
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 192937, 192938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESAF", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192939)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

' *** 11/04/02 - INÍCIO Luiz G.F.Nogueira ***
Public Function TabelaPreco_Le_AliqICMS(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Lê os dados de uma tabela de preços a partir da alíquota de ICMS passada como parâmetro
'objTabelaPreco RECEBE(Input) a alíquota de ICMS que servirá de parâmetro para leitura
'objTabelaPreco RETORNA(Output) os dados lidos

Dim lComando As Long
Dim lErro As Long
Dim tTabelasDePreco As typeTabelasDePreco

On Error GoTo Erro_TabelaPreco_Le_AliqICMS

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94998
    'Cyntia
    tTabelasDePreco.sDescricao = String(STRING_TABELAPRECO_DESCRICAO, 0)
    
    'Procura no BD uma tabela de preços com a alíquota de ICMS passada como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, AliquotaICMS FROM TabelasDePreco WHERE AliquotaICMS =?", tTabelasDePreco.iCodigo, tTabelasDePreco.sDescricao, tTabelasDePreco.dAliquotaICMS, objTabelaPreco.dAliquotaICMS)
    If lErro <> AD_SQL_SUCESSO Then gError 94999
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102000
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102001
    
    'Transfere os dados lidos para o obj
    With tTabelasDePreco
    
        objTabelaPreco.iCodigo = .iCodigo
        objTabelaPreco.sDescricao = .sDescricao
        objTabelaPreco.dAliquotaICMS = .dAliquotaICMS
    
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TabelaPreco_Le_AliqICMS = SUCESSO
    
    Exit Function
    
Erro_TabelaPreco_Le_AliqICMS:

    TabelaPreco_Le_AliqICMS = gErr
    
    Select Case gErr
    
        Case 94998
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94999, 102000
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GENERICO", gErr, "TabelasDePreco")
            
        Case 102001 'SEM DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
' *** 11/04/02 - FIM Luiz G.F.Nogueira ***

