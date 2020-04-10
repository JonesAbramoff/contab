VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ALTERAÇÕES INPAL

'1 - 09/2000 Marcio - Incluido os campos ValorOriginal e MotivoDiferenca na leitura das ParcelasPag

Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function NfsPag_Le_FaturaPagar(ByVal objTituloPagar As ClassTituloPagar, ByVal colNFPagVinculada As ColNFsPag) As Long
'Lê as Notas Fiscais vinculadas ao Título passado por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tNFPag As typeNFsPag

On Error GoTo Erro_NfsPag_Le_FaturaPagar

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18614

    'Seleciona as NF's vinculadas ao Título passado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Filial, NumIntDoc, NumNotaFiscal, DataEmissao, Status, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido FROM NFsPag WHERE Fornecedor = ? AND NumIntTitPag = ? AND Status <> ? ORDER BY DataEmissao", tNFPag.iFilialEmpresa, tNFPag.iFilial, tNFPag.lNumIntDoc, tNFPag.lNumNotaFiscal, tNFPag.dtDataEmissao, tNFPag.iStatus, tNFPag.dtDataVencimento, tNFPag.dValorTotal, tNFPag.dValorSeguro, tNFPag.dValorFrete, tNFPag.dOutrasDespesas, tNFPag.dValorProdutos, tNFPag.dValorIRRF, tNFPag.dValorICMS, tNFPag.dValorICMSSubst, tNFPag.iCreditoICMS, tNFPag.dValorIPI, tNFPag.iCreditoIPI, tNFPag.dPISRetido, tNFPag.dCOFINSRetido, tNFPag.dCSLLRetido, objTituloPagar.lFornecedor, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18620

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18622
    
    If lErro = AD_SQL_SEM_DADOS Then Error 26020

    'Loop de Notas Fiscais
    Do While lErro = AD_SQL_SUCESSO

        'Armazena NF em colNFPagVinculada
        With tNFPag
            colNFPagVinculada.Add .lNumIntDoc, giFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, .lNumNotaFiscal, .dtDataEmissao, .iStatus, objTituloPagar.lNumIntDoc, .dtDataVencimento, .dValorTotal, .dValorSeguro, .dValorFrete, .dOutrasDespesas, .dValorProdutos, .dValorIRRF, .dValorICMS, .dValorICMSSubst, .iCreditoICMS, .dValorIPI, .iCreditoIPI, .dPISRetido, .dCOFINSRetido, .dCSLLRetido
        End With

        'Busca a próxima NF vinculada
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18621

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    NfsPag_Le_FaturaPagar = SUCESSO

    Exit Function

Erro_NfsPag_Le_FaturaPagar:

    Select Case Err

        Case 18614
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 18620, 18621, 18622
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", Err)

        Case 26020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Sub ParcelasRec_Le_BorderoCobranca1(ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao, sSelect As String, iParamFilialEmpresa As Integer, iParamTipoCliente As Integer)
'auxiliar a ParcelasRec_Le_BorderoCobranca
'monta o SELECT dinamicamente
Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String

    sFields = " Clientes.NomeReduzido, Clientes.RazaoSocial, Clientes.CodPadraoCobranca"
    sFields = sFields & ", TitulosRec.NumTitulo, TitulosRec.FilialEmpresa, TitulosRec.Cliente, TitulosRec.Filial, TitulosRec.SiglaDocumento"
    sFields = sFields & ", ParcelasRec.NumIntDoc, ParcelasRec.NumParcela, ParcelasRec.DataVencimento, ParcelasRec.Saldo"

    sFrom = " FROM ParcelasRec, TitulosRec, Clientes, FiliaisClientes"
    sWhere = " WHERE ParcelasRec.DataVencimento <= ? AND ParcelasRec.Status = ? AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.Cliente = Clientes.Codigo AND ParcelasRec.Cobrador = ? AND ParcelasRec.CarteiraCobranca = ? AND Clientes.Codigo = FiliaisClientes.CodCliente AND TitulosRec.Filial = FiliaisClientes.CodFilial AND ( FiliaisClientes.CodCobrador = ? OR FiliaisClientes.CodCobrador = 0 )"

    'Se entrar como filial centralizadora da cobranca pode pegar titulos de todas as filiais,
    If giFilialEmpresa = gobjCR.iFilialCentralizadora Then
        iParamFilialEmpresa = 0
    Else 'Senao pegar titulos somente da filial corrente
        sWhere = sWhere & " AND TitulosRec.FilialEmpresa = ?"
        iParamFilialEmpresa = 1
    End If

    If objBorderoCobrancaEmissao.sSiglaEstado <> "" Then
        sWhere = sWhere & " AND FiliaisClientes.EnderecoCobranca = Enderecos.Codigo AND Enderecos.SiglaEstado = '" & objBorderoCobrancaEmissao.sSiglaEstado & "'"
        sFrom = sFrom & ", Enderecos"
    End If
    
    If objBorderoCobrancaEmissao.iTipoCliente <> 0 Then
        sWhere = sWhere & " AND Clientes.Tipo = ?"
        iParamTipoCliente = 1
    Else
        iParamTipoCliente = 0
    End If
    
    '??? modificado p/INPAL sOrderBy = " ORDER BY TitulosRec.Cliente, TitulosRec.Filial, TitulosRec.NumTitulo, ParcelasRec.NumParcela"
    sOrderBy = " ORDER BY TitulosRec.NumTitulo, ParcelasRec.DataVencimento, ParcelasRec.NumParcela"
    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub
