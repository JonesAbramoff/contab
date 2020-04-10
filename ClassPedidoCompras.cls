VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPedidoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvariComprador As Integer
Private mvarsContato As String
Private mvariTipoDestino As Integer
Private mvarlFornCliDestino As Long
Private mvariFilialDestino As Integer
Private mvardtData As Date
Private mvardtDataEmissao As Date
Private mvardtDataEnvio As Date
Private mvardtDataAlteracao As Date
Private mvariCondicaoPagto As Integer
Private mvardOutrasDespesas As Double
Private mvardValorFrete As Double
Private mvardValorSeguro As Double
Private mvardValorDesconto As Double
Private mvardValorTotal As Double
Private mvardValorIPI As Double
Private mvarsObservacao As String
Private mvarsTipoFrete As String
Private mvariTransportadora As Integer
Private mvariProxSeqBloqueio As Integer
Private mvariTipoBaixa As Integer
Private mvarsMotivoBaixa As String
Private mvarsAlcada As String
Private mvarcolItens As New Collection
Private mvarcolBloqueiosPC As New Collection
Private mvardValorProdutos As Double
Private mvarlObservacao As Long
Private mvardtDataBaixa As Date
Private mvariTipoOrig As Integer
Private mvarlPedCotacao As Long
Private mvarlMaiorAtraso As Long
Private mvarlAtrasoAcumulado As Long
Private mvarlItensPCRecebidos As Long
Private mvariStatus As Integer
Private mvarsNomeFornFilial As String
Private mvarsCondPagto As String
Private mvarsComprador As String
Private mvarsFilialEmp As String
Private mvarsFornecedor As String
Private mvarsFilialForn As String
Private mvarcolNotasPedCompras As New Collection
Private mvariEmbalagem As Integer
Private mvariMoeda As Integer
Private mvardTaxa As Double
Private mvardCotacaoMoeda As Double
Private mvarsObsEmbalagem As String
Private mvardtDataFixa As Date
Private mvarlCodigoPV As Long
Private mvardtDataRefFluxo As Date
Private mvarobjInfoAdic As New ClassInfoAdic
Private mvarobjProjetoInfo As Object

Public iTabelaPreco As Integer

Private mvarbInclusao As Boolean
Private mvarbEnvio As Boolean
Private mvarbAprovacao As Boolean
Private mvarbBaixa As Boolean
Private mvarbRecebimento As Boolean
Private mvarsUsuReg As String
Private mvarsUsuRegEnvio As String
Private mvarsUsuRegAprov As String
Private mvarsUsuRegBaixa As String
Private mvarName As String
Private mvariSelecionado As Integer
Private mvardtDataRegAprov As Date

Public Property Let dtDataRegAprov(ByVal vData As Date)
    mvardtDataRegAprov = vData
End Property

Public Property Get dtDataRegAprov() As Date
    dtDataRegAprov = mvardtDataRegAprov
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Let sUsuRegBaixa(ByVal vData As String)
    mvarsUsuRegBaixa = vData
End Property

Public Property Get sUsuRegBaixa() As String
    sUsuRegBaixa = mvarsUsuRegBaixa
End Property

Public Property Let sUsuRegAprov(ByVal vData As String)
    mvarsUsuRegAprov = vData
End Property

Public Property Get sUsuRegAprov() As String
    sUsuRegAprov = mvarsUsuRegAprov
End Property

Public Property Let sUsuRegEnvio(ByVal vData As String)
    mvarsUsuRegEnvio = vData
End Property

Public Property Get sUsuRegEnvio() As String
    sUsuRegEnvio = mvarsUsuRegEnvio
End Property

Public Property Let sUsuReg(ByVal vData As String)
    mvarsUsuReg = vData
End Property

Public Property Get sUsuReg() As String
    sUsuReg = mvarsUsuReg
End Property

Public Property Let bRecebimento(ByVal vData As Boolean)
    mvarbRecebimento = vData
End Property

Public Property Get bRecebimento() As Boolean
    bRecebimento = mvarbRecebimento
End Property

Public Property Let bBaixa(ByVal vData As Boolean)
    mvarbBaixa = vData
End Property

Public Property Get bBaixa() As Boolean
    bBaixa = mvarbBaixa
End Property

Public Property Let bAprovacao(ByVal vData As Boolean)
    mvarbAprovacao = vData
End Property

Public Property Get bAprovacao() As Boolean
    bAprovacao = mvarbAprovacao
End Property

Public Property Let bEnvio(ByVal vData As Boolean)
    mvarbEnvio = vData
End Property

Public Property Get bEnvio() As Boolean
    bEnvio = mvarbEnvio
End Property

Public Property Let bInclusao(ByVal vData As Boolean)
    mvarbInclusao = vData
End Property

Public Property Get bInclusao() As Boolean
    bInclusao = mvarbInclusao
End Property

Public Property Set objInfoAdic(ByVal vData As ClassInfoAdic)
    Set mvarobjInfoAdic = vData
End Property

Public Property Get objInfoAdic() As ClassInfoAdic
    Set objInfoAdic = mvarobjInfoAdic
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let dtDataRefFluxo(ByVal vData As Date)
    mvardtDataRefFluxo = vData
End Property

Public Property Get dtDataRefFluxo() As Date
    dtDataRefFluxo = mvardtDataRefFluxo
End Property

Public Property Let lCodigoPV(ByVal vData As Long)
    mvarlCodigoPV = vData
End Property

Public Property Get lCodigoPV() As Long
    lCodigoPV = mvarlCodigoPV
End Property

Public Property Let dtDataFixa(ByVal vData As Date)
    mvardtDataFixa = vData
End Property

Public Property Get dtDataFixa() As Date
    dtDataFixa = mvardtDataFixa
End Property

Public Property Let sObsEmbalagem(ByVal vData As String)
    mvarsObsEmbalagem = vData
End Property

Public Property Get sObsEmbalagem() As String
    sObsEmbalagem = mvarsObsEmbalagem
End Property

Public Property Let dCotacaoMoeda(ByVal vData As Double)
    mvardCotacaoMoeda = vData
End Property

Public Property Get dCotacaoMoeda() As Double
    dCotacaoMoeda = mvardCotacaoMoeda
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Set colNotasPedCompras(ByVal vData As Object)
    Set mvarcolNotasPedCompras = vData
End Property

Public Property Get colNotasPedCompras() As Collection
    Set colNotasPedCompras = mvarcolNotasPedCompras
End Property

Public Property Let sFilialForn(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFilialForn = 5
    mvarsFilialForn = vData
End Property


Public Property Get sFilialForn() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFilialForn
    sFilialForn = mvarsFilialForn
End Property




Public Property Let sFornecedor(ByVal vData As String)
    mvarsFornecedor = vData
End Property

Public Property Get sFornecedor() As String
    sFornecedor = mvarsFornecedor
End Property

Public Property Let sFilialEmp(ByVal vData As String)
    mvarsFilialEmp = vData
End Property

Public Property Get sFilialEmp() As String
    sFilialEmp = mvarsFilialEmp
End Property

Public Property Let sComprador(ByVal vData As String)
    mvarsComprador = vData
End Property

Public Property Get sComprador() As String
    sComprador = mvarsComprador
End Property

Public Property Let sCondPagto(ByVal vData As String)
    mvarsCondPagto = vData
End Property

Public Property Get sCondPagto() As String
    sCondPagto = mvarsCondPagto
End Property

Public Property Let sNomeFornFilial(ByVal vData As String)
    mvarsNomeFornFilial = vData
End Property


Public Property Get sNomeFornFilial() As String
    sNomeFornFilial = mvarsNomeFornFilial
End Property


Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lItensPCRecebidos(ByVal vData As Long)
    mvarlItensPCRecebidos = vData
End Property

Public Property Get lItensPCRecebidos() As Long
    lItensPCRecebidos = mvarlItensPCRecebidos
End Property

Public Property Let lAtrasoAcumulado(ByVal vData As Long)
    mvarlAtrasoAcumulado = vData
End Property

Public Property Get lAtrasoAcumulado() As Long
    lAtrasoAcumulado = mvarlAtrasoAcumulado
End Property

Public Property Let lMaiorAtraso(ByVal vData As Long)
    mvarlMaiorAtraso = vData
End Property

Public Property Get lMaiorAtraso() As Long
    lMaiorAtraso = mvarlMaiorAtraso
End Property

Public Property Let iTipoOrig(ByVal vData As Integer)
    mvariTipoOrig = vData
End Property

Public Property Get iTipoOrig() As Integer
    iTipoOrig = mvariTipoOrig
End Property

Public Property Let lPedCotacao(ByVal vData As Long)
Attribute lPedCotacao.VB_Description = "código do Pedido de Cotação"
    mvarlPedCotacao = vData
End Property

Public Property Get lPedCotacao() As Long
    lPedCotacao = mvarlPedCotacao
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
Attribute dtDataBaixa.VB_Description = "Data da Baixa do Pedido de Compras."
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let lObservacao(ByVal vData As Long)
    mvarlObservacao = vData
End Property

Public Property Get lObservacao() As Long
    lObservacao = mvarlObservacao
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Set colBloqueiosPC(ByVal vData As Object)
    Set mvarcolBloqueiosPC = vData
End Property

Public Property Get colBloqueiosPC() As Collection
    Set colBloqueiosPC = mvarcolBloqueiosPC
End Property

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sAlcada(ByVal vData As String)
    mvarsAlcada = vData
End Property

Public Property Get sAlcada() As String
    sAlcada = mvarsAlcada
End Property

Public Property Let sMotivoBaixa(ByVal vData As String)
    mvarsMotivoBaixa = vData
End Property

Public Property Get sMotivoBaixa() As String
    sMotivoBaixa = mvarsMotivoBaixa
End Property

Public Property Let iTipoBaixa(ByVal vData As Integer)
    mvariTipoBaixa = vData
End Property

Public Property Get iTipoBaixa() As Integer
    iTipoBaixa = mvariTipoBaixa
End Property

Public Property Let iProxSeqBloqueio(ByVal vData As Integer)
    mvariProxSeqBloqueio = vData
End Property

Public Property Get iProxSeqBloqueio() As Integer
    iProxSeqBloqueio = mvariProxSeqBloqueio
End Property

Public Property Let iTransportadora(ByVal vData As Integer)
    mvariTransportadora = vData
End Property

Public Property Get iTransportadora() As Integer
    iTransportadora = mvariTransportadora
End Property

Public Property Let sTipoFrete(ByVal vData As String)
    mvarsTipoFrete = vData
End Property

Public Property Get sTipoFrete() As String
    sTipoFrete = mvarsTipoFrete
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dOutrasDespesas(ByVal vData As Double)
    mvardOutrasDespesas = vData
End Property

Public Property Get dOutrasDespesas() As Double
    dOutrasDespesas = mvardOutrasDespesas
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let dtDataAlteracao(ByVal vData As Date)
    mvardtDataAlteracao = vData
End Property

Public Property Get dtDataAlteracao() As Date
    dtDataAlteracao = mvardtDataAlteracao
End Property

Public Property Let dtDataEnvio(ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio() As Date
    dtDataEnvio = mvardtDataEnvio
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialDestino(ByVal vData As Integer)
    mvariFilialDestino = vData
End Property

Public Property Get iFilialDestino() As Integer
    iFilialDestino = mvariFilialDestino
End Property

Public Property Let lFornCliDestino(ByVal vData As Long)
    mvarlFornCliDestino = vData
End Property

Public Property Get lFornCliDestino() As Long
    lFornCliDestino = mvarlFornCliDestino
End Property

Public Property Let iTipoDestino(ByVal vData As Integer)
    mvariTipoDestino = vData
End Property

Public Property Get iTipoDestino() As Integer
    iTipoDestino = mvariTipoDestino
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let iComprador(ByVal vData As Integer)
    mvariComprador = vData
End Property

Public Property Get iComprador() As Integer
    iComprador = mvariComprador
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim objUsu As ClassUsuarios, sNome As String
Dim objComprador As ClassComprador

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                        
        Case "CodigoPedido"
            objMnemonicoValor.colValor.Add lCodigo
            
        Case "FilialEmpresa"
            objMnemonicoValor.colValor.Add iFilialEmpresa
            
        Case "Inclusao_PedCompra"
            objMnemonicoValor.colValor.Add bInclusao
            
        Case "Envio_PedCompra"
            objMnemonicoValor.colValor.Add bEnvio
            
        Case "Aprovacao_PedCompra"
            objMnemonicoValor.colValor.Add bAprovacao
        
        Case "Baixa_PedCompra"
            objMnemonicoValor.colValor.Add bBaixa
                       
        Case "Comprador"
        
            If iComprador <> 0 Then
                Set objComprador = New ClassComprador
                objComprador.iCodigo = iComprador
                lErro = CF("Comprador_Le", objComprador)
                If lErro <> SUCESSO And lErro <> 50064 Then gError ERRO_SEM_MENSAGEM
                sComprador = objComprador.sCodUsuario
            End If
        
            objMnemonicoValor.colValor.Add sComprador
            
        Case "UsuCadastro"
            objMnemonicoValor.colValor.Add sUsuReg
            
        Case "Aprovador"
            objMnemonicoValor.colValor.Add sUsuRegAprov
            
        Case "UsuEnvio"
            objMnemonicoValor.colValor.Add sUsuRegEnvio
             
        Case "UsuBaixa"
            objMnemonicoValor.colValor.Add sUsuRegBaixa
            
        Case "EmailComprador"
        
            Set objUsu = New ClassUsuarios
            sEmail = ""
            
            If iComprador <> 0 Then
                Set objComprador = New ClassComprador
                objComprador.iCodigo = iComprador
                lErro = CF("Comprador_Le", objComprador)
                If lErro <> SUCESSO And lErro <> 50064 Then gError ERRO_SEM_MENSAGEM
                
                objUsu.sCodUsuario = objComprador.sCodUsuario
                sEmail = objComprador.sEmail
            End If
            
            If Len(Trim(objUsu.sCodUsuario)) > 0 And Len(Trim(sEmail)) = 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If

            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailUsuCadastro"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuReg
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If

            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailAprovador"

            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegAprov
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailUsuEnvio"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuEnvio
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
             
        Case "EmailUsuBaixa"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegBaixa
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
            
        Case "NomeUsuCadastro"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuReg
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If

            objMnemonicoValor.colValor.Add sNome
            
        Case "NomeAprovador"

            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegAprov
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
            
        Case "NomeUsuEnvio"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuEnvio
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
             
        Case "NomeUsuBaixa"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegBaixa
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
              
        Case Else
            gError 178683

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 178683
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178684)

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    mvardtDataRegAprov = DATA_NULA
    mvardtDataRefFluxo = DATA_NULA
End Sub
