VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariImprimeItemAItem As Integer
Private mvarsNatOpPadrao As String
Private mvarsMensagemCupom As String
Private mvariOperadorIgualVendedor As Integer
Private mvariTabelaPreco As Integer
Private mvariTelaVendaMP As Integer
Private mvarobjECF As New ClassECF
Private mvariAtualizacaoECF As Integer
Private mvarsSimboloMoeda As String
Private mvariHorarioVerao As Integer
Private mvarcolAliquotaICMS As New Collection
Private mvariCupomDescreveFormaPagto As Integer
Private mvarlEspacoEntreLinhas As Long
Private mvariLinhasEntreCupons As Integer
Private mvarsNumLimRO As String
Private mvarsTruncamentoArredondamento As String
Private mvariGerenteAutoriza As Integer
Private mvarsContaContabil As String
Private mvariSemTeclado As Integer
Private mvariVendedorObrigatorio As Integer
Private mvariAbreAposFechamento As Integer
Private mvarcolTiposMovtoCaixa As New ColTipoMovtoCaixa
Private mvarsFTPURL As String
Private mvarsFTPUserName As String
Private mvarsFTPPassword As String
Private mvarsFTPDiretorio As String
Private mvarlIntervaloTrans As Long
Private mvarsNomeArqParam As String
Private mvarcolModulo As AdmColModulo

Public Property Set colModulo(ByVal vData As AdmColModulo)
    Set mvarcolModulo = vData
End Property

Public Property Get colModulo() As AdmColModulo
    Set colModulo = mvarcolModulo
End Property

Public Property Let sNomeArqParam(ByVal vData As String)
    mvarsNomeArqParam = vData
End Property

Public Property Get sNomeArqParam() As String
    sNomeArqParam = mvarsNomeArqParam
End Property

Public Property Let lIntervaloTrans(ByVal vData As Long)
    mvarlIntervaloTrans = vData
End Property

Public Property Get lIntervaloTrans() As Long
    lIntervaloTrans = mvarlIntervaloTrans
End Property

Public Property Let sFTPDiretorio(ByVal vData As String)
    mvarsFTPDiretorio = vData
End Property

Public Property Get sFTPDiretorio() As String
    sFTPDiretorio = mvarsFTPDiretorio
End Property

Public Property Let sFTPPassword(ByVal vData As String)
    mvarsFTPPassword = vData
End Property

Public Property Get sFTPPassword() As String
    sFTPPassword = mvarsFTPPassword
End Property

Public Property Let sFTPUserName(ByVal vData As String)
    mvarsFTPUserName = vData
End Property

Public Property Get sFTPUserName() As String
    sFTPUserName = mvarsFTPUserName
End Property

Public Property Let sFTPURL(ByVal vData As String)
    mvarsFTPURL = vData
End Property

Public Property Get sFTPURL() As String
    sFTPURL = mvarsFTPURL
End Property

Public Property Set colTiposMovtoCaixa(ByVal vData As ColTipoMovtoCaixa)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTiposMovtoCaixa = Form1
    Set mvarcolTiposMovtoCaixa = vData
End Property


Public Property Get colTiposMovtoCaixa() As ColTipoMovtoCaixa
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTiposMovtoCaixa
    Set colTiposMovtoCaixa = mvarcolTiposMovtoCaixa
End Property

Public Property Let iAbreAposFechamento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariAbreAposFechamento = vData
End Property


Public Property Get iAbreAposFechamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iAbreAposFechamento = mvariAbreAposFechamento
End Property


Public Property Let iVendedorObrigatorio(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariVendedorObrigatorio = vData
End Property


Public Property Get iVendedorObrigatorio() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iVendedorObrigatorio = mvariVendedorObrigatorio
End Property

Public Property Let iSemTeclado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariSemTeclado = vData
End Property


Public Property Get iSemTeclado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iSemTeclado = mvariSemTeclado
End Property

Public Property Let sContaContabil(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
End Property

Public Property Let iGerenteAutoriza(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariGerenteAutoriza = vData
End Property


Public Property Get iGerenteAutoriza() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iGerenteAutoriza = mvariGerenteAutoriza
End Property

Public Property Let sTruncamentoArredondamento(ByVal vData As String)
Attribute sTruncamentoArredondamento.VB_Description = "Indica se o caixa deve arredondar ou truncar o total do cupom."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTruncamentoArredondamento = 5
    mvarsTruncamentoArredondamento = vData
End Property


Public Property Get sTruncamentoArredondamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTruncamentoArredondamento
    sTruncamentoArredondamento = mvarsTruncamentoArredondamento
End Property



Public Property Let sNumLimRO(ByVal vData As String)
Attribute sNumLimRO.VB_Description = "Número Limite de Resumo de Operação (Lote) para boletos manuais."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumLimRO = 5
    mvarsNumLimRO = vData
End Property


Public Property Get sNumLimRO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumLimRO
    sNumLimRO = mvarsNumLimRO
End Property



Public Property Let iLinhasEntreCupons(ByVal vData As Integer)
Attribute iLinhasEntreCupons.VB_Description = "Determina o númeri de linhas entre cupons."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariLinhasEntreCupons = vData
End Property


Public Property Get iLinhasEntreCupons() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iLinhasEntreCupons = mvariLinhasEntreCupons
End Property



Public Property Let lEspacoEntreLinhas(ByVal vData As Long)
Attribute lEspacoEntreLinhas.VB_Description = "Determina o espaço entre as linhas do cupom."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lEspacoEntreLinhas = 5
    mvarlEspacoEntreLinhas = vData
End Property


Public Property Get lEspacoEntreLinhas() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lEspacoEntreLinhas
    lEspacoEntreLinhas = mvarlEspacoEntreLinhas
End Property



Public Property Let iCupomDescreveFormaPagto(ByVal vData As Integer)
Attribute iCupomDescreveFormaPagto.VB_Description = "Determinará se no cupom virá escrita a forma de pagamento."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCupomDescreveFormaPagto = 5
    mvariCupomDescreveFormaPagto = vData
End Property


Public Property Get iCupomDescreveFormaPagto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCupomDescreveFormaPagto
    iCupomDescreveFormaPagto = mvariCupomDescreveFormaPagto
End Property



Public Property Set colAliquotaICMS(ByVal vData As Collection)
Attribute colAliquotaICMS.VB_Description = "Guarda as alíquotas dos produtos que serão utilizadas pela Loja"
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colAliquotaICMS = Form1
    Set mvarcolAliquotaICMS = vData
End Property


Public Property Get colAliquotaICMS() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colAliquotaICMS
    Set colAliquotaICMS = mvarcolAliquotaICMS
End Property



Public Property Let iHorarioVerao(ByVal vData As Integer)
Attribute iHorarioVerao.VB_Description = "Informa se a loja está em horário de verão."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iHorarioVerao = 5
    mvariHorarioVerao = vData
End Property


Public Property Get iHorarioVerao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iHorarioVerao
    iHorarioVerao = mvariHorarioVerao
End Property



Public Property Let sSimboloMoeda(ByVal vData As String)
Attribute sSimboloMoeda.VB_Description = "Símbolo da moeda utilizada. Não deve conter o $."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSimboloMoeda = 5
    mvarsSimboloMoeda = vData
End Property


Public Property Get sSimboloMoeda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSimboloMoeda
    sSimboloMoeda = mvarsSimboloMoeda
End Property



Public Property Let iAtualizacaoECF(ByVal vData As Integer)
Attribute iAtualizacaoECF.VB_Description = "Sequencial de atualização de ECF. Se a caixa estiver com um sequencial diferente saberá que deverá pegar as últimas configurações."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAtualizacaoECF = 5
    mvariAtualizacaoECF = vData
End Property


Public Property Get iAtualizacaoECF() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAtualizacaoECF
    iAtualizacaoECF = mvariAtualizacaoECF
End Property



Friend Property Set objECF(ByVal vData As ClassECF)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objECF = Form1
    Set mvarobjECF = vData
End Property


Friend Property Get objECF() As ClassECF
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objECF
    Set objECF = mvarobjECF
End Property




Public Property Let iTelaVendaMP(ByVal vData As Integer)
Attribute iTelaVendaMP.VB_Description = "Determina se a tela de venda utilizada será a de poucos ou a de muitos produtos. (VendaP ou VendaM)."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTelaVendaMP = 5
    mvariTelaVendaMP = vData
End Property


Public Property Get iTelaVendaMP() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTelaVendaMP
    iTelaVendaMP = mvariTelaVendaMP
End Property



Public Property Let iTabelaPreco(ByVal vData As Integer)
Attribute iTabelaPreco.VB_Description = "Determina a tabela de preço utilizada pela Loja."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTabelaPreco = 5
    mvariTabelaPreco = vData
End Property


Public Property Get iTabelaPreco() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTabelaPreco
    iTabelaPreco = mvariTabelaPreco
End Property



Public Property Let iOperadorIgualVendedor(ByVal vData As Integer)
Attribute iOperadorIgualVendedor.VB_Description = "Indica se o operador de caixa é o próprio vendedor."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOperadorIgualVendedor = 5
    mvariOperadorIgualVendedor = vData
End Property


Public Property Get iOperadorIgualVendedor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOperadorIgualVendedor
    iOperadorIgualVendedor = mvariOperadorIgualVendedor
End Property



Public Property Let sMensagemCupom(ByVal vData As String)
Attribute sMensagemCupom.VB_Description = "Mensagem que sairá mpressa no final do cupom fiscal."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMensagemCupom = 5
    mvarsMensagemCupom = vData
End Property


Public Property Get sMensagemCupom() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMensagemCupom
    sMensagemCupom = mvarsMensagemCupom
End Property



Public Property Let sNatOpPadrao(ByVal vData As String)
Attribute sNatOpPadrao.VB_Description = "Natureza Operação padrão da Loja"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNatOpPadrao = 5
    mvarsNatOpPadrao = vData
End Property


Public Property Get sNatOpPadrao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNatOpPadrao
    sNatOpPadrao = mvarsNatOpPadrao
End Property



Public Property Let iImprimeItemAItem(ByVal vData As Integer)
Attribute iImprimeItemAItem.VB_Description = "Determina se o cupom deverá ser impresso item a item ou de uma única vez ao final do pagamento."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iImprimeItemAItem = 5
    mvariImprimeItemAItem = vData
End Property


Public Property Get iImprimeItemAItem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iImprimeItemAItem
    iImprimeItemAItem = mvariImprimeItemAItem
End Property

Public Function Inicializa() As Long

Dim objConfiguraLoja As New ClassLoja
Dim lErro As Long

On Error GoTo Erro_Inicializa

    'Lê as configurações em LojaConfig
    lErro = CF("LojaConfig_Le", objConfiguraLoja)
    If lErro <> SUCESSO Then gError 80088

    mvariCupomDescreveFormaPagto = objConfiguraLoja.iCupomDescreveFormaPagto
    mvarlEspacoEntreLinhas = objConfiguraLoja.lEspacoEntreLinhas
    mvariLinhasEntreCupons = objConfiguraLoja.iLinhasEntreCupons
    mvariImprimeItemAItem = objConfiguraLoja.iImprimeItemAItem
    mvarsNatOpPadrao = objConfiguraLoja.sNatOpPadrao
    mvarsNumLimRO = objConfiguraLoja.sNumLimRO
    mvarsTruncamentoArredondamento = objConfiguraLoja.sTruncamentoArredondamento
    mvariOperadorIgualVendedor = objConfiguraLoja.iOperadorIgualVendedor
    mvariTabelaPreco = objConfiguraLoja.iTabelaPreco
    mvarsSimboloMoeda = objConfiguraLoja.sSimboloMoeda
    mvarsMensagemCupom = objConfiguraLoja.sMensagemCupom
    mvariHorarioVerao = objConfiguraLoja.iHorarioVerao
    mvariSemTeclado = objConfiguraLoja.iSemTeclado
    mvarsContaContabil = objConfiguraLoja.sContaContabil
    mvariGerenteAutoriza = objConfiguraLoja.iGerenteAutoriza
    mvariVendedorObrigatorio = objConfiguraLoja.iVendedorObrigatorio
    mvariAbreAposFechamento = objConfiguraLoja.iAbreAposFechamento
    mvarsFTPURL = objConfiguraLoja.sFTPURL
    mvarsFTPUserName = objConfiguraLoja.sFTPUserName
    mvarsFTPPassword = objConfiguraLoja.sFTPPassword
    mvarsFTPDiretorio = objConfiguraLoja.sFTPDiretorio
    mvarlIntervaloTrans = objConfiguraLoja.lIntervaloTrans
    
    Set mvarcolAliquotaICMS = objConfiguraLoja.colAliquotaICMS
    
    
    'lê os tipos de movimento de Caixa
    lErro = CF("TiposMovtoCaixa_Le_Todos", mvarcolTiposMovtoCaixa)
    If lErro <> SUCESSO And lErro <> 113575 Then gError 113576
    
    'se não encontrar nenhum-> erro
    If lErro = 113575 Then gError 113577

    Inicializa = SUCESSO

    Exit Function

Erro_Inicializa:

    Inicializa = gErr

    Select Case gErr

        Case 80088
        
        Case 113576
        
        Case 113577
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVTOCAIXA_VAZIA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150908)

    End Select

    Exit Function

End Function

Function Gravar_Trans(ByVal objLoja As ClassLoja, ByVal ColConfig As ColLojaConfig) As Long
'Grava no bd a configuração atual

Dim lErro As Long

On Error GoTo Erro_Gravar_Trans
    
    'Chama LojaConfig_Grava_Configs
    lErro = CF("ConfiguraLoja_GravarTrans", objLoja, ColConfig)
    If lErro <> SUCESSO Then gError 109359
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = gErr
    
    Select Case gErr
    
        Case 109359
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150909)
    
    End Select
        
    Exit Function
    
End Function
