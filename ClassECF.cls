VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECF"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFabricante As Integer
Private mvariCodigo As Integer
Private mvariFilialEmpresa As Integer
Private mvariCaixa As Integer
Private mvariReinicioOperacao As Integer
Private mvariImpressoraCheque As Integer
Private mvariLeitoraCodigoBarras As Integer
Private mvarsPOS As String
Private mvariTipoTEF As Integer
Private mvariHorarioVerao As Integer
Private mvariAtualizacaoECF As Integer
Private mvarsNome_Fabricante As String
Private mvarsNome_Caixa As String
Private mvariImpressoraECF As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvariCodModeloECF As Integer

Public Property Let iCodModeloECF(ByVal vData As Integer)
    mvariCodModeloECF = vData
End Property

Public Property Get iCodModeloECF() As Integer
    iCodModeloECF = mvariCodModeloECF
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iImpressoraECF(ByVal vData As Integer)
    mvariImpressoraECF = vData
End Property

Public Property Get iImpressoraECF() As Integer
    iImpressoraECF = mvariImpressoraECF
End Property

Public Property Let sNome_Fabricante(ByVal vData As String)
    mvarsNome_Fabricante = vData
End Property

Public Property Get sNome_Fabricante() As String
    sNome_Fabricante = mvarsNome_Fabricante
End Property

Public Property Let sNome_Caixa(ByVal vData As String)
    mvarsNome_Caixa = vData
End Property

Public Property Get sNome_Caixa() As String
    sNome_Caixa = mvarsNome_Caixa
End Property

Public Property Let iAtualizacaoECF(ByVal vData As Integer)
    mvariAtualizacaoECF = vData
End Property

Public Property Get iAtualizacaoECF() As Integer
    iAtualizacaoECF = mvariAtualizacaoECF
End Property

Public Property Let iHorarioVerao(ByVal vData As Integer)
    mvariHorarioVerao = vData
End Property

Public Property Get iHorarioVerao() As Integer
    iHorarioVerao = mvariHorarioVerao
End Property

Public Property Let iTipoTEF(ByVal vData As Integer)
    mvariTipoTEF = vData
End Property

Public Property Get iTipoTEF() As Integer
    iTipoTEF = mvariTipoTEF
End Property

Public Property Let sPOS(ByVal vData As String)
    mvarsPOS = vData
End Property

Public Property Get sPOS() As String
    sPOS = mvarsPOS
End Property

Public Property Let iLeitoraCodigoBarras(ByVal vData As Integer)
    mvariLeitoraCodigoBarras = vData
End Property

Public Property Get iLeitoraCodigoBarras() As Integer
    iLeitoraCodigoBarras = mvariLeitoraCodigoBarras
End Property

Public Property Let iImpressoraCheque(ByVal vData As Integer)
    mvariImpressoraCheque = vData
End Property

Public Property Get iImpressoraCheque() As Integer
    iImpressoraCheque = mvariImpressoraCheque
End Property

Public Property Let iReinicioOperacao(ByVal vData As Integer)
    mvariReinicioOperacao = vData
End Property

Public Property Get iReinicioOperacao() As Integer
    iReinicioOperacao = mvariReinicioOperacao
End Property

Public Property Let iCaixa(ByVal vData As Integer)
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iFabricante(ByVal vData As Integer)
    mvariFabricante = vData
End Property

Public Property Get iFabricante() As Integer
    iFabricante = mvariFabricante
End Property


'Comandos operacionais
'Serão em geral acionados das telas de Venda e de Pagamento
'______________________
Function EstadoImpressora(bFimPapel, bPoucoPapel, bErroRelogio, bErroImpressora, bErroComando, bCupomAberto, bNumParametrosInvalido, bMemoriaFiscalLotada, bErroRAMCMOS, bAliquotaNaoProgramada, bCapacidadeAliquotasEsgotada, bCancelamentoNaoPermitido, bCGCProprietarioNaoProgramado, bComandoNaoExecutado)
'Retorna os erros possíveis, se b... for True o erro correspondente aconteceu.

End Function
Function MonitoramentoLinhas()

End Function

Function AberturaCupom()
'coloca o cabeçalho do Cupom

End Function
Function VendaItem(objItemCF As ClassItemCupomFiscal)
'obs: ClassItemCF deve ter tudo do BD + nome reduzido do Produto
'Se entrar com unidade de medida em geral chama um comando do driver e sem a UM outro
'Converte de número para sigla T1, T2, etc... para usar na chamada do driver
'(deve criar funcao p/ obter aliquota a partir de Produto e
'local da FilialEmpresa- vamos fazer por enquanto so para ICMS)

End Function

Function CancelaItem(iItem As Integer)

End Function

Function CancelaCupom()

End Function

Function IniciaFechamento(Optional sTipoDesconto As String, Optional dDesconto As Double)
'sTipoDesconto --> "AP" -> Acréscimo Percentual  "AV" -> Acréscimo Valor "DP" -> Desconto Percentual  "DV" -> Desconto Valor
'dDesconto --> Valor do acréscimo/desconto
'Inicia o fechamento do cupom. Disparado quando passamos para o Pagamento no Sistema.

End Function

Function EfetuaPagamento(iFormaPagto As Integer, dValor As Double, Optional sDescFormaPagto As String)
'Tem tabela de FormaPagto com a descricao correspondente ao código
'A Configuracao do loja por FilialEmpresa determina se a descricao entra ou nao

End Function

Function TerminaFechamento()
'Puxa a mensagem promocional da tabela LojaConfig

End Function
Function AbreGaveta()

End Function

Function EstadoGaveta(iEstado As Integer)
'iEstado=0 -> gaveta aberta --> GAVETA_ABERTA=0
'iEstado=1 -> gaveta fechada --> GAVETA_FECHADA=1

End Function

Function NumeroCupom(lNumero As Long)
'Retorna o número do cupom que acabou de imprimir

End Function

'Comandos de ABERTURA ou FECHAMENTO de caixa ou eventualmente
'________________________________________

'Vários desses comandos tem parametros setados
'em LojaConfig (Configuração de Loja) e aqui deve verificar
'se esses parametros estão bem setados para a Impressora.
'Sugestão: Reprogramá-los em cada abertura de caixa para não
'precisar ficar verificando se mudou e guardar o valor antigo
'em cada ECF

Function AdicionaAliquotaTrib(dAliquota As Double)

Dim lErro As Long

    Select Case iFabricante
    
        Case FAB_BEMATECH
            Dim objBematech As New ClassBematech
            lErro = objBematech.AdicionaAliquotaTrib(dAliquota)
    
        Case FAB_SWEDA
            Dim objSweda As New ClassSweda
            
            '......... continua
        
    End Select



'lErro = ComandoECF("AdicionaAliquotaTrib", dAliquota)

End Function
Function RetornoAliquotas(ParamArray dAliquotas())
'Retorna aliquotas para conferir se há necessidade de alterar o cadastramento
'Deve DEPOIS comparar as aliquotas que estão no ECF com o conjunto de aliquotas
'de Loja (tabela ProdutoFilial) para TODOS os Produtos que podem ser vendidos.

End Function

Function HorarioVeraoAtiva()

End Function
Function HorarioVeraoDesativa()

End Function
Function TruncArredond(sTrunca As String)
'Se sTrunca="T" é truncamento se for "A" é arredondamento

End Function
Function EspacoEntreLinhas(iDots As Integer)

End Function
Function LinhasEntreCupons(iLinhas As Integer)

End Function

Function ProgramaFormasPagamento()
'No caso da Bematech em cada redução Z apaga as formas de pagto
'havendo necessidade de colocar a programação na ABERTURA CAIXA
'Verificar se para OUTROS fabricantes o mesmo ocorre.


'Programa as formas de pagamento
'Disparada automaticamente diariamente na ABERTURA do caixa
'Usa constantes públicas:
'PAGTO_DINHEIRO=1
'PAGTO_CHEQUE_VISTA=2
'PAGTO_CHEQUE_PRE=3
'PAGTO_CARTAO_VISTA=4
'PAGTO_CARTAO_PRAZO=5
'PAGTO_VALE_TICKET=6
'Usa tabela FormaPagto para pegar as descricoes

End Function


Function LeituraX()
'Chamado automaticamente na ABERTURA do caixa e a qualquer momento.

End Function

Function ReducaoZ()
'Reducao Z. Chamada no FECHAMENTO do ECF

End Function


'Comando para sangrias, suprimentos
'______________

Function Sangria(dValor As Double)
'Na bematech chama Comprovante Nao Fiscal Nao Vinculado com parametro "SA"

End Function

Function Suprimento(dValor As Double)
'Na bematech chama Comprovante Nao Fiscal Nao Vinculado com parametro "SU"

End Function

'Comando para emitir comprovante do TEF
'______________

Function AbreComprovanteNaoFiscal(iFormaPagto As Integer)

End Function

Function ImprimeComprovanteNaoFiscal(sComprovante As String)
'Funcao chamada quando obtemos o retorno do TEF. Antes tem de chamar a Abertura.

End Function

Function FechamentoComprovanteNaoFiscal()
'No caso Bematech usa Fechamento de Relatório Gerencial.

End Function

'Comandos eventuais
'________________

Function Reset()
'Resseta a impressora

End Function

Function LeituraMemoriaFiscalData(dtDataInicial As Date, dtDataFinal As Date)
'Faz a leitura da memoria fiscal entre as datas
'Contém o total de vendas diário nessa faixa de datas
End Function

Function LeituraMemoriaFiscalRed(lReducaoInicial As Long, lReducaoFinal As Long)
'Faz a leitura da memoria fiscal entre reduções

End Function
Function LeituraUltimaReducao(objReducaoZ As ClassReducaoZ)
'Lê dados da última redução Z. Caso a impressão da Reducao Z
'dê problema, poderá recorrer a essa leitura e imprimir usando relatório de Reducao Z.
End Function



'Comandos de Impressão de Cheque
'_______________________________
Function ImprimeCheque(sCodBanco, sValor, sFavorecido, sLocal, sDia, sMes, sAno)
'Imprime cheque com parametros passados. O favorecido será a FilialEmpresaLoja.
'O Código do Banco deve orientar a impressora no layout do cheque.

End Function
Function CancelaImpressaoCheque()
'Cancelamento da impressao caso o cheque não seja posicionado em const pública segundos.
'Quando comandar ImprimeCheque deve ativar um Timer que de segundo em segundo verifica
'status do Cheque para saber se ainda está aguardando inserção. Se estourar o tempo
'de aguardar cancela. Se status mudar para imprimindo desativa o timer.

End Function
Function StatusCheque(iStatus)
'iStatus -> 0 -> aguardando inserção   1 -> imprimindo
End Function

Private Sub Class_Initialize()
'Faz essa classe apontar para gobjLoja.objECF

End Sub

