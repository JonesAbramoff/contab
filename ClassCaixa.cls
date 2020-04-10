VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariFilialEmpresa As Integer
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvariStatus As Integer
Private mvardtDataInicial As Date
Private mvarlProxSeqMov As Long
'Private mvarsPOS As String
Private mvariTeclado As Integer
'Private mvarsContaContabil As String
'Private mvariLeitoraCheque As Integer
'Private mvarlTipoCaixa As Long
'Private mvariLeitoraCodigoBarras As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvariTEF As Integer
Private mvariBoletoManual As Integer
Private mvariPOS As Integer
Private mvariOrcamentoECF As Integer
Private mvariUltimoOperador As Integer
Private mvariSessaoStatus As Integer

Public Property Let iSessaoStatus(ByVal vData As Integer)
    mvariSessaoStatus = vData
End Property

Public Property Get iSessaoStatus() As Integer
    iSessaoStatus = mvariSessaoStatus
End Property

Public Property Let iUltimoOperador(ByVal vData As Integer)
    mvariUltimoOperador = vData
End Property

Public Property Get iUltimoOperador() As Integer
    iUltimoOperador = mvariUltimoOperador
End Property

Public Property Let iOrcamentoECF(ByVal vData As Integer)
    mvariOrcamentoECF = vData
End Property

Public Property Get iOrcamentoECF() As Integer
    iOrcamentoECF = mvariOrcamentoECF
End Property

Public Property Let iPOS(ByVal vData As Integer)
    mvariPOS = vData
End Property

Public Property Get iPOS() As Integer
    iPOS = mvariPOS
End Property

Public Property Let iBoletoManual(ByVal vData As Integer)
    mvariBoletoManual = vData
End Property

Public Property Get iBoletoManual() As Integer
    iBoletoManual = mvariBoletoManual
End Property

Public Property Let iTEF(ByVal vData As Integer)
    mvariTEF = vData
End Property

Public Property Get iTEF() As Integer
    iTEF = mvariTEF
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

'Public Property Let iLeitoraCodigoBarras(ByVal vData As Integer)
'    mvariLeitoraCodigoBarras = vData
'End Property
'
'Public Property Get iLeitoraCodigoBarras() As Integer
'    iLeitoraCodigoBarras = mvariLeitoraCodigoBarras
'End Property


'Public Property Let lTipoCaixa(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.lTipoCaixa = 5
'    mvarlTipoCaixa = vData
'End Property
'
'
'Public Property Get lTipoCaixa() As Long
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.lTipoCaixa
'    lTipoCaixa = mvarlTipoCaixa
'End Property



'Public Property Let iLeitoraCheque(ByVal vData As Integer)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.iLeitoraCheque = 5
'    mvariLeitoraCheque = vData
'End Property
'
'
'Public Property Get iLeitoraCheque() As Integer
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.iLeitoraCheque
'    iLeitoraCheque = mvariLeitoraCheque
'End Property


Public Property Let lProxSeqMov(ByVal vData As Long)
Attribute lProxSeqMov.VB_Description = "Número sequencial para os movimentos de caixa."
    mvarlProxSeqMov = vData
End Property

Public Property Get lProxSeqMov() As Long
    lProxSeqMov = mvarlProxSeqMov
End Property

'Public Property Let sContaContabil(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.sContaContabil = 5
'    mvarsContaContabil = vData
'End Property


'Public Property Get sContaContabil() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.sContaContabil
'    sContaContabil = mvarsContaContabil
'End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data em que esse caixa começou a trabalhar"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "Armazena o Status que a caixa está no momento."
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do caixa"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
Attribute sNomeReduzido.VB_Description = "Nome Reduzido do Caixa."
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial empresa a qual esse caixa pertence"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Código do Caixa"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

'Public Property Let sPOS(ByVal vData As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.sDescricao = 5
'    mvarsPOS = vData
'End Property
'
'
'Public Property Get sPOS() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.sDescricao
'    sPOS = mvarsPOS
'End Property


Public Property Let iTeclado(ByVal vData As Integer)
    mvariTeclado = vData
End Property

Public Property Get iTeclado() As Integer
    iTeclado = mvariTeclado
End Property


