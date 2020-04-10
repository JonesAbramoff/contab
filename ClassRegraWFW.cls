VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegraWFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsModulo As String
Private mvariTransacao As Integer
Private mvariItem As Integer
Private mvarsUsuario As String
Private mvarsRegra As String
Private mvardRegraRet As Double
Private mvariTipoBloqueio As Integer
Private mvarsEmailPara As String
Private mvarsEmailParaRet As String
Private mvarsEmailAssunto As String
Private mvarsEmailAssuntoRet As String
Private mvarsEmailMsg As String
Private mvarsEmailMsgRet As String
Private mvarsAvisoMsg As String
Private mvarsAvisoMsgRet As String
Private mvarsLogDoc As String
Private mvarsLogDocRet As String
Private mvarsLogMsg As String
Private mvarsLogMsgRet As String
Private mvarcolUsuarios As New Collection
Private mvardtDataUltExec As Date
Private mvardHoraUltExec As Double
Private mvarsRelModulo As String
Private mvarsRelNome As String
Private mvarsRelOpcao As String
Private mvarsBrowseModulo As String
Private mvarsBrowseNome As String
Private mvarsBrowseOpcao As String

Public iRelPorEmail As Integer
Public sRelSel As String
Public sRelAnexo As String
Public sRelSelRet As String
Public sRelAnexoRet As String

Public Property Let sBrowseOpcao(ByVal vData As String)
    mvarsBrowseOpcao = vData
End Property

Public Property Get sBrowseOpcao() As String
    sBrowseOpcao = mvarsBrowseOpcao
End Property

Public Property Let sBrowseNome(ByVal vData As String)
    mvarsBrowseNome = vData
End Property

Public Property Get sBrowseNome() As String
    sBrowseNome = mvarsBrowseNome
End Property

Public Property Let sBrowseModulo(ByVal vData As String)
    mvarsBrowseModulo = vData
End Property

Public Property Get sBrowseModulo() As String
    sBrowseModulo = mvarsBrowseModulo
End Property

Public Property Let sRelOpcao(ByVal vData As String)
    mvarsRelOpcao = vData
End Property

Public Property Get sRelOpcao() As String
    sRelOpcao = mvarsRelOpcao
End Property

Public Property Let sRelNome(ByVal vData As String)
    mvarsRelNome = vData
End Property

Public Property Get sRelNome() As String
    sRelNome = mvarsRelNome
End Property

Public Property Let sRelModulo(ByVal vData As String)
    mvarsRelModulo = vData
End Property

Public Property Get sRelModulo() As String
    sRelModulo = mvarsRelModulo
End Property

Public Property Let dHoraUltExec(ByVal vData As Double)
    mvardHoraUltExec = vData
End Property

Public Property Get dHoraUltExec() As Double
    dHoraUltExec = mvardHoraUltExec
End Property

Public Property Let dtDataUltExec(ByVal vData As Date)
    mvardtDataUltExec = vData
End Property

Public Property Get dtDataUltExec() As Date
    dtDataUltExec = mvardtDataUltExec
End Property

Public Property Set colUsuarios(ByVal vData As Object)
    Set mvarcolUsuarios = vData
End Property

Public Property Get colUsuarios() As Collection
    Set colUsuarios = mvarcolUsuarios
End Property

Public Property Let sLogMsgRet(ByVal vData As String)
    mvarsLogMsgRet = vData
End Property

Public Property Get sLogMsgRet() As String
    sLogMsgRet = mvarsLogMsgRet
End Property

Public Property Let sLogMsg(ByVal vData As String)
    mvarsLogMsg = vData
End Property

Public Property Get sLogMsg() As String
    sLogMsg = mvarsLogMsg
End Property

Public Property Let sLogDocRet(ByVal vData As String)
    mvarsLogDocRet = vData
End Property

Public Property Get sLogDocRet() As String
    sLogDocRet = mvarsLogDocRet
End Property

Public Property Let sLogDoc(ByVal vData As String)
    mvarsLogDoc = vData
End Property

Public Property Get sLogDoc() As String
    sLogDoc = mvarsLogDoc
End Property

Public Property Let sAvisoMsgRet(ByVal vData As String)
    mvarsAvisoMsgRet = vData
End Property

Public Property Get sAvisoMsgRet() As String
    sAvisoMsgRet = mvarsAvisoMsgRet
End Property

Public Property Let sAvisoMsg(ByVal vData As String)
    mvarsAvisoMsg = vData
End Property

Public Property Get sAvisoMsg() As String
    sAvisoMsg = mvarsAvisoMsg
End Property

Public Property Let sEmailMsgRet(ByVal vData As String)
    mvarsEmailMsgRet = vData
End Property

Public Property Get sEmailMsgRet() As String
    sEmailMsgRet = mvarsEmailMsgRet
End Property

Public Property Let sEmailMsg(ByVal vData As String)
    mvarsEmailMsg = vData
End Property

Public Property Get sEmailMsg() As String
    sEmailMsg = mvarsEmailMsg
End Property

Public Property Let sEmailAssuntoRet(ByVal vData As String)
    mvarsEmailAssuntoRet = vData
End Property

Public Property Get sEmailAssuntoRet() As String
    sEmailAssuntoRet = mvarsEmailAssuntoRet
End Property

Public Property Let sEmailAssunto(ByVal vData As String)
    mvarsEmailAssunto = vData
End Property

Public Property Get sEmailAssunto() As String
    sEmailAssunto = mvarsEmailAssunto
End Property

Public Property Let sEmailParaRet(ByVal vData As String)
    mvarsEmailParaRet = vData
End Property

Public Property Get sEmailParaRet() As String
    sEmailParaRet = mvarsEmailParaRet
End Property

Public Property Let sEmailPara(ByVal vData As String)
    mvarsEmailPara = vData
End Property

Public Property Get sEmailPara() As String
    sEmailPara = mvarsEmailPara
End Property

Public Property Let iTipoBloqueio(ByVal vData As Integer)
    mvariTipoBloqueio = vData
End Property

Public Property Get iTipoBloqueio() As Integer
    iTipoBloqueio = mvariTipoBloqueio
End Property

Public Property Let dRegraRet(ByVal vData As Double)
    mvardRegraRet = vData
End Property

Public Property Get dRegraRet() As Double
    dRegraRet = mvardRegraRet
End Property

Public Property Let sRegra(ByVal vData As String)
    mvarsRegra = vData
End Property

Public Property Get sRegra() As String
    sRegra = mvarsRegra
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sModulo(ByVal vData As String)
    mvarsModulo = vData
End Property

Public Property Get sModulo() As String
    sModulo = mvarsModulo
End Property




