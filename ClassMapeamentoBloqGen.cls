VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMapeamentoBloqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariTipoTelaBloqueio As Integer
Private mvariClassePossuiFilEmp As Integer
Private mvarsNomeTelaEditaDocBloq As String
Private mvarsProjetoClasseDocBloq As String
Private mvarsNomeClasseDocBloq As String
Private mvarsClasseNomeCampoChave As String
Private mvarsNomeFuncLeDoc As String
Private mvarsNomeColecaoBloqDoc As String
Private mvarsNomeBrowseChave As String
Private mvarsNomeTabelaBloqueios As String
Private mvarsTabelaNomeCampoChave As String
Private mvariTabelaBloqPossuiTipoTela As Integer
Private mvarsNomeTelaTestaPermissao As String
Private mvarsNomeFuncLiberaCust As String
Private mvarsNomeViewLeBloqueios As String
Private mvarsClasseDocNomeQTD As String
Private mvarsClasseDocNomeQTDReservada As String
Private mvarsClasseDocNomeUM As String
Private mvarsClasseDocNomeColItem As String
Private mvariClasseDocQTDNoItem As Integer
Private mvarsClasseDocNomeProduto As String
Private mvarsClasseDocNomeValor As String
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sClasseDocNomeValor(ByVal vData As String)
    mvarsClasseDocNomeValor = vData
End Property

Public Property Get sClasseDocNomeValor() As String
    sClasseDocNomeValor = mvarsClasseDocNomeValor
End Property

Public Property Let sClasseDocNomeProduto(ByVal vData As String)
    mvarsClasseDocNomeProduto = vData
End Property

Public Property Get sClasseDocNomeProduto() As String
    sClasseDocNomeProduto = mvarsClasseDocNomeProduto
End Property

Public Property Let iClasseDocQTDNoItem(ByVal vData As Integer)
    mvariClasseDocQTDNoItem = vData
End Property

Public Property Get iClasseDocQTDNoItem() As Integer
    iClasseDocQTDNoItem = mvariClasseDocQTDNoItem
End Property

Public Property Let sClasseDocNomeColItem(ByVal vData As String)
    mvarsClasseDocNomeColItem = vData
End Property

Public Property Get sClasseDocNomeColItem() As String
    sClasseDocNomeColItem = mvarsClasseDocNomeColItem
End Property

Public Property Let sClasseDocNomeUM(ByVal vData As String)
    mvarsClasseDocNomeUM = vData
End Property

Public Property Get sClasseDocNomeUM() As String
    sClasseDocNomeUM = mvarsClasseDocNomeUM
End Property

Public Property Let sClasseDocNomeQTDReservada(ByVal vData As String)
    mvarsClasseDocNomeQTDReservada = vData
End Property

Public Property Get sClasseDocNomeQTDReservada() As String
    sClasseDocNomeQTDReservada = mvarsClasseDocNomeQTDReservada
End Property

Public Property Let sClasseDocNomeQTD(ByVal vData As String)
    mvarsClasseDocNomeQTD = vData
End Property

Public Property Get sClasseDocNomeQTD() As String
    sClasseDocNomeQTD = mvarsClasseDocNomeQTD
End Property

Public Property Let sNomeColecaoBloqDoc(ByVal vData As String)
    mvarsNomeColecaoBloqDoc = vData
End Property

Public Property Get sNomeColecaoBloqDoc() As String
    sNomeColecaoBloqDoc = mvarsNomeColecaoBloqDoc
End Property

Public Property Let sNomeFuncLeDoc(ByVal vData As String)
    mvarsNomeFuncLeDoc = vData
End Property

Public Property Get sNomeFuncLeDoc() As String
    sNomeFuncLeDoc = mvarsNomeFuncLeDoc
End Property

Public Property Let iTabelaBloqPossuiTipoTela(ByVal vData As Integer)
    mvariTabelaBloqPossuiTipoTela = vData
End Property

Public Property Get iTabelaBloqPossuiTipoTela() As Integer
    iTabelaBloqPossuiTipoTela = mvariTabelaBloqPossuiTipoTela
End Property

Public Property Let sTabelaNomeCampoChave(ByVal vData As String)
    mvarsTabelaNomeCampoChave = vData
End Property

Public Property Get sTabelaNomeCampoChave() As String
    sTabelaNomeCampoChave = mvarsTabelaNomeCampoChave
End Property

Public Property Let sNomeViewLeBloqueios(ByVal vData As String)
    mvarsNomeViewLeBloqueios = vData
End Property

Public Property Get sNomeViewLeBloqueios() As String
    sNomeViewLeBloqueios = mvarsNomeViewLeBloqueios
End Property

Public Property Let sProjetoClasseDocBloq(ByVal vData As String)
    mvarsProjetoClasseDocBloq = vData
End Property

Public Property Get sProjetoClasseDocBloq() As String
    sProjetoClasseDocBloq = mvarsProjetoClasseDocBloq
End Property

Public Property Let sNomeFuncLiberaCust(ByVal vData As String)
    mvarsNomeFuncLiberaCust = vData
End Property

Public Property Get sNomeFuncLiberaCust() As String
    sNomeFuncLiberaCust = mvarsNomeFuncLiberaCust
End Property

Public Property Let sNomeTelaTestaPermissao(ByVal vData As String)
    mvarsNomeTelaTestaPermissao = vData
End Property

Public Property Get sNomeTelaTestaPermissao() As String
    sNomeTelaTestaPermissao = mvarsNomeTelaTestaPermissao
End Property

Public Property Let sNomeBrowseChave(ByVal vData As String)
    mvarsNomeBrowseChave = vData
End Property

Public Property Get sNomeBrowseChave() As String
    sNomeBrowseChave = mvarsNomeBrowseChave
End Property

Public Property Let sClasseNomeCampoChave(ByVal vData As String)
    mvarsClasseNomeCampoChave = vData
End Property

Public Property Get sClasseNomeCampoChave() As String
    sClasseNomeCampoChave = mvarsClasseNomeCampoChave
End Property


Public Property Let sNomeClasseDocBloq(ByVal vData As String)
    mvarsNomeClasseDocBloq = vData
End Property

Public Property Get sNomeClasseDocBloq() As String
    sNomeClasseDocBloq = mvarsNomeClasseDocBloq
End Property

Public Property Let sNomeTelaEditaDocBloq(ByVal vData As String)
    mvarsNomeTelaEditaDocBloq = vData
End Property

Public Property Get sNomeTelaEditaDocBloq() As String
    sNomeTelaEditaDocBloq = mvarsNomeTelaEditaDocBloq
End Property

Public Property Let iClassePossuiFilEmp(ByVal vData As Integer)
    mvariClassePossuiFilEmp = vData
End Property

Public Property Get iClassePossuiFilEmp() As Integer
    iClassePossuiFilEmp = mvariClassePossuiFilEmp
End Property

Public Property Let iTipoTelaBloqueio(ByVal vData As Integer)
    mvariTipoTelaBloqueio = vData
End Property

Public Property Get iTipoTelaBloqueio() As Integer
    iTipoTelaBloqueio = mvariTipoTelaBloqueio
End Property

Public Property Let sNomeTabelaBloqueios(ByVal vData As String)
    mvarsNomeTabelaBloqueios = vData
End Property

Public Property Get sNomeTabelaBloqueios() As String
    sNomeTabelaBloqueios = mvarsNomeTabelaBloqueios
End Property



