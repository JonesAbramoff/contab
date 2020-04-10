VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCriaControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarsNome As String
Private mvarcolControles As New Collection
Private mvariTipo As Integer
Private mvariOrdem As Integer
Private mvarsTabela As String
Private mvarsClasse As String
Private mvarsTipo As String
Private mvarsFrame As String
Private mvarsGrid As String
Private mvarsScript(1 To 500) As String
Private mvarlTopUltimo As Long

Public Property Let lTopUltimo(ByVal vData As Long)
    mvarlTopUltimo = vData
End Property

Public Property Get lTopUltimo() As Long
    lTopUltimo = mvarlTopUltimo
End Property

Public Property Let sScript(ByVal iIndice As Integer, ByVal vData As String)
    mvarsScript(iIndice) = vData
End Property

Public Property Get sScript(ByVal iIndice As Integer) As String
    sScript = mvarsScript(iIndice)
End Property

Public Property Let sGrid(ByVal vData As String)
    mvarsGrid = vData
End Property

Public Property Get sGrid() As String
    sGrid = mvarsGrid
End Property

Public Property Let sFrame(ByVal vData As String)
    mvarsFrame = vData
End Property

Public Property Get sFrame() As String
    sFrame = mvarsFrame
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let sClasse(ByVal vData As String)
    mvarsClasse = vData
End Property

Public Property Get sClasse() As String
    sClasse = mvarsClasse
End Property

Public Property Let sTabela(ByVal vData As String)
    mvarsTabela = vData
End Property

Public Property Get sTabela() As String
    sTabela = mvarsTabela
End Property

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Set colControles(ByVal vData As Collection)
    Set mvarcolControles = vData
End Property

Public Property Get colControles() As Collection
    Set colControles = mvarcolControles
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property
