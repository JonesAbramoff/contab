VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarsCcl As String
Private mvarsDescCcl As String
Private mvariTipoCcl As Integer
Private mvariAtivo As Integer

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iTipoCcl(ByVal vData As Integer)
    mvariTipoCcl = vData
End Property

Public Property Get iTipoCcl() As Integer
    iTipoCcl = mvariTipoCcl
End Property

Public Property Let sDescCcl(ByVal vData As String)
    mvarsDescCcl = vData
End Property

Public Property Get sDescCcl() As String
    sDescCcl = mvarsDescCcl
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Private Sub Class_Initialize()
    mvarsDescCcl = String(STRING_CCL_DESCRICAO, 0)
    mvarsCcl = String(STRING_CCL, 0)
End Sub
