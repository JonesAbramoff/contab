VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassIPICodEnquadramento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodigo AS String
Private mvarsGrupoCST AS String
Private mvarsDescCompleta AS String
Private mvariTipoIPI AS Integer

Public Property Let sCodigo (ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo () AS String
    sCodigo= mvarsCodigo
End Property

Public Property Let sGrupoCST (ByVal vData As String)
    mvarsGrupoCST = vData
End Property

Public Property Get sGrupoCST () AS String
    sGrupoCST= mvarsGrupoCST
End Property

Public Property Let sDescCompleta (ByVal vData As String)
    mvarsDescCompleta = vData
End Property

Public Property Get sDescCompleta () AS String
    sDescCompleta= mvarsDescCompleta
End Property

Public Property Let iTipoIPI (ByVal vData As Integer)
    mvariTipoIPI = vData
End Property

Public Property Get iTipoIPI () AS Integer
    iTipoIPI= mvariTipoIPI
End Property
