VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassMnemonicoMalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsMnemonico AS String
Private mvarsDescricao AS String
Private mvariTipoObj AS Integer
Private mvarsNomeCampoObj AS String
Private mvariTipo AS Integer

Public Property Let sMnemonico (ByVal vData As String)
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico () AS String
    sMnemonico= mvarsMnemonico
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let iTipoObj (ByVal vData As Integer)
    mvariTipoObj = vData
End Property

Public Property Get iTipoObj () AS Integer
    iTipoObj= mvariTipoObj
End Property

Public Property Let sNomeCampoObj (ByVal vData As String)
    mvarsNomeCampoObj = vData
End Property

Public Property Get sNomeCampoObj () AS String
    sNomeCampoObj= mvarsNomeCampoObj
End Property

Public Property Let iTipo (ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo () AS Integer
    iTipo= mvariTipo
End Property
