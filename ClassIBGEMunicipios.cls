VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassIBGEMunicipios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodUF AS String
Private mvarsUF AS String
Private mvarsCodMunic AS String
Private mvarsMunic AS String

Public Property Let sCodUF (ByVal vData As String)
    mvarsCodUF = vData
End Property

Public Property Get sCodUF () AS String
    sCodUF= mvarsCodUF
End Property

Public Property Let sUF (ByVal vData As String)
    mvarsUF = vData
End Property

Public Property Get sUF () AS String
    sUF= mvarsUF
End Property

Public Property Let sCodMunic (ByVal vData As String)
    mvarsCodMunic = vData
End Property

Public Property Get sCodMunic () AS String
    sCodMunic= mvarsCodMunic
End Property

Public Property Let sMunic (ByVal vData As String)
    mvarsMunic = vData
End Property

Public Property Get sMunic () AS String
    sMunic= mvarsMunic
End Property
