VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMOTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigoMO As Long
Private mvariTipoMO As Integer

Public Property Let iTipoMO(ByVal vData As Integer)
    mvariTipoMO = vData
End Property

Public Property Get iTipoMO() As Integer
    iTipoMO = mvariTipoMO
End Property

Public Property Let lCodigoMO(ByVal vData As Long)
    mvarlCodigoMO = vData
End Property

Public Property Get lCodigoMO() As Long
    lCodigoMO = mvarlCodigoMO
End Property


