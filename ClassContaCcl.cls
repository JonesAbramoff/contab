VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassContaCcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCcl As String
Private mvarsConta As String
Private mvardSldIni As Double
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dSldIni(ByVal vData As Double)
Attribute dSldIni.VB_Description = "Saldo Inicial da associação conta x centro de custo/lucro quando o centro de custo/lucro é extra-contábil"
    mvardSldIni = vData
End Property

Public Property Get dSldIni() As Double
    dSldIni = mvardSldIni
End Property

Public Property Let sConta(ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Private Sub Class_Initialize()
    mvarsConta = String(STRING_CONTA, 0)
    mvarsCcl = String(STRING_CCL, 0)
End Sub
