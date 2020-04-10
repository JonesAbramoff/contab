VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassValorLiberadoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvariAno As Integer
Private mvaradValorLiberado(1 To 12) As Double

Public Property Let adValorLiberado(ByVal iIndice As Integer, ByVal vData As Double)
    mvaradValorLiberado(iIndice) = vData
End Property

Public Property Get adValorLiberado(ByVal iIndice As Integer) As Double
    adValorLiberado = mvaradValorLiberado(iIndice)
End Property

Public Property Let iAno(ByVal vData As Integer)
Attribute iAno.VB_Description = "Ano a que se referem as liberações"
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Codigo do Usuario que libera o crédito"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property



