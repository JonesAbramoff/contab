VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSegmento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Segmentos de conta, etc."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariNivel As Integer
Private mvariTipo As Integer
Private mvariTamanho As Integer
Private mvarsDelimitador As String
Private mvariPreenchimento As Integer
Private mvarsCodigo As String

Public Property Let sCodigo(ByVal vData As String)
Attribute sCodigo.VB_Description = "Campo ao qual se refere o segmento. Pode ser ""conta"",  ""simples"", ""ccl"", etc."
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iPreenchimento(ByVal vData As Integer)
    mvariPreenchimento = vData
End Property

Public Property Get iPreenchimento() As Integer
    iPreenchimento = mvariPreenchimento
End Property

Public Property Let sDelimitador(ByVal vData As String)
    mvarsDelimitador = vData
End Property

Public Property Get sDelimitador() As String
    sDelimitador = mvarsDelimitador
End Property

Public Property Let iTamanho(ByVal vData As Integer)
    mvariTamanho = vData
End Property

Public Property Get iTamanho() As Integer
    iTamanho = mvariTamanho
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property



Private Sub Class_Initialize()
    mvarsDelimitador = String(STRING_SEGMENTO_DELIMITADOR, 0)
End Sub
