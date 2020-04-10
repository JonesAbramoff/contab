VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassConfiguraADM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariConfiguracaoOK As Integer
Private mvarcolModulosConfigurar As Collection

Public Property Set colModulosConfigurar(ByVal vData As Object)
Attribute colModulosConfigurar.VB_Description = "Coleção contendo o nome dos módulos que devem ser configurados"
    Set mvarcolModulosConfigurar = vData
End Property

Public Property Get colModulosConfigurar() As Collection
    Set colModulosConfigurar = mvarcolModulosConfigurar
End Property

Public Property Let iConfiguracaoOK(ByVal vData As Integer)
Attribute iConfiguracaoOK.VB_Description = "TRUE = configuracao foi feita com sucesso FALSE = em caso contrário"
    mvariConfiguracaoOK = vData
End Property

Public Property Get iConfiguracaoOK() As Integer
    iConfiguracaoOK = mvariConfiguracaoOK
End Property

Private Sub Class_Initialize()
    Set mvarcolModulosConfigurar = New Collection
End Sub

Private Sub Class_Terminate()
    Set mvarcolModulosConfigurar = Nothing
End Sub
