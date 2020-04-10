VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassModuloFilEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodEmpresa As Long
Private mvariCodFilial As Integer
Private mvarsSiglaModulo As String
Private mvariConfigurado As Integer

Public Property Let iConfigurado(ByVal vData As Integer)
Attribute iConfigurado.VB_Description = "0 -> Nao configurado   1 -> Configurado"
    mvariConfigurado = vData
End Property

Public Property Get iConfigurado() As Integer
    iConfigurado = mvariConfigurado
End Property

Public Property Let sSiglaModulo(ByVal vData As String)
Attribute sSiglaModulo.VB_Description = "Sigla que representa o modulo. Ex: CTB/CPR/FAT..."
    mvarsSiglaModulo = vData
End Property

Public Property Get sSiglaModulo() As String
    sSiglaModulo = mvarsSiglaModulo
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
Attribute iCodFilial.VB_Description = "Codigo da Filial da Empresa - Filial 0 significa EMPRESA"
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodEmpresa(ByVal vData As Long)
Attribute lCodEmpresa.VB_Description = "Codigo da Empresa"
    mvarlCodEmpresa = vData
End Property

Public Property Get lCodEmpresa() As Long
    lCodEmpresa = mvarlCodEmpresa
End Property



