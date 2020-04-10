VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegApuracaoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocApuracao As Long
Private mvariTipoReg As Integer
Private mvarsDescricao As String
Private mvardtData As Date
Private mvardValor As Double
'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property




Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor do Item"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data que vai Fazer o Link Provisório"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Item"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipoReg(ByVal vData As Integer)
Attribute iTipoReg.VB_Description = "Codigo do Tipo de Registro de ICMS"
    mvariTipoReg = vData
End Property

Public Property Get iTipoReg() As Integer
    iTipoReg = mvariTipoReg
End Property

Public Property Let lNumIntDocApuracao(ByVal vData As Long)
Attribute lNumIntDocApuracao.VB_Description = "identifica a apuracao a que se refere este item"
    mvarlNumIntDocApuracao = vData
End Property

Public Property Get lNumIntDocApuracao() As Long
    lNumIntDocApuracao = mvarlNumIntDocApuracao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "NumIntDoc deste Item"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
