VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLinhaCabecalhoExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsAlinhamento As String 'local copy
Private mvarsNegrito As String 'local copy
Private mvarsItalico As String 'local copy
Private mvarsSublinhado As String 'local copy
Private mvarsTexto As String 'local copy
Private mvarsExibePagina As String 'local copy
Private mvarsFonte As String 'local copy
Private mvariTamanhoFonte As Integer 'local copy
Private mvariSecao As Integer 'local copy
Private mvariLinha As Integer 'local copy
Private mvariEspacoLinha As Integer 'local copy
Public Property Let iEspacoLinha(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iEspacoLinha = 5
    mvariEspacoLinha = vData
End Property


Public Property Get iEspacoLinha() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iEspacoLinha
    iEspacoLinha = mvariEspacoLinha
End Property



Public Property Let iLinha(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinha = 5
    mvariLinha = vData
End Property


Public Property Get iLinha() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinha
    iLinha = mvariLinha
End Property



Public Property Let iSecao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSecao = 5
    mvariSecao = vData
End Property


Public Property Get iSecao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSecao
    iSecao = mvariSecao
End Property



Public Property Let iTamanhoFonte(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTamanhoFonte = 5
    mvariTamanhoFonte = vData
End Property


Public Property Get iTamanhoFonte() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTamanhoFonte
    iTamanhoFonte = mvariTamanhoFonte
End Property



Public Property Let sFonte(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFonte = 5
    mvarsFonte = vData
End Property


Public Property Get sFonte() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFonte
    sFonte = mvarsFonte
End Property



Public Property Let sExibePagina(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sExibePagina = 5
    mvarsExibePagina = vData
End Property


Public Property Get sExibePagina() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sExibePagina
    sExibePagina = mvarsExibePagina
End Property



Public Property Let sTexto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTexto = 5
    mvarsTexto = vData
End Property


Public Property Get sTexto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTexto
    sTexto = mvarsTexto
End Property



Public Property Let sSublinhado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSublinhado = 5
    mvarsSublinhado = vData
End Property


Public Property Get sSublinhado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSublinhado
    sSublinhado = mvarsSublinhado
End Property



Public Property Let sItalico(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sItalico = 5
    mvarsItalico = vData
End Property


Public Property Get sItalico() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sItalico
    sItalico = mvarsItalico
End Property



Public Property Let sNegrito(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNegrito = 5
    mvarsNegrito = vData
End Property


Public Property Get sNegrito() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNegrito
    sNegrito = mvarsNegrito
End Property



Public Property Let sAlinhamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAlinhamento = 5
    mvarsAlinhamento = vData
End Property


Public Property Get sAlinhamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAlinhamento
    sAlinhamento = mvarsAlinhamento
End Property



