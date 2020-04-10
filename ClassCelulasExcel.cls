VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCelulasExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvardFonteTamanho As Double 'local copy
Private mvarbFonteItalico As Boolean 'local copy
Private mvarbFonteNegrito As Boolean 'local copy
Private mvarvValor As Variant 'local copy
Private mvariFonteCor As Integer 'local copy
Private mvariBordaCor As Integer 'local copy
Private mvarlBordaLargura As Long 'local copy
Private mvariCelulaCor As Integer 'local copy
Private mvariNumeroFormato As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarvAlinhamentoHorizontal As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarsFonte As String 'local copy
'local variable(s) to hold property value(s)
Private mvarbCelulaTamanhoAuto As Boolean 'local copy
Private mvarsNumberFormat As String

Public Property Let sNumberFormat(ByVal vData As String)
    mvarsNumberFormat = vData
End Property

Public Property Get sNumberFormat() As String
    sNumberFormat = mvarsNumberFormat
End Property

Public Property Let bCelulaTamanhoAuto(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bCelulaTamanhoAuto = 5
    mvarbCelulaTamanhoAuto = vData
End Property


Public Property Get bCelulaTamanhoAuto() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bCelulaTamanhoAuto
    bCelulaTamanhoAuto = mvarbCelulaTamanhoAuto
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



Public Property Let vAlinhamentoHorizontal(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vAlinhamentoHorizontal = 5
    mvarvAlinhamentoHorizontal = vData
End Property


Public Property Set vAlinhamentoHorizontal(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.vAlinhamentoHorizontal = Form1
    Set mvarvAlinhamentoHorizontal = vData
End Property


Public Property Get vAlinhamentoHorizontal() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vAlinhamentoHorizontal
    If IsObject(mvarvAlinhamentoHorizontal) Then
        Set vAlinhamentoHorizontal = mvarvAlinhamentoHorizontal
    Else
        vAlinhamentoHorizontal = mvarvAlinhamentoHorizontal
    End If
End Property



Public Property Let iNumeroFormato(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumeroFormato = 5
    mvariNumeroFormato = vData
End Property


Public Property Get iNumeroFormato() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumeroFormato
    iNumeroFormato = mvariNumeroFormato
End Property



Public Property Let iCelulaCor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCelulaCor = 5
    mvariCelulaCor = vData
End Property


Public Property Get iCelulaCor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCelulaCor
    iCelulaCor = mvariCelulaCor
End Property



Public Property Let lBordaLargura(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lBordaLargura = 5
    mvarlBordaLargura = vData
End Property


Public Property Get lBordaLargura() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lBordaLargura
    lBordaLargura = mvarlBordaLargura
End Property



Public Property Let iBordaCor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iBordaCor = 5
    mvariBordaCor = vData
End Property


Public Property Get iBordaCor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iBordaCor
    iBordaCor = mvariBordaCor
End Property



Public Property Let iFonteCor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFonteCor = 5
    mvariFonteCor = vData
End Property


Public Property Get iFonteCor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFonteCor
    iFonteCor = mvariFonteCor
End Property



Public Property Let vValor(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vValor = 5
    mvarvValor = vData
End Property


Public Property Set vValor(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.vValor = Form1
    Set mvarvValor = vData
End Property


Public Property Get vValor() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vValor
    If IsObject(mvarvValor) Then
        Set vValor = mvarvValor
    Else
        vValor = mvarvValor
    End If
End Property



Public Property Let bFonteNegrito(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bFonteNegrito = 5
    mvarbFonteNegrito = vData
End Property


Public Property Get bFonteNegrito() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bFonteNegrito
    bFonteNegrito = mvarbFonteNegrito
End Property



Public Property Let bFonteItalico(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bFonteItalico = 5
    mvarbFonteItalico = vData
End Property


Public Property Get bFonteItalico() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bFonteItalico
    bFonteItalico = mvarbFonteItalico
End Property



Public Property Let dFonteTamanho(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dFonteTamanho = 5
    mvardFonteTamanho = vData
End Property


Public Property Get dFonteTamanho() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dFonteTamanho
    dFonteTamanho = mvardFonteTamanho
End Property

Private Sub Class_Initialize()
    mvarvAlinhamentoHorizontal = -1
End Sub
