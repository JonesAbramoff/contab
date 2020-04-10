VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelOOWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'API para trabalhar com Writer do OpenOffice
'Considera somente a criacao de somente
'um documento...

Option Explicit

Private gobjDoc As Object 'aponta p/documento do Writer
Private objServiceManager As Object, objDesktop As Object

Private bFaltaPularPagina As Boolean
Private bFaltaConfigurarPagina As Boolean

Public Function CriaDoc() As Long
'Cria uma instancia do Writer e um documento na mesma, que será o corrente

Dim Args(), document As Object, dispatcher As Object
Dim args1(0 To 1) As Object
'Dim objCoreReflection As Object

On Error GoTo Erro_CriaDoc

    'se gobjDoc foi instanciado
    If Not (gobjDoc Is Nothing) Then Error 2012

    'The service manager is always the starting point
    'If there is no office running then an office is started up
    Set objServiceManager = CreateObject("com.sun.star.ServiceManager")

'    'Create the CoreReflection service that is later used to create structs
'    Set objCoreReflection = objServiceManager.createInstance("com.sun.star.reflection.CoreReflection")

    'Create the Desktop
    Set objDesktop = objServiceManager.createInstance("com.sun.star.frame.Desktop")
    
    'Open a new empty writer document
    Set gobjDoc = objDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, Args)
    
'    Set dispatcher = objServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
'    Set document = gobjDoc.currentcontroller.Frame
'    document.getComponentWindow.setVisible (False)
    
'    Set args1(0) = MakePropertyValue("Hidden", True)
'    Set gobjDoc = objDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, args1)

'    Call Command1_Click
    CriaDoc = SUCESSO

    Exit Function

Erro_CriaDoc:

    CriaDoc = Err

    Select Case Err

        Case 2012 'Erro que indica que ja existe um documento instanciado
        'essa API so trabalha com 1 documento... a ser tratado na rotina chamadora
        
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly
           
    End Select

    Exit Function

End Function

Public Function SalvaDoc(ByVal sNomeDoc As String) As Long
'Salva o documento corrente da aplicacao passada

Dim mFileProperties(1) As Object
Dim SaveParam(1) As Object 'Parameters to save the doc
Dim bGeraPDF As Boolean

On Error GoTo Erro_SalvaDoc

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1003
    
    'Verifica se o nome que o documento
    'Sera salvo nao esta vazio
    If Len(Trim(sNomeDoc)) = 0 Then Error 1004
    
    bGeraPDF = False
    If UCase(Right(sNomeDoc, 3)) = "PDF" Or UCase(Right(sNomeDoc, 4)) = "PDF""" Then bGeraPDF = True
    If (UCase(Right(sNomeDoc, 3)) = "TMP" Or UCase(Right(sNomeDoc, 4)) = "TMP""") And (UCase(Right(gsExtensaoGerRelExp, 3)) = "PDF" Or UCase(Right(gsExtensaoGerRelExp, 4)) = "PDF""") Then bGeraPDF = True
   
    If bGeraPDF Then
    
        Set mFileProperties(0) = MakePropertyValue("FilterName", "writer_pdf_Export")
        
        Set mFileProperties(1) = MakePropertyValue("Overwrite", True)
        
        Call gobjDoc.storeToURL(ConvertToUrl(sNomeDoc), mFileProperties())

        
    Else
        Set mFileProperties(0) = MakePropertyValue("FilterName", "MS Word 97")
           
        Set mFileProperties(1) = MakePropertyValue("Overwrite", True)
    
        'Salva o documento
        Call gobjDoc.storeAsURL(ConvertToUrl(sNomeDoc), mFileProperties())
           
    End If
        
    SalvaDoc = SUCESSO

    Exit Function

Erro_SalvaDoc:

    SalvaDoc = Err

    Select Case Err
    
        Case 1003
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
        
        Case 1004
            MsgBox "Erro - Nome do arquivo a ser salvo é nulo.", vbOKOnly, Err
        
        Case 1005
            MsgBox "Erro - Nome do arquivo a ser salvo inválido.", vbOKOnly, Err
        
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    End Select

    Exit Function

End Function

Public Function InsereFigura(ByVal sPathFigura As String) As Long
'Insere uma Figura no documento passado

Dim dispatcher As Object, i
Dim args1(0 To 3) As Object, document As Object
Dim oCursor As Object
'Dim oGrphObj As Object

On Error GoTo Erro_InsereFigura
      
    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1002
    
    Set dispatcher = objServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
    Set document = gobjDoc.currentcontroller.Frame

'    Set oGrphObj = gobjDoc.createInstance("com.sun.star.text.GraphicObject")
'    Call oGrphObj.setPropertyValue("GraphicURL", ConvertToUrl(sPathFigura))
'    Set oCursor = gobjDoc.Text.CreateTextCursor()
'    Call gobjDoc.Text.insertTextContent(oCursor, oGrphObj, True)
'    Call oCursor.Gotoend(False)
'    'Move down a line
'    Call dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())
'
    
    Rem ----------------------------------------------------------------------
    Set args1(0) = MakePropertyValue("FileName", ConvertToUrl(sPathFigura))
    Set args1(1) = MakePropertyValue("FilterName", "<Todos os formatos>")
    Set args1(2) = MakePropertyValue("AsLink", False)
    Set args1(3) = MakePropertyValue("Style", "Figuras")

    Call dispatcher.executeDispatch(document, ".uno:InsertGraphic", "", 0, args1())

    Set oCursor = gobjDoc.Text.CreateTextCursor()
    Call oCursor.Gotoend(False)
    
    For i = 1 To 20
    
        'envia um ESC para deselecionar a figura
        Call dispatcher.executeDispatch(document, ".uno:Escape", "", 0, Array())
    
        'às vezes abria o dialogo de figura como se o ESC nao tivesse sido processado
        DoEvents
    
    Next
    
    InsereFigura = SUCESSO

    Exit Function

Erro_InsereFigura:

    InsereFigura = Err

    Select Case Err

       Case 1002
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
       
        Case 5152
            MsgBox "Erro na inserção da figura " & sPathFigura
            
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    
    End Select

    Exit Function

End Function

Public Function PulaPagina() As Long
'Pula uma página após a inserção de uma figura...

Dim document   As Object, oCursor As Object
Dim dispatcher As Object
Dim args2(0 To 3) As Object, i

On Error GoTo Erro_PulaPagina

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1001

    Set dispatcher = objServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
    Set document = gobjDoc.currentcontroller.Frame
    
    Set args2(0) = MakePropertyValue("Kind", 3)
    Set args2(1) = MakePropertyValue("TemplateName", "")
    Set args2(2) = MakePropertyValue("PageNumber", 0)
    
    Call dispatcher.executeDispatch(document, ".uno:InsertBreak", "", 0, args2())

    Set oCursor = gobjDoc.Text.CreateTextCursor()
    Call oCursor.Gotoend(False)
    
    For i = 1 To 20
    
        'envia um ESC para deselecionar a figura
        Call dispatcher.executeDispatch(document, ".uno:Escape", "", 0, Array())
    
        'às vezes abria o dialogo de figura como se o ESC nao tivesse sido processado
        DoEvents
    
    Next
    
'    gobjDoc.Application.Selection.InsertBreak Type:=wdPageBreak
    
    PulaPagina = SUCESSO
    
    Exit Function

Erro_PulaPagina:

    PulaPagina = Err

    Select Case Err

       Case 1001
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
       
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    End Select

    Exit Function

End Function

Public Function FechaDoc() As Long
'fecha o documento e a aplicacao do word

On Error GoTo Erro_FechaDoc

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1000

    gobjDoc.Close (True)
    
    'Call objDesktop.Terminate
    
    Set objDesktop = Nothing
    Set objServiceManager = Nothing
    
'    'Fecha o documento sem salvar
'    gobjDoc.Application.Quit (0)
    
    Set gobjDoc = Nothing
    
    FechaDoc = SUCESSO
    
    Exit Function
    
Erro_FechaDoc:

    FechaDoc = Err

    Select Case Err

        Case 1000
           MsgBox "Erro - Documento não criado.", vbOKOnly, Err

        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    End Select

    Exit Function

End Function

Public Function InsereFigura_E_PulaPagina(ByVal sPathFigura As String) As Long
'insere uma figura na pagina corrente e pula para proxima...
'foi feita para que o usuario da API nao precise chamar a funcao
'de pular toda vez q inserir uma figura (se ele nao chamar, ele
'ira sobrescrever a figura que inseriu anteriormente)

Dim lErro As Long

On Error GoTo Erro_InsereFigura_E_PulaPagina

    If bFaltaPularPagina Then
        'pula a pagina
        lErro = PulaPagina()
        If lErro <> SUCESSO Then Error 2021
        bFaltaPularPagina = False
    End If
    
    'insere a figura no documento
    lErro = InsereFigura(sPathFigura)
    If lErro <> SUCESSO Then Error 2020
    
    bFaltaPularPagina = True
    
    InsereFigura_E_PulaPagina = SUCESSO

    Exit Function

Erro_InsereFigura_E_PulaPagina:

    InsereFigura_E_PulaPagina = Err
    
    Select Case Err

        Case 2020, 2021
           
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    bFaltaPularPagina = False
    bFaltaConfigurarPagina = True
End Sub

'''Essa sub nao faz parte da API
'''Ela eh so para testes...
''
''Private Sub Command1_Click()
''
''Dim objDoc As Object
''
''    Call InsereFigura("c:\windows\areia.bmp")
''    Call PulaPagina
''    Call InsereFigura("c:\windows\web\bullet.gif")
''    Call SalvaDoc("c:\passaro.doc")
''    Call FechaDoc
''
''End Sub

Private Sub Class_Terminate()

    'Verifica se o documento foi criado e nao foi destruido
    If Not (gobjDoc Is Nothing) Then Call FechaDoc
    
End Sub

Public Function ConfiguraPagina(ByVal sLandscape As String, ByVal lMargEsq As Long, ByVal lMargSup As Long, ByVal lMargDir As Long, ByVal lMargInf As Long) As Long

Dim dMargMin As Double
Dim oText
Dim oCursor
'Dim oNewStyle
Dim Args()
Dim s
Dim oStyle, altura, largura
Dim lMargEsq2 As Long, lMargSup2 As Long, lMargDir2 As Long, lMargInf2 As Long

On Error GoTo Erro_ConfiguraPagina
      
    If bFaltaConfigurarPagina Then
    
        'Verifica se o documento foi criado...
        If gobjDoc Is Nothing Then Error 1002
        
        'init the main object
        Set oText = gobjDoc.GetText
        Set oCursor = oText.CreateTextCursor
        
        'getting the old style
        s = oCursor.PageStyleName
        Set oStyle = gobjDoc.StyleFamilies.getByName("PageStyles").getByName(s)
      
        If sLandscape = "S" Then
        
            oStyle.setPropertyValue "IsLandscape", True
            
            altura = oStyle.getPropertyValue("Height")
            largura = oStyle.getPropertyValue("Width")
            oStyle.setPropertyValue "Height", largura
            oStyle.setPropertyValue "Width", altura
            
            oStyle.setPropertyValue "LeftMargin", 1500
            oStyle.setPropertyValue "TopMargin", 800
            oStyle.setPropertyValue "RightMargin", 1500
            oStyle.setPropertyValue "BottomMargin", 1000
        
        Else
        
            oStyle.setPropertyValue "IsLandscape", False
        
            oStyle.setPropertyValue "LeftMargin", 1500
            oStyle.setPropertyValue "TopMargin", 1500
            oStyle.setPropertyValue "RightMargin", 1500
            oStyle.setPropertyValue "BottomMargin", 1500
        
        End If
        
'        dMargMin = 1500 ' = 0.5 cm
'
'        lMargEsq2 = TwipsParaCentesimoMM(lMargEsq)
'        lMargSup2 = TwipsParaCentesimoMM(lMargSup)
'        lMargDir2 = TwipsParaCentesimoMM(lMargDir)
'        lMargInf2 = TwipsParaCentesimoMM(lMargInf)
'
'        oStyle.setPropertyValue "LeftMargin", CLng(IIf(lMargEsq2 > dMargMin, lMargEsq2, dMargMin))
'        oStyle.setPropertyValue "TopMargin", CLng(IIf(lMargSup2 > dMargMin, lMargSup2, dMargMin))
'        oStyle.setPropertyValue "RightMargin", CLng(IIf(lMargDir2 > dMargMin, lMargDir2, dMargMin))
'        oStyle.setPropertyValue "BottomMargin", CLng(IIf(lMargInf2 > dMargMin, lMargInf2, dMargMin))
'
'        oStyle.setPropertyValue "LeftMargin", 0
'        oStyle.setPropertyValue "TopMargin", 0
'        oStyle.setPropertyValue "RightMargin", 0
'        oStyle.setPropertyValue "BottomMargin", 0
        
        bFaltaConfigurarPagina = False
        
    End If
    
    ConfiguraPagina = SUCESSO

    Exit Function

Erro_ConfiguraPagina:

    ConfiguraPagina = Err

    Select Case Err

       Case 1002
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
       
        Case Else
           MsgBox "Erro Fornecido Pelo VB. (" & Err & ")" & Error$, vbOKOnly

    End Select

    Exit Function

End Function

'Function createStruct(strTypeName)
'   Dim aStruct
'
'   Set classSize = objCoreReflection.forName(strTypeName)
'   classSize.CreateObject aStruct
'   Set createStruct = aStruct
'End Function

Function MakePropertyValue(cName, uValue) As Object
    
  Dim oPropertyValue As Object
  Dim oSM As Object
    
  Set oSM = CreateObject("com.sun.star.ServiceManager")
  Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oPropertyValue.Name = cName
  oPropertyValue.Value = uValue
      
  Set MakePropertyValue = oPropertyValue

End Function

'
'Converts a Ms Windows local pathname in URL (RFC 1738)
'Todo : UNC pathnames, more character conversions
'
Public Function ConvertToUrl(strFile) As String
'    strFile = Replace(strFile, "\", "/")
'    strFile = Replace(strFile, ":", "|")
'    strFile = Replace(strFile, " ", "%20")
'    strFile = "file:///" + strFile
'    ConvertToUrl = strFile
Dim strFileAux As String
    strFileAux = strFile
    strFileAux = Replace(strFileAux, "\", "/")
    strFileAux = Replace(strFileAux, ":", "|")
    strFileAux = Replace(strFileAux, " ", "%20")
    strFileAux = "file:///" + strFileAux
    strFileAux = Replace(strFileAux, """", "")
    ConvertToUrl = strFileAux
End Function

' Given a URL to an external graphic resource,
'  load that graphic permanently into this drawing document,
'  and return a new URL to the internal resource.
' The new URL can be used in place of the old URL.
Function LoadGraphicIntoDocument(oDoc As Object, cUrl As String, cInternalName As String) As String
Dim oBitmaps As Object, cNewUrl As String
   ' Get the BitmapTable from this drawing document.
   ' It is a service that maintains a list of bitmaps that are internal
   '  to the document.
   Set oBitmaps = oDoc.createInstance("com.sun.star.drawing.BitmapTable")
    
   ' Add an external graphic to the BitmapTable of this document.
   Call oBitmaps.insertByName(cInternalName, cUrl)
    
   ' Now ask for it back.
   ' What we get back is an different Url that points to a graphic
   '  which is inside this document, and remains with the document.
   cNewUrl = oBitmaps.getByName(cInternalName)
    
   LoadGraphicIntoDocument = cNewUrl
End Function
 

'Function MakePoint(ByVal x As Long, ByVal y As Long) As com.sun.star.awt.Point
'   oPoint = createUnoStruct("com.sun.star.awt.Point")
'   oPoint.x = x
'   oPoint.y = y
'   MakePoint = oPoint
'End Function
'
'Function MakeSize(ByVal width As Long, ByVal height As Long) As com.sun.star.awt.Size
'   oSize = createUnoStruct("com.sun.star.awt.Size")
'   oSize.width = width
'   oSize.height = height
'   MakeSize = oSize
'End Function
'
'Function MakeRectangleShape(oDoc As Object, Optional oPosition As com.sun.star.awt.Point, Optional oSize As com.sun.star.awt.Size) As com.sun.star.drawing.RectangleShape
'   oShape = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
'   If Not IsMissing(oPosition) Then
'      oShape.Position = oPosition
'   End If
'   If Not IsMissing(oSize) Then
'      oShape.Size = oSize
'   End If
'   MakeRectangleShape = oShape
'End Function
'
'Function MakeGraphicObjectShape(oDoc As Object, Optional oPosition As com.sun.star.awt.Point, Optional oSize As com.sun.star.awt.Size) As com.sun.star.drawing.GraphicObjectShape
'   oShape = oDoc.createInstance("com.sun.star.drawing.GraphicObjectShape")
'   If Not IsMissing(oPosition) Then
'      oShape.Position = oPosition
'   End If
'   If Not IsMissing(oSize) Then
'      oShape.Size = oSize
'   End If
'   MakeGraphicObjectShape = oShape
'End Function
'
'
Private Function TwipsParaCentesimoMM(lTwips As Long) As Long
    TwipsParaCentesimoMM = (2540 * (lTwips / 20)) / 72
End Function
