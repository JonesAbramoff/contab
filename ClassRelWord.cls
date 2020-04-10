VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'API para trabalhar com WORD
'Considera somente a criacao de somente
'um documento...

Option Explicit

Private gobjDoc As Object 'aponta p/documento do Word
Private bFaltaPularPagina As Boolean
Private bFaltaConfigurarPagina As Boolean

'Constante de opcao: Quebra de pagina
Private Const wdPageBreak = 7
Private Const wdStory = 6
Private Const wdMove = 0
Private Const wdExtend = 1
Private Const wdGoToLast = -1
Private Const wdGoToLine = 3
Private Const wdSelectionInlineShape = 7
Private Const wdSelectionIP = 1

Private Const wdOrientPortrait = 0
Private Const wdOrientLandscape = 1

Public Function CriaDoc() As Long
'Cria uma instancia do Word e um documento na mesma, que será o corrente

Dim objWord As Object

On Error GoTo Erro_CriaDoc

    'Call MsgBox("CriaDoc1")

    'se gobjDoc foi instanciado
    If Not (gobjDoc Is Nothing) Then Error 2012

    'Instanciando uma nova aplicacao do Word
    Set objWord = CreateObject("Word.Application")
    
    'objWord.Visible = True
                 
    'Criando um novo documento dentro da aplicacao
    Set gobjDoc = objWord.Documents.Add

    'Call MsgBox("CriaDoc2")

    CriaDoc = SUCESSO

    Exit Function

Erro_CriaDoc:

    CriaDoc = Err

    Select Case Err

        Case 2012 'Erro que indica que ja existe um documento instanciado
        'essa API so trabalha com 1 documento... a ser tratado na rotina chamadora
        
        Case Else
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
           
    End Select

    Exit Function

End Function

Public Function SalvaDoc(ByVal sNomeDoc As String) As Long
'Salva o documento corrente da aplicacao passada

Dim bGeraPDF As Boolean

On Error GoTo Erro_SalvaDoc

    'Call MsgBox("SalvaDoc1")

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1003
    
    'Verifica se o nome que o documento
    'Sera salvo nao esta vazio
    If Len(Trim(sNomeDoc)) = 0 Then Error 1004

'    'Call MsgBox(sNomeDoc)
    
'    If InStr(1, sNomeDoc, "DANFE") <> 0 Then
'
'        sNomeDoc = "c:\lixo\danfe.doc"
'
'        With gobjDoc
'            .Application.ActiveWindow.ActivePane.View.SeekView = 10
'            .Application.Selection.Font.Name = "Code128bWinLarge"
'            .Application.Selection.Font.Color = -603914241
'            .Application.Selection.TypeText Text:="123456111"
'            .EmbedTrueTypeFonts = True
'            .SaveSubsetFonts = False
'            .DoNotEmbedSystemFonts = False
'        End With
'
'    End If

    
    Select Case gobjDoc.Application.Version
    
        Case "2.0", "6.0", "7.0", "8.0", "8.5", "9.0", "10.0", "11.0"
            'Salva o documento
            gobjDoc.SaveAs FileName:=sNomeDoc
        
        Case Else
        
            bGeraPDF = False
            If UCase(Right(sNomeDoc, 3)) = "PDF" Or UCase(Right(sNomeDoc, 4)) = "PDF""" Then bGeraPDF = True
            If (UCase(Right(sNomeDoc, 3)) = "TMP" Or UCase(Right(sNomeDoc, 4)) = "TMP""") And (UCase(Right(gsExtensaoGerRelExp, 3)) = "PDF" Or UCase(Right(gsExtensaoGerRelExp, 4)) = "PDF""") Then bGeraPDF = True
            
            If bGeraPDF Then
                'Salva o documento
                gobjDoc.ExportAsFixedFormat OutputFileName:=Replace(sNomeDoc, """", ""), ExportFormat:=17
            Else
                gobjDoc.SaveAs FileName:=sNomeDoc, FileFormat:=0
            End If
        
    End Select
    
    'MsgBox "Version:" & gobjDoc.Application.Version & " NomeDoc: " & sNomeDoc

    'Call MsgBox("SalvaDoc2")

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
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

    End Select

    Exit Function

End Function

Public Function InsereFigura(ByVal sPathFigura As String) As Long
'Insere uma Figura no documento passado

On Error GoTo Erro_InsereFigura
      
    'Call MsgBox("InsereFigura1")
      
    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1002
    
    'gobjDoc.Application.selection.TypeText "AddPicture " & sPathFigura
    
    gobjDoc.Application.selection.InlineShapes.AddPicture FileName:=sPathFigura, _
        LinkToFile:=True, SaveWithDocument:=True

    'Call MsgBox("InsereFigura2")

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
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

    End Select

    Exit Function

End Function

Public Function PulaPagina() As Long
'Pula uma página após a inserção de uma figura...

On Error GoTo Erro_PulaPagina

    'Call MsgBox("PulaPagina1")

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1001

    'gobjDoc.Application.selection.TypeText "InsertBreak"

    'gobjDoc.Application.selection.InsertBreak Type:=wdPageBreak
    
    'Call MsgBox("PulaPagina2")
    
    PulaPagina = SUCESSO
    
    Exit Function

Erro_PulaPagina:

    PulaPagina = Err

    Select Case Err

       Case 1001
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
       
        Case Else
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

    End Select

    Exit Function

End Function

Public Function FechaDoc() As Long
'fecha o documento e a aplicacao do word

On Error GoTo Erro_FechaDoc

    'Call MsgBox("FechaDoc1")

    'Verifica se o documento foi criado...
    If gobjDoc Is Nothing Then Error 1000

    'Fecha o documento sem salvar
    gobjDoc.Application.Quit (0)
    
    Set gobjDoc = Nothing
    
    'Call MsgBox("FechaDoc2")
    
    FechaDoc = SUCESSO
    
    Exit Function
    
Erro_FechaDoc:

    FechaDoc = Err

    Select Case Err

        Case 1000
           MsgBox "Erro - Documento não criado.", vbOKOnly, Err

        Case Else
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

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

    'Call MsgBox("InsereFigura_E_PulaPagina1")

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
    
    'Call MsgBox("InsereFigura_E_PulaPagina2")
    
    InsereFigura_E_PulaPagina = SUCESSO

    Exit Function

Erro_InsereFigura_E_PulaPagina:

    InsereFigura_E_PulaPagina = Err
    
    Select Case Err

        Case 2020, 2021
           
        Case Else
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    bFaltaPularPagina = False
    bFaltaConfigurarPagina = True
End Sub

'Essa sub nao faz parte da API
'Ela eh so para testes...

'Private Sub Command1_Click()
'
'Dim objDoc As Object
'
'    Call CriaDoc(objDoc)
'    Call InsereFigura(objDoc, "c:\oi.gif")
'
'    Call PulaPagina(objDoc)
'    Call InsereFigura(objDoc, "c:\bird.gif")
'    Call SalvaDoc(objDoc, "c:\passaro.doc")
'    Call FechaDoc(objDoc)
'
'End Sub

Private Sub Class_Terminate()

    'Verifica se o documento foi criado e nao foi destruido
    If Not (gobjDoc Is Nothing) Then Call FechaDoc
    
End Sub

Sub VaiFim()
Dim myRange As Object

    gobjDoc.Application.selection.EndKey wdStory, wdMove
    
    Set myRange = gobjDoc.GoTo(wdGoToLine, wdGoToLast)

End Sub

Public Function ConfiguraPagina(ByVal sLandscape As String, ByVal lMargEsq As Long, ByVal lMargSup As Long, ByVal lMargDir As Long, ByVal lMargInf As Long) As Long

Dim dMargMin As Double
Dim dMargMinV As Double

On Error GoTo Erro_ConfiguraPagina
    
    'Call MsgBox("ConfiguraPagina1")

    If bFaltaConfigurarPagina Then

        'Verifica se o documento foi criado...
        If gobjDoc Is Nothing Then Error 1002
        
        dMargMin = 36 / 2.54
        'dMargMinV = 18 / 2.54
        dMargMinV = 36 / 2.54
        
        With gobjDoc.PageSetup
            .Orientation = IIf(sLandscape = "S", wdOrientLandscape, wdOrientPortrait)
            .LeftMargin = IIf(lMargEsq / 20 > dMargMin, lMargEsq / 20, dMargMin)
            .TopMargin = IIf(lMargSup / 20 > dMargMinV, lMargSup / 20, dMargMinV)
            .RightMargin = IIf(lMargDir / 20 > dMargMin, lMargDir / 20, dMargMin)
            .BottomMargin = IIf(lMargInf / 20 > dMargMinV, lMargInf / 20, dMargMinV)
        End With
    
        bFaltaConfigurarPagina = False
        
    End If
    
    'Call MsgBox("ConfiguraPagina2")
    
    ConfiguraPagina = SUCESSO

    Exit Function

Erro_ConfiguraPagina:

    ConfiguraPagina = Err

    Select Case Err

       Case 1002
            MsgBox "Erro - Documento não criado.", vbOKOnly, Err
       
        Case Else
           MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err

    End Select

    Exit Function

End Function

