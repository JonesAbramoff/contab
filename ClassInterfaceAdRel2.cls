VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInterfaceAdRel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Essa classe funciona como interface para acesso aos metodos de
'outras classes que estao contidas nesse projeto
'(por enquanto classRelEmail e classRelWord)
'
'isso serve para que os usuarios desse servico possam utiliza-lo
'sem ter conhecimento dos objs responsaveis pelo mesmo
'essa classe pode ser considerada como uma controladora, ou seja,
'a que recebe todas as requisicoes e, posteriormente,
'repassa para o objeto que deve executa-la

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private gobjRelWord As Object

'*********************************************************************
'*************** Funcoes Interface de ClassRelWord
'*********************************************************************

Public Function Inicia_Servidor_Word() As Long
'Essa funcao eh a responsavel por iniciar o servidor
'do word, ela basicamente instancia o obj
'caso seja chamada 2x, ignora a 2a chamada deixando de instanciar
'o obj
Dim lRetorno As Long
Dim sPgm As String
Dim sConteudo As String
Dim objADMLe As Object
Dim lErro As Long
Dim sExtensaoGerRelExp As String

On Error GoTo Erro_Inicia_Servidor_Word

    'se o objeto ja foi iniciado
    If gobjRelWord Is Nothing Then
        
        sPgm = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "Mail", "", sPgm, 255, "ADM100.INI")
        sPgm = Left(sPgm, lRetorno)

        If Len(Trim(sPgm)) = 0 Then
            sPgm = String(255, 0)
            lRetorno = GetPrivateProfileString("Forprint", "PgmOffice", "", sPgm, 255, "ADM100.INI")
            sPgm = Left(sPgm, lRetorno)
        End If
                
        sExtensaoGerRelExp = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "ExtensaoGerRelExp", "", sExtensaoGerRelExp, 255, "ADM100.INI")
        sExtensaoGerRelExp = Left(sExtensaoGerRelExp, lRetorno)
        gsExtensaoGerRelExp = sExtensaoGerRelExp
        If gsExtensaoGerRelExp = "" Then gsExtensaoGerRelExp = ".doc"

'        If Len(Trim(sPgm)) = 0 Then
'
'            Set objADMLe = CreateObject("RotinasADM.ADMSelect")
'
'            lErro = objADMLe.Config_Le("AdmConfig", "PGM_PADRAO_OFFICE", 0, sPgm)  'CF("Config_Le", "AdmConfig", "PGM_PADRAO_OFFICE", 0, sPgm)
'            If lErro <> SUCESSO Then Error 9999
'        End If
        
        'instancia o obj da classe ClassRelEmail
        Select Case UCase(sPgm)
        
            Case "OOWRITER", "OPENOFFICE", "OPEN OFFICE", "OO", "2"
                Set gobjRelWord = New ClassRelOOWriter
                        
            Case Else
                Set gobjRelWord = New ClassRelWord
                
        End Select
       
    End If
       
    Inicia_Servidor_Word = SUCESSO
       
    Exit Function

Erro_Inicia_Servidor_Word:

    Inicia_Servidor_Word = Err

    Select Case Err
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB.", vbOKOnly, Err

    End Select

    Exit Function
    
End Function
    
Public Function Finaliza_Servidor_Word() As Long
'Essa funcao eh a responsavel por finalizar o servidor
'ela simplismente libera o obj
'caso nao exista obj instanciado, ou seja, se o servidor
'nao esta iniciado, ela nao critica

On Error GoTo Erro_Finaliza_Servidor_Word

    'se o objeto ja foi iniciado
    Set gobjRelWord = Nothing
        
    Finaliza_Servidor_Word = SUCESSO
       
    Exit Function

Erro_Finaliza_Servidor_Word:

    Finaliza_Servidor_Word = Err

    Select Case Err
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select

    Exit Function
    
End Function

Public Function Cria_Doc() As Long
'Interface para a chamada do metodo "CriaDoc" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_Cria_Doc

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2000
    
    lErro = gobjRelWord.CriaDoc()
    If lErro <> SUCESSO Then Error 2001
    
    Cria_Doc = SUCESSO
    
    Exit Function
        
Erro_Cria_Doc:

    Cria_Doc = Err

    Select Case Err
    
        Case 2000
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2001
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

Public Function Salva_Doc(ByVal sNomeDoc As String) As Long
'Interface para a chamada do metodo "SalvaDoc" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_Salva_Doc

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2002
    
    lErro = gobjRelWord.SalvaDoc(Replace(sNomeDoc, """""", """"))
    If lErro <> SUCESSO Then Error 2003
    
    lErro = gobjRelWord.FechaDoc()
    If lErro <> SUCESSO Then Error 2009
    
    Salva_Doc = SUCESSO
    
    Exit Function
        
Erro_Salva_Doc:

    Salva_Doc = Err

    Select Case Err
    
        Case 2002
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2003
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

Public Function InsereFigura(ByVal sPathFigura As String) As Long
'Interface para a chamada do metodo "InsereFigura" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_InsereFigura

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2004
    
    lErro = gobjRelWord.InsereFigura(sPathFigura)
    If lErro <> SUCESSO Then Error 2005
    
    InsereFigura = SUCESSO
    
    Exit Function
        
Erro_InsereFigura:

    InsereFigura = Err

    Select Case Err
    
        Case 2004
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2005
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

Public Function PulaPagina() As Long
'Interface para a chamada do metodo "PulaPagina" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_PulaPagina

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2006
    
    lErro = gobjRelWord.PulaPagina()
    If lErro <> SUCESSO Then Error 2007
    
    PulaPagina = SUCESSO
    
    Exit Function
        
Erro_PulaPagina:

    PulaPagina = Err

    Select Case Err
    
        Case 2006
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2007
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function
   
Public Function FechaDoc() As Long
'Interface para a chamada do metodo "FechaDoc" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_FechaDoc

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2008
    
    lErro = gobjRelWord.FechaDoc()
    If lErro <> SUCESSO Then Error 2050
    
    FechaDoc = SUCESSO
    
    Exit Function
        
Erro_FechaDoc:

    FechaDoc = Err

    Select Case Err
    
        Case 2008
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2009
        
        Case 2050
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function
 
Public Function InsereFigura_E_PulaPagina(ByVal sPathFigura As String) As Long
'Interface para a chamada do metodo "PulaPagina" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_InsereFigura_E_PulaPagina

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2022
    
    lErro = gobjRelWord.InsereFigura_E_PulaPagina(sPathFigura)
    If lErro <> SUCESSO Then Error 2023
    
    InsereFigura_E_PulaPagina = SUCESSO
    
    Exit Function
        
Erro_InsereFigura_E_PulaPagina:

    InsereFigura_E_PulaPagina = Err

    Select Case Err
    
        Case 2022
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2023
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function
 
'*********************************************************************
'*************** FIM - Funcoes Interface de ClassRelWord
'*********************************************************************
 
'------------------------------------------------------------------------
 
'*********************************************************************
'*************** Funcoes Interface de ClassRelEmail
'*********************************************************************

Public Function Email_Enviar(ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sAliasAnexo As String, ByVal sCc As String, ByVal sOutrosAnexos As String, ByVal sCco As String) As Long
'Interface para a chamada do metodo "EmailEnviar" que esta
'implementado em ClassRelEmail

Dim lErro As Long
Dim objRelEmail As ClassRelEmail

On Error GoTo Erro_Email_Enviar

    'instancia um objeto da classe ClassRelEmail
    Set objRelEmail = New ClassRelEmail

    lErro = objRelEmail.Email_Enviar(sTo, sSubject, sTexto, sAnexo, sAliasAnexo, sCc, sOutrosAnexos, sCco)
    If lErro <> SUCESSO Then Error 2011
    
    Email_Enviar = SUCESSO
    
    Exit Function
        
Erro_Email_Enviar:

    Email_Enviar = Err

    Select Case Err
    
        Case 2011
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

Public Function Email_Enviar2(ByVal sSMPTHost As String, ByVal sFrom As String, ByVal sFromName As String, ByVal sUser As String, ByVal sPassword As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sCco As String) As Long
'Interface para a chamada do metodo "EmailEnviar" que esta
'implementado em ClassRelEmail

Dim lErro As Long
Dim objRelEmail As ClassRelEmail

On Error GoTo Erro_Email_Enviar2

    'instancia um objeto da classe ClassRelEmail
    Set objRelEmail = New ClassRelEmail

    lErro = objRelEmail.Email_Enviar2(sSMPTHost, sFrom, sFromName, sUser, sPassword, sTo, sSubject, sTexto, sAnexo, sCc, sCco)
    If lErro <> SUCESSO Then Error 2011
    
    Email_Enviar2 = SUCESSO
    
    Exit Function
        
Erro_Email_Enviar2:

    Email_Enviar2 = Err

    Select Case Err
    
        Case 2011
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function


Public Function ConfiguraPagina(ByVal sLandscape As String, ByVal lMargEsq As Long, ByVal lMargSup As Long, ByVal lMargDir As Long, ByVal lMargInf As Long) As Long
'Interface para a chamada do metodo "ConfiguraPagina" que esta
'implementado em ClassRelWord

Dim lErro As Long

On Error GoTo Erro_ConfiguraPagina

    'se o servidor nao estiver inicializado => erro
    If Not Esta_Inicializado_Servidor_Word Then Error 2022
    
    lErro = gobjRelWord.ConfiguraPagina(sLandscape, lMargEsq, lMargSup, lMargDir, lMargInf)
    If lErro <> SUCESSO Then Error 2023
    
    ConfiguraPagina = SUCESSO
    
    Exit Function
        
Erro_ConfiguraPagina:

    ConfiguraPagina = Err

    Select Case Err
    
        Case 2022
            'erro de servidor nao inicializado que deve
            'ser tratado na rotina chamadora
    
        Case 2023
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

Public Function Email_Enviar3(ByVal sSMPTHost As String, ByVal sFrom As String, ByVal sFromName As String, ByVal sUser As String, ByVal sPassword As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sCco As String, ByVal sReplyToAddress As String, ByVal bConfirmaEnvio As Boolean, Optional ByVal objClasseTrataEventos As Object, Optional ByVal sNomeProgresso As String, Optional ByVal sNomeFalha As String, Optional ByVal sNomeSucesso As String, Optional ByVal sNomeStatus As String) As Long
'Interface para a chamada do metodo "EmailEnviar" que esta
'implementado em ClassRelEmail

Dim lErro As Long
Dim objRelEmail As ClassRelEmail

On Error GoTo Erro_Email_Enviar3

    'instancia um objeto da classe ClassRelEmail
    Set objRelEmail = New ClassRelEmail

    lErro = objRelEmail.Email_Enviar3(sSMPTHost, sFrom, sFromName, sUser, sPassword, sTo, sSubject, sTexto, sAnexo, sCc, sCco, sReplyToAddress, bConfirmaEnvio, objClasseTrataEventos, sNomeProgresso, sNomeFalha, sNomeSucesso, sNomeStatus)
    If lErro <> SUCESSO Then Error 2011
    
    Email_Enviar3 = SUCESSO
    
    Exit Function
        
Erro_Email_Enviar3:

    Email_Enviar3 = Err

    Select Case Err
    
        Case 2011
    
        Case Else
            MsgBox "Erro Fornecido Pelo VB", vbOKOnly, Err

    End Select
    
    Exit Function

End Function

'*********************************************************************
'*************** FIM - Funcoes Interface de ClassRelEmail
'*********************************************************************

'------------------------------------------------------------------------

'*********************************************************************
'*************** Funcoes Privadas, para uso interno dessa classe
'*********************************************************************

Private Function Esta_Inicializado_Servidor_Word() As Boolean
'Essa funcao verifica se o servidor do Word esta inicializado, retornando
'verdadeiro ou falso caso esteja inicializado ou nao inicializado
'respectivamente

    Esta_Inicializado_Servidor_Word = True

    If gobjRelWord Is Nothing Then Esta_Inicializado_Servidor_Word = False

End Function

'*********************************************************************
'*************** FIM - Funcoes Privadas, para uso interno dessa classe
'*********************************************************************

