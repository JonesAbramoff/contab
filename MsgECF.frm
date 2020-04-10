VERSION 5.00
Begin VB.Form MsgECF 
   Caption         =   "Comunicação Caixa - Caixa Central"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   Icon            =   "MsgECF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox PausarProcessamento 
      Caption         =   "Pausar o processamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   2
      Top             =   90
      Width           =   3300
   End
   Begin VB.CheckBox ExibirTodasMsgs 
      Caption         =   "Exibir todas as mensagens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   1
      Top             =   75
      Width           =   6585
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   2505
   End
   Begin VB.ListBox ListaMsg 
      Height          =   2205
      ItemData        =   "MsgECF.frx":014A
      Left            =   150
      List            =   "MsgECF.frx":014C
      TabIndex        =   0
      Top             =   450
      Width           =   12345
   End
End
Attribute VB_Name = "MsgECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gobjTransfECF As ClassTransfECF
Public gsMsg As String
Public gsMsg1 As String
Public gdtTime As Date
Public giUsouFTP As Integer

'os campos abaixo foram incluidos para poder tentar transmitir direto pela rede local os arquivos entre o caixa e o servidor
Private gsDirDadosECFSemFTP As String
Private gsDirDadosCCSemFTP As String
Private gsDirXMLSemFTP As String

Private gbExecutando As Boolean

Private gobjECFConfig As Object

Const NUM_MAX_TENTATIVAS = 2
Const TEMPO_EXPERA_UPLOAD_ARQ = 5000 'Em milissegundos

Private Function ArquivoCopiaPelaRede(ByVal sArqOrigem As String, ByVal sArqDestino As String, Optional ByVal bApagaExistente As Boolean = False) As Long

Dim lErro As Long
Dim objFSO As New FileSystemObject

On Error GoTo Erro_ArquivoCopiaPelaRede

    If UCase(sArqOrigem) <> UCase(sArqDestino) Then
    
        'copia da origem para o destino com um nome temporario
        Call objFSO.CopyFile(sArqOrigem, sArqDestino & ".tmp", True)
        
        'se o destino existir
        If objFSO.FileExists(sArqDestino) Then
        
            If bApagaExistente Then
                Call objFSO.DeleteFile(sArqDestino)
            Else
                Name sArqDestino As sArqDestino & ".old." & right(Str(CDbl(Time)), 10)
            End If
        
        End If
        
        If objFSO.GetFile(sArqOrigem).Size <> objFSO.GetFile(sArqDestino & ".tmp").Size Then gError 216236
        'testar tamanho
        
        'renomeia o arquivo destino com o nome definitivo
        Name sArqDestino & ".tmp" As sArqDestino
                
        If objFSO.GetFile(sArqOrigem).Size <> objFSO.GetFile(sArqDestino).Size Then gError 216237
        'testar tamanho
        
    End If
    
    ArquivoCopiaPelaRede = SUCESSO
    
    Exit Function

Erro_ArquivoCopiaPelaRede:

    ArquivoCopiaPelaRede = gErr

    Select Case gErr
    
        Case 216236
            Call Arquivo_Log_Grava("(" & CStr(gErr) & ") Falha na cópia do arquivo pela rede: SIZE " & sArqOrigem & " = " & CStr(objFSO.GetFile(sArqOrigem).Size) & " e " & sArqDestino & ".tmp" & " = " & CStr(objFSO.GetFile(sArqDestino & ".tmp").Size))

        Case 216237
            Call Arquivo_Log_Grava("(" & CStr(gErr) & ") Falha na cópia do arquivo pela rede: SIZE " & sArqOrigem & " = " & CStr(objFSO.GetFile(sArqOrigem).Size) & " e " & sArqDestino & " = " & CStr(objFSO.GetFile(sArqDestino).Size))

        Case Else
            Call Arquivo_Log_Grava("(216238) Falha na cópia do arquivo pela rede: " & CStr(gErr) & "-" & Err.Description)
            'Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163185)

    End Select

    Exit Function

End Function

Private Function ObterDirConfigurado(ByVal sSecao As String, ByVal sChave As String) As String

Dim sRetorno As String
Dim lTamanho As Long
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(sSecao, sChave, "", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = Trim(StringZ(sRetorno))
    
    If Len(sRetorno) <> 0 Then
    
        'se o diretorio nao for terminado por \  ===> acrescentar
        If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"

    End If
    
    ObterDirConfigurado = sRetorno

End Function

Private Sub Form_Load()

Dim sRetorno As String
Dim lTamanho As Long
Dim sRet As String
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'para nao deixar descarregar a classe para nao perder a conexao com o paf
    Set gobjECFConfig = New ClassECFConfig
    
    Call ShowWindow(Me.hWnd, SW_SHOWMINNOACTIVE)

    giFilialEmpresa = 0
    giCodEmpresa = 0
    giCodCaixa = 0

    lErro = CF_ECF("Carrega_Caixa_Config")
    If lErro <> SUCESSO Then gError 133518

    lErro = CF_ECF("AbreBDs_PAFECF")
    If lErro <> SUCESSO Then gError 214161

    giUsouFTP = 0

    'os campos abaixo foram incluidos para poder tentar transmitir direto pela rede local os arquivos entre o caixa e o servidor
    gsDirDadosECFSemFTP = ObterDirConfigurado(APLICACAO_DADOS, "DirDadosECFSemFTP")
    gsDirDadosCCSemFTP = ObterDirConfigurado(APLICACAO_DADOS, "DirDadosCCSemFTP")
    gsDirXMLSemFTP = ObterDirConfigurado(APLICACAO_DADOS, "DirXMLSemFTP")
    'fim da inclusao
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosECF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 133372
    
    gobjTransfECF.sDirDadosECF = sRetorno
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    gobjTransfECF.sDirDadosCCC = sRetorno
    
    'se o diretorio nao for terminado por \  ===> acrescentar
    If right(gobjTransfECF.sDirDadosECF, 1) <> "\" Then gobjTransfECF.sDirDadosECF = gobjTransfECF.sDirDadosECF & "\"
    
    sRet = Dir(left(gobjTransfECF.sDirDadosECF, Len(gobjTransfECF.sDirDadosECF) - 1), vbDirectory)
        
    'se o diretorio DirDadadosECF\back nao existir ==> cria
    If sRet = "" Then MkDir (left(gobjTransfECF.sDirDadosECF, Len(gobjTransfECF.sDirDadosECF) - 1))
    
    sRet = Dir(gobjTransfECF.sDirDadosECF & "back", vbDirectory)
    
    'se o diretorio DirDadadosECF\back nao existir ==> cria
    If sRet = "" Then MkDir (gobjTransfECF.sDirDadosECF & "back")

    gbExecutando = False
    
    gdtTime = Now

    If gobjTransfECF.lIntervaloTrans > 0 Then
        
        'intervalo = 1 minuto ==> dentro do Timer vai controlar o intervalo
        Timer1.Interval = 1
        
        'subtrai o intervalo da data atual para que
        'o timer seja disparado imediatamente ao se chamar a primeira vez esta rotina
        gdtTime = DateAdd("n", -gobjTransfECF.lIntervaloTrans, Now)

    End If


    Exit Sub
    
Erro_Form_Load:

    Select Case gErr

        Case 133372
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "DirDadosECF", APLICACAO_DADOS, NOME_ARQUIVO_CAIXA)

        Case 133574
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "DirDadosCC", APLICACAO_DADOS, NOME_ARQUIVO_CAIXA)

        Case 133396, 133518, 133557, 133558, 214161

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163178)

    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gbExecutando Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo Erro_Form_UnLoad

    Set gobjECFConfig = Nothing
    Set gobjTransfECF = Nothing
    
    Exit Sub
    
Erro_Form_UnLoad:
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163179)

    End Select

    Exit Sub
    
End Sub


Private Sub Timer1_Timer()

Dim lMinutos As Long
Dim lErro As Long

On Error GoTo Erro_Timer1_Timer

    If Not gbExecutando And PausarProcessamento.Value = vbUnchecked Then

        lMinutos = DateDiff("n", gdtTime, Now)
    
        If lMinutos >= gobjTransfECF.lIntervaloTrans Then
            
            gbExecutando = True
        
            Call Limpa_Lista
            
            gdtTime = Now
            
            lErro = Transmite_Arq1()
            If lErro <> SUCESSO Then gError 133397
            
            If Len(gobjTransfECF.sFTPURL) > 0 Then
            
                lErro = Upload_Arq()
                If lErro <> SUCESSO Then gError 133398
            
            End If
            
            gbExecutando = False
        
        End If
        
    End If
        
    Exit Sub

Erro_Timer1_Timer:

    gbExecutando = False
    
    Select Case gErr

        Case 133397, 133398

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163180)

    End Select

    Exit Sub

End Sub

Private Function Arquivo_Log_Grava(sMsg As String, Optional ByVal iExibirMsg As Integer = 0) As Long

Dim sNomeArq As String
Dim bAbriuArq As Boolean
Dim iFreeFile As Integer

On Error GoTo Erro_Arquivo_Log_Grava

    bAbriuArq = False

    If ExibirTodasMsgs.Value = MARCADO Or iExibirMsg = MARCADO Then

        ListaMsg.AddItem sMsg
        ListaMsg.ListIndex = ListaMsg.NewIndex

    End If

    sNomeArq = gobjTransfECF.sDirDadosECF & giCodEmpresa & "_" & giFilialEmpresa & "_" & giCodCaixa & "_" & CStr(Format(gdtDataHoje, "ddmmyy")) & (".log")
    
    iFreeFile = FreeFile()
    
    Open sNomeArq For Append As #iFreeFile
    bAbriuArq = True
    
    'Inseri no Arquivo
    Print #iFreeFile, sMsg
    
    'Fecha o Arquivo
    Close #iFreeFile

    Arquivo_Log_Grava = SUCESSO

    Exit Function
    
Erro_Arquivo_Log_Grava:

    Arquivo_Log_Grava = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 163182)

    End Select

    If bAbriuArq Then Close #iFreeFile

    Exit Function

End Function

Private Function Upload_Arq_Xml(ByVal sFTPDiretorio As String, ByVal objFTP As Object, ByVal sTabela As String, ByVal sCampoArqXml As String, ByVal sCampoTransmitido As String, ByVal sCampoChaveAcesso As String, ByRef bAcabou As Boolean) As Long

Dim lErro As Long
Dim sErro As String
Dim alComando(1 To 2) As Long, iPosBarra As Integer
Dim iIndice As Integer, sArqXml As String, lTransacao As Long
Dim bSubiu As Boolean, sMeio As String

On Error GoTo Erro_Upload_Arq_Xml

    bAcabou = False
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 201552
    Next

    sArqXml = String(255, 0)
    lErro = Comando_Executar(alComando(1), "SELECT " & sCampoArqXml & " FROM " & sTabela & " WHERE " & sCampoTransmitido & " = 0 AND " & sCampoArqXml & " <> '' ORDER BY " & sCampoChaveAcesso, sArqXml)
    If lErro <> AD_SQL_SUCESSO Then gError 201553
        
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201554

    If lErro <> AD_SQL_SEM_DADOS Then
    
        iPosBarra = InStrRev(sArqXml, "\")
        
        bSubiu = False
        If gsDirXMLSemFTP <> "" Then
            lErro = ArquivoCopiaPelaRede(sArqXml, gsDirXMLSemFTP & Mid(sArqXml, iPosBarra + 1))
            If lErro = SUCESSO Then
                bSubiu = True
                sMeio = "pela rede"
            End If
            'se der erro tem que seguir para tentar por ftp
        End If
        
        If Not bSubiu Then
            sMeio = "por ftp"
            For iIndice = 1 To NUM_MAX_TENTATIVAS
                lErro = objFTP.Upload_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, left(sArqXml, iPosBarra), Mid(sArqXml, iPosBarra + 1), sErro)
                If lErro = SUCESSO Or iIndice = NUM_MAX_TENTATIVAS Then Exit For
                Sleep (TEMPO_EXPERA_UPLOAD_ARQ)
            Next
            If lErro <> SUCESSO Then
                Call Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sArqXml & " NÃO FOI transmitido com sucesso.", MARCADO)
                gError 214526
            End If
        End If
        
        lErro = CF_ECF("Caixa_Transacao_Abrir", lTransacao)
        If lErro <> SUCESSO Then gError 204617

        lErro = Comando_Executar(alComando(2), "UPDATE " & sTabela & " SET " & sCampoTransmitido & " = 1 WHERE " & sCampoArqXml & " = ?", sArqXml)
        If lErro <> AD_SQL_SUCESSO Then gError 201558
        
        'Função que Executa o Encerramento da Sessão
        lErro = CF_ECF("Caixa_Transacao_Fechar", lTransacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sArqXml & " foi transmitido " & sMeio & " com sucesso.", MARCADO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else
    
        bAcabou = True
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    Upload_Arq_Xml = SUCESSO

    Exit Function

Erro_Upload_Arq_Xml:

    'MsgBox (CStr(gErr))
    
    Upload_Arq_Xml = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 214526, 201558, 201554
            Call Rotina_ErroECF(vbOKOnly, ERRO_UPLOAD_ARQUIVO_FTP1, gErr, sArqXml, sFTPDiretorio, sErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 163183)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function

End Function

Function Upload_Arq() As Long

Dim sArquivo As String
Dim sDir As String
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String
Dim bAcabou As Boolean, sDir2 As String
Dim iIndice As Integer
Dim bSubiu As Boolean
Dim iPosBarra As Integer, sMeio As String

On Error GoTo Erro_Upload_Arq

    sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
    If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
    
    Set objFTP = CreateObject("SGEUtil.FTP1")
        
    'envia xmls de nfce
    Do
    
        DoEvents
    
        If PausarProcessamento.Value = vbChecked Then GoTo InterrompidoPeloUsuario
                
        lErro = Upload_Arq_Xml(sFTPDiretorio, objFTP, "NFCeInfo", "ArqXml", "XmlTransmitido", "chNFe", bAcabou)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Loop Until bAcabou

    'envia xmls de cancelamento de nfce
    Do
    
        DoEvents
    
        If PausarProcessamento.Value = vbChecked Then GoTo InterrompidoPeloUsuario
                
        lErro = Upload_Arq_Xml(sFTPDiretorio, objFTP, "NFCeInfo", "CancArqXml", "CancXmlTransmitido", "chNFe", bAcabou)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Loop Until bAcabou

    'envia xmls de sat
    Do
    
        DoEvents
    
        If PausarProcessamento.Value = vbChecked Then GoTo InterrompidoPeloUsuario
                
        lErro = Upload_Arq_Xml(sFTPDiretorio, objFTP, "SATInfo", "ArqXml", "XmlTransmitido", "ChaveAcesso", bAcabou)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Loop Until bAcabou

    'envia xmls de cancelamento de sat
    Do
    
        DoEvents
    
        If PausarProcessamento.Value = vbChecked Then GoTo InterrompidoPeloUsuario
                
        lErro = Upload_Arq_Xml(sFTPDiretorio, objFTP, "SATInfo", "CancArqXml", "CancXmlTransmitido", "ChaveAcesso", bAcabou)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Loop Until bAcabou

    'envia arquivos .ccc (movimentos de caixa)
    sArquivo = Dir(gobjTransfECF.sDirDadosECF & giCodEmpresa & "_" & giFilialEmpresa & "_" & giCodCaixa & "_" & "*.ccc")
    
    Do While sArquivo <> ""
    
        DoEvents
    
        If PausarProcessamento.Value = vbChecked Then GoTo InterrompidoPeloUsuario
        
        bSubiu = False
        If gsDirDadosECFSemFTP <> "" Then
            lErro = ArquivoCopiaPelaRede(gobjTransfECF.sDirDadosECF & sArquivo, gsDirDadosECFSemFTP & sArquivo)
            If lErro = SUCESSO Then
                bSubiu = True
                sMeio = "pela rede"
            End If
            'se der erro tem que seguir para tentar por ftp
        End If
        
        If Not bSubiu Then
            sMeio = "por ftp"
            For iIndice = 1 To NUM_MAX_TENTATIVAS
                lErro = objFTP.Upload_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosECF, sArquivo, sErro)
                If lErro = SUCESSO Or iIndice = NUM_MAX_TENTATIVAS Then Exit For
                Sleep (TEMPO_EXPERA_UPLOAD_ARQ)
            Next
            If lErro <> SUCESSO Then
                Call Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & gobjTransfECF.sDirDadosECF & sArquivo & " NÃO FOI transmitido com sucesso.", MARCADO)
                gError 214526
            End If
        End If
        
        sDir2 = Dir(gobjTransfECF.sDirDadosECF & "back\" & sArquivo)
    
        If Len(sDir2) > 0 Then
            Kill gobjTransfECF.sDirDadosECF & "back\" & sArquivo
        End If
    
        'coloca o arquivo no diretorio DirDadosECF\back
        Name gobjTransfECF.sDirDadosECF & sArquivo As gobjTransfECF.sDirDadosECF & "back\" & sArquivo
        
        lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & gobjTransfECF.sDirDadosECF & sArquivo & " foi transmitido " & sMeio & " com sucesso.", MARCADO)
        If lErro <> SUCESSO Then gError 133637
        
        sArquivo = Dir(gobjTransfECF.sDirDadosECF & giCodEmpresa & "_" & giFilialEmpresa & "_" & giCodCaixa & "_" & "*.ccc")
            
    Loop
    
InterrompidoPeloUsuario:

    Upload_Arq = SUCESSO

    Exit Function

Erro_Upload_Arq:

    'MsgBox (CStr(gErr))
    
    Upload_Arq = gErr

    Select Case gErr

        Case 133637, ERRO_SEM_MENSAGEM

        Case 214526
            Call Rotina_ErroECF(vbOKOnly, ERRO_UPLOAD_ARQUIVO_FTP1, gErr, gobjTransfECF.sDirDadosECF & sArquivo, sFTPDiretorio & sArquivo, sErro)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 163183)

    End Select

    Exit Function

End Function

Private Function Transmite_Arq1() As Long

Dim lErro As Long
Dim sNomeArqGerado As String

On Error GoTo Erro_Transmite_Arq1

    lErro = CF_ECF("Transmitir_Arquivo", TRANSMISSAO_ARQ_BATCH, sNomeArqGerado)
    If lErro <> SUCESSO And lErro <> 53 And lErro <> 117565 Then gError 133393
    
    If lErro = 53 Then gError 133395
    
    If lErro <> SUCESSO Then
        lErro = Arquivo_Log_Grava(CStr(Now) & " - Nao ha arquivo a ser gerado no momento.", MARCADO)
        If lErro <> SUCESSO Then gError 133525
    Else
        lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sNomeArqGerado & " foi gerado.", MARCADO)
        If lErro <> SUCESSO Then gError 133525
    End If
    
    Transmite_Arq1 = SUCESSO

    Exit Function

Erro_Transmite_Arq1:

    Transmite_Arq1 = gErr

    Select Case gErr

        Case 133392 To 133394, 133525

        Case 133395
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ABERTO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163184)

    End Select

    Exit Function

End Function

Private Sub Limpa_Lista()

Dim iIndice As Integer

    If ListaMsg.ListCount > 100 Then
        For iIndice = ListaMsg.ListCount - 50 To 0 Step -1
            ListaMsg.RemoveItem (iIndice)
        Next
    End If

End Sub

Public Function Download_DadosCC() As Long

Dim lTeste As Long
Dim lPos1 As Long
Dim sDir As String
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String
Dim sArquivoSemDiretorio As String, bBaixou As Boolean, sMeio As String

On Error GoTo Erro_Download_DadosCC

    sArquivoSemDiretorio = giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC

    bBaixou = False
    If gsDirDadosCCSemFTP <> "" Then
        lErro = ArquivoCopiaPelaRede(gsDirDadosCCSemFTP & sArquivoSemDiretorio, gobjTransfECF.sDirDadosCCC & sArquivoSemDiretorio, True)
        If lErro = SUCESSO Then
            bBaixou = True
            sMeio = "pela rede"
        End If
        'se der erro tem que seguir para tentar por ftp
    End If
    
    If Not bBaixou Then
    
        sMeio = "por ftp"
    
        Set objFTP = CreateObject("SGEUtil.FTP1")
        
        sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
        If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
        lErro = objFTP.Download_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosCCC, sArquivoSemDiretorio, sErro)
        If lErro <> SUCESSO Then
            Call Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & gobjTransfECF.sDirDadosCCC & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC & " NÃO FOI baixado com sucesso.", MARCADO)
            'gError 133576 '??? deixar carregar pelo dadoscc atual
        End If
                        
    End If
                        
    lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sArquivoSemDiretorio & " foi baixado " & sMeio & " com sucesso.", MARCADO)
    If lErro <> SUCESSO Then gError 133638

    Download_DadosCC = SUCESSO

    Exit Function

Erro_Download_DadosCC:

    Download_DadosCC = gErr

    Select Case gErr

        Case 133573
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_FTP_NAO_ENCONTRADO, gErr, gobjTransfECF.sFTPDiretorio & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC)

        Case 133575
            Call Rotina_ErroECF(vbOKOnly, ERRO_COMUNICACAO_FTP, gErr)
        
        Case 133576
            Call Rotina_ErroECF(vbOKOnly, ERRO_DOWNLOAD_ARQUIVO_FTP1, gErr, gobjTransfECF.sFTPDiretorio & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC, gobjTransfECF.sDirDadosCCC & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC, sErro)
        
        Case 133638
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163185)

    End Select

    Exit Function

End Function

