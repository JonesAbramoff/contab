VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MsgCC 
   Caption         =   "Comunicação Caixa Central - Caixa"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
   Icon            =   "MsgCC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   13380
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
      Left            =   4320
      TabIndex        =   3
      Top             =   30
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
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   5955
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   135
      TabIndex        =   1
      Top             =   2655
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox ListaMsg 
      Height          =   2205
      ItemData        =   "MsgCC.frx":014A
      Left            =   90
      List            =   "MsgCC.frx":014C
      TabIndex        =   0
      Top             =   330
      Width           =   12915
   End
   Begin VB.Timer Timer1 
      Left            =   90
      Top             =   2400
   End
End
Attribute VB_Name = "MsgCC"
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
Public iUpDown As Integer

Private gbExecutando As Boolean
Private gbPrimeiraVez As Boolean

Private Sub ExecutarSobDemanda(ByVal iUpDownAux As Integer)

    If Len(gobjTransfECF.sFTPDiretorio) > 0 Then
    
        Select Case iUpDownAux
        
            Case 1
                Call Upload_DadosCC
            
            Case 2
                Call Download_DadosBack
    
            Case 3
                Call Upload_DadosBack
    
            Case 4
                Call Download_DadosCCB
    
            Case 5
                Call Upload_DadosCCB
    
            Case 6
                Call Download_Arq
    
        End Select

    End If

End Sub

Private Sub PrimeiraVezNoTimer()

Dim sRetorno As String
Dim lTamanho As Long
Dim sRet As String
Dim lErro As Long
Dim objLojaConfig As New ClassLojaConfig
Dim sDirDadosCC As String

On Error GoTo Erro_PrimeiraVezNoTimer

    objLojaConfig.iFilialEmpresa = EMPRESA_TODA
    objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOS
    
    lErro = CF("LojaConfig_Le1", objLojaConfig)
    If lErro <> SUCESSO And lErro <> 126361 Then gError 133408
    
    'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
    If lErro = 126361 Then objLojaConfig.sConteudo = CurDir
    
    gobjTransfECF.sDirDadosECF = objLojaConfig.sConteudo
    
    'se o diretorio nao for terminado por \  ===> acrescentar
    If right(gobjTransfECF.sDirDadosECF, 1) <> "\" Then gobjTransfECF.sDirDadosECF = gobjTransfECF.sDirDadosECF & "\"
    
    sRet = Dir(left(gobjTransfECF.sDirDadosECF, Len(gobjTransfECF.sDirDadosECF) - 1), vbDirectory)
    
    'se o diretorio DirDadadosECF nao existir ==> cria
    If sRet = "" Then MkDir (left(gobjTransfECF.sDirDadosECF, Len(gobjTransfECF.sDirDadosECF) - 1))
    
    sRet = Dir(gobjTransfECF.sDirDadosECF & "back", vbDirectory)
    
    'se o diretorio DirDadadosECF\back nao existir ==> cria
    If sRet = "" Then MkDir (gobjTransfECF.sDirDadosECF & "back")
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Or giLocalOperacao = LOCALOPERACAO_BACKOFFICE Or giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
    
        objLojaConfig.iFilialEmpresa = EMPRESA_TODA
        objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOSCCBACK
        
        lErro = CF("LojaConfig_Le1", objLojaConfig)
        If lErro <> SUCESSO And lErro <> 126361 Then gError 133621
        
        'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
        If lErro = 126361 Then objLojaConfig.sConteudo = CurDir
        
        sDirDadosCC = objLojaConfig.sConteudo
    
        'se o diretorio nao for terminado por \  ===> acrescentar
        If right(sDirDadosCC, 1) <> "\" Then sDirDadosCC = sDirDadosCC & "\"
        
        gobjTransfECF.sDirDadosCCC = sDirDadosCC
        
        sRet = Dir(left(gobjTransfECF.sDirDadosCCC, Len(gobjTransfECF.sDirDadosCCC) - 1), vbDirectory)
        
        'se o diretorio DirDadosCC nao existir ==> cria
        If sRet = "" Then MkDir (left(gobjTransfECF.sDirDadosCCC, Len(gobjTransfECF.sDirDadosCCC) - 1))
        
        sRet = Dir(gobjTransfECF.sDirDadosCCC & "back", vbDirectory)
        
        'se o diretorio DirDadosCC\back nao existir ==> cria
        If sRet = "" Then MkDir (gobjTransfECF.sDirDadosCCC & "back")
    
    End If
    
    Exit Sub
    
Erro_PrimeiraVezNoTimer:

    Select Case gErr

        Case 133408, 133435, 133559, 133560, 133621

        Case 133587
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_ARQUIVO_CONFIG", gErr, "DirDadosCC", APLICACAO_DADOS, NOME_ARQUIVO_ADM)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163164)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim sRetorno As String
Dim lTamanho As Long
Dim sRet As String
Dim lErro As Long
Dim objLojaConfig As New ClassLojaConfig
Dim sDirDadosCC As String

On Error GoTo Erro_Form_Load

    Call ShowWindow(Me.hWnd, SW_SHOWMINNOACTIVE)

    giUsouFTP = 0

    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    gdtTime = Now

    gbPrimeiraVez = True
    gbExecutando = False
    
    Timer1.Interval = 5000

    Exit Sub
    
Erro_Form_Load:

    Select Case gErr

        Case 133408, 133435, 133559, 133560, 133621

        Case 133587
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_ARQUIVO_CONFIG", gErr, "DirDadosCC", APLICACAO_DADOS, NOME_ARQUIVO_ADM)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163164)

    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gbExecutando Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
On Error GoTo Erro_Form_UnLoad

    Set gobjTransfECF = Nothing

    Exit Sub
    
Erro_Form_UnLoad:
    
    Select Case gErr

        Case 133509

        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 163165)

    End Select

    Exit Sub
    
End Sub



Private Sub Timer1_Timer()

Dim lMinutos As Long
Dim lErro As Long, iUpDownAux As Integer

On Error GoTo Erro_Timer1_Timer

    If Not gbExecutando And PausarProcessamento.Value = vbUnchecked Then

        gbExecutando = True
        
        If gbPrimeiraVez Then
        
            gbPrimeiraVez = False
            Call PrimeiraVezNoTimer
            
            If iUpDown <> 0 Then
                
                iUpDownAux = iUpDown
                iUpDown = 0
                
                Call ExecutarSobDemanda(iUpDownAux)
        
            End If
            
            If gobjTransfECF.lIntervaloTrans = 0 Then
            
                'se a rotina for chamado com o intervalo zerado ===>
                'so transmite uma vez os arquivos que estao no diretorio.
                
                gbExecutando = False
                
                Unload Me
        
            End If
        
        Else
        
            lMinutos = DateDiff("n", gdtTime, Now)
        
            If iUpDown <> 0 Or lMinutos >= gobjTransfECF.lIntervaloTrans Then
            
                gdtTime = Now
                
                Call Limpa_Lista
                
                If iUpDown <> 0 Then
                    
                    iUpDownAux = iUpDown
                    iUpDown = 0
                    
                    Call ExecutarSobDemanda(iUpDownAux)
                    
                Else
                
                    If Len(gobjTransfECF.sFTPDiretorio) > 0 Then
                    
                        lErro = Download_Arq()
                        If lErro <> SUCESSO Then gError 133507
                    
                    End If
                    
                    lErro = Carrega_Arq()
                    If lErro <> SUCESSO Then gError 133508
                    
                    lErro = Gera_Arq_CCB()
                    If lErro <> SUCESSO Then gError 133585
            
                    If Len(gobjTransfECF.sFTPDiretorio) > 0 Then
            
                        lErro = Upload_DadosCCB()
                        If lErro <> SUCESSO Then gError 133595
                        
                    End If
            
                    If Len(gobjTransfECF.sFTPDiretorio) > 0 Then
                    
                        lErro = Download_DadosCCB()
                        If lErro <> SUCESSO Then gError 133604
                    
                    End If
            
                    lErro = Carrega_Arq_CCB()
                    If lErro <> SUCESSO Then gError 133605
        
                End If
                
            End If
        
        End If
        
        gbExecutando = False
    
    End If
        
    Exit Sub

Erro_Timer1_Timer:

    gbExecutando = False

    Select Case gErr

        Case 133506 To 133508, 133585, 133589, 133590, 133595, 133604, 133605, 133606

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163166)

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

    sNomeArq = gobjTransfECF.sDirDadosECF & "CARGACC" & CStr(Format(gdtDataHoje, "ddmmyy")) & (".log")
    
    iFreeFile = FreeFile()
    
    Open sNomeArq For Append As #iFreeFile
    bAbriuArq = True
    
    'Inseri no Arquivo
    Print #iFreeFile, sMsg
    
    'Fecha o Arquivo
    Close #iFreeFile
    bAbriuArq = False

    Arquivo_Log_Grava = SUCESSO

    Exit Function
    
Erro_Arquivo_Log_Grava:

    Arquivo_Log_Grava = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163168)

    End Select
    
    If bAbriuArq Then Close #iFreeFile

    Exit Function

End Function

Function Download_Arq() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim objLojaConfig As New ClassLojaConfig
Dim lPos As Long
Dim lPos1 As Long
Dim vMsg1 As Variant
Dim sFile As String
Dim sDir As String
Dim iResult As VbMsgBoxResult
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String
Dim sRetorno As String
Dim iPos1 As Integer
Dim iPos As Integer, sDirXml As String

On Error GoTo Erro_Download_Arq

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Or giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
    
        Set objFTP = CreateObject("SGEUtil.FTP1")
        
        sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
        If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
        lErro = objFTP.Download_CC(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosECF, sErro, ".ccc", sRetorno)
        If lErro <> SUCESSO Then
            Call Arquivo_Log_Grava(CStr(Now) & " - Houve erro na recepção de arquivos de movimentação dos caixas (ccc).", MARCADO)
            gError 214524
        End If
                            
        iPos1 = 1
                            
        If Len(sRetorno) = 0 Then
            iPos1 = 0
        Else
            iPos = InStr(sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
        End If
                            
        Do While iPos1 < Len(sRetorno)
        
            lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & Mid(sRetorno, iPos1, iPos + 1 - iPos1) & " foi baixado com sucesso.", MARCADO)
            If lErro <> SUCESSO Then gError 133639
            
            iPos1 = iPos + 2
            
            iPos = InStr(iPos1, sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
            
        Loop
        
        lErro = CF("NFe_Obtem_Dir_Xml", sDirXml)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If right(sDirXml, 1) <> "\" Then sDirXml = sDirXml & "\"

        lErro = objFTP.Download_CC(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, sDirXml, sErro, ".xml", sRetorno)
        If lErro <> SUCESSO Then
            Call Arquivo_Log_Grava(CStr(Now) & " - Houve erro na recepção de arquivos de movimentação dos caixas (xml).", MARCADO)
            gError 214524
        End If
                            
        iPos1 = 1
                            
        If Len(sRetorno) = 0 Then
            iPos1 = 0
        Else
            iPos = InStr(sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
        End If
                            
        Do While iPos1 < Len(sRetorno)
        
            lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & Mid(sRetorno, iPos1, iPos + 1 - iPos1) & " foi baixado com sucesso.", MARCADO)
            If lErro <> SUCESSO Then gError 133639
            
            iPos1 = iPos + 2
            
            iPos = InStr(iPos1, sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
            
        Loop
    
    End If

    Download_Arq = SUCESSO

    Exit Function

Erro_Download_Arq:

    Download_Arq = gErr

    Select Case gErr

        Case 133639, ERRO_SEM_MENSAGEM

        Case 214524
            Call Rotina_Erro(vbOKOnly, "ERRO_DOWNLOAD_CCC_FTP", gErr, sErro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163169)

    End Select

    Exit Function

End Function

Private Function Carrega_Arq() As Long

Dim lErro As Long
Dim lErro1 As Long
Dim sArq As String
Dim objBarraProgresso As Object
Dim colFile As New Collection
Dim sDir As String
Dim vFile As Variant
Dim iCarregou As Integer
Dim sFile As String
Dim iCarregouAlgo As Integer
Dim fs As New FileSystemObject


On Error GoTo Erro_Carrega_Arq

    'Set fs = CreateObject("Scripting.FileSystemObject")

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Or giLocalOperacao = LOCALOPERACAO_BACKOFFICE Or giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then

        Set objBarraProgresso = BarraProgresso
    
        iCarregou = 1
        iCarregouAlgo = 0
    
        Do While iCarregou = 1
    
            iCarregou = 0
    
            sFile = Dir(gobjTransfECF.sDirDadosECF & "*.ccc")
    
            Do While Len(sFile) > 0
            
                If PausarProcessamento.Value = vbChecked Then Exit Do
                
                DoEvents
                
                sArq = gobjTransfECF.sDirDadosECF & sFile
        
                'nao para de executar mesmo que tenha dado erro
                lErro = CF("Verifica_Nome_Arquivo1", sFile)
                
                'se o arquivo já foi processado
                If lErro = 133433 Then
                
                    'e já está na pasta back
                    If fs.FileExists(gobjTransfECF.sDirDadosECF & "back\" & sFile) Then Kill sArq
                
                End If
                
                If lErro = SUCESSO Then
                
                    BarraProgresso.Value = 0
                
                    'nao para de executar mesmo que tenha dado erro
                    lErro = CF("Rotina_Carga_ECF_Caixa_Central", sArq, objBarraProgresso, TRANSMISSAO_ARQ_BATCH)
                    If lErro = SUCESSO Then
                        
                        iCarregou = 1
                        iCarregouAlgo = 1
                                   
                        lErro = Arquivo_Log_Grava(CStr(Now) & " - Carregou o arquivo " & sArq, MARCADO)
                        If lErro <> SUCESSO Then gError 133438
                        
'                        sDir = Dir(gobjTransfECF.sDirDadosECF & "back\" & sFile)
    
'                        If Len(sDir) > 0 Then
                        If fs.FileExists(gobjTransfECF.sDirDadosECF & "back\" & sFile) Then
                            Kill gobjTransfECF.sDirDadosECF & "back\" & sFile
                        End If
    
                        Name gobjTransfECF.sDirDadosECF & sFile As gobjTransfECF.sDirDadosECF & "back\" & sFile
                        
'                        colFile.Add sFile
                        
                    Else
                        
                        lErro1 = Arquivo_Log_Grava(CStr(Now) & " - Aconteceu um erro na carga do arquivo " & sArq, MARCADO)
                        '???? If lErro <> SUCESSO Or lErro1 <> SUCESSO Then gError 133527
                        
                    End If
        
                End If
        
                sFile = Dir
        
            Loop
    
        Loop
    
        If iCarregouAlgo = 0 Then
        
            lErro = Arquivo_Log_Grava(CStr(Now) & " - Nao ha arquivo dos caixas a ser carregado no momento.", MARCADO)
            If lErro <> SUCESSO Then gError 133526
        
        End If

    End If

    Carrega_Arq = SUCESSO

    Exit Function

Erro_Carrega_Arq:

    Carrega_Arq = gErr

    Select Case gErr

        Case 133438, 133526, 133527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163170)

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

Public Function Upload_DadosCC() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim sDir As String
Dim sDiretorio As String
Dim lPos As Long
Dim objLojaConfig As New ClassLojaConfig
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String
Dim colMsg As New Collection

On Error GoTo Erro_Upload_DadosCC

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Or giLocalOperacao = LOCALOPERACAO_BACKOFFICE Or giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then


        objLojaConfig.iFilialEmpresa = EMPRESA_TODA
        objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOSCCBACK
        
        lErro = CF("LojaConfig_Le1", objLojaConfig)
        If lErro <> SUCESSO And lErro <> 126361 Then gError 133626
        
        'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
        If lErro = 126361 Then objLojaConfig.sConteudo = CurDir
        
        sDiretorio = objLojaConfig.sConteudo
       
        'se o diretorio nao for terminado por \  ===> acrescentar
        If right(sDiretorio, 1) <> "\" Then sDiretorio = sDiretorio & "\"
    
        sArquivo = Dir(sDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC)
        
        If sArquivo <> "" Then
        
            Set objFTP = CreateObject("SGEUtil.FTP1")
            
            sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
            If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
            lErro = objFTP.Upload_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, sDiretorio, sArquivo, sErro)
            If lErro <> SUCESSO Then
                Call Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sDiretorio & sArquivo & " NÃO FOI TRANSMITIDO: " & sErro, MARCADO)
                gError 214525
            End If
                                
            lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sDiretorio & sArquivo & " foi transmitido com sucesso.", MARCADO)
            If lErro <> SUCESSO Then gError 133640
    
        End If


    End If

    Upload_DadosCC = SUCESSO

    Exit Function

Erro_Upload_DadosCC:

    Upload_DadosCC = gErr

    Select Case gErr

        Case 214525
            Call Rotina_Erro(vbOKOnly, "ERRO_UPLOAD_ARQUIVO_FTP1", gErr, sDiretorio & sArquivo, sFTPDiretorio & sArquivo, sErro)

        Case 133626, 133640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163171)

    End Select

    Exit Function

End Function

Private Function Gera_Arq_CCB() As Long

Dim lNumIntDocInicial As Long
Dim lNumIntDocFinal As Long
Dim objBarraProgresso As Object
Dim sArqCCBGerado As String
Dim colNomeArq As New Collection
Dim sNomeArq As String
Dim vNomeArq As Variant
Dim lErro As Long

On Error GoTo Erro_Gera_Arq_CCB

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            
        Set objBarraProgresso = BarraProgresso
        
        lErro = CF("Rotina_Gravacao_CC_Back", colNomeArq, lNumIntDocFinal, objBarraProgresso, 0)
        If lErro <> SUCESSO Then gError 133584
        
        For Each vNomeArq In colNomeArq
            lErro = Arquivo_Log_Grava(CStr(Now) & " O arquivo " & vNomeArq & "foi gerado.", MARCADO)
            If lErro <> SUCESSO Then gError 133586
        Next
        
        If colNomeArq.Count = 0 Then
            lErro = Arquivo_Log_Grava(CStr(Now) & " - Nao ha arquivo a ser transmitido para a retaguarda no momento.", MARCADO)
            If lErro <> SUCESSO Then gError 133583
        End If
    
    End If
    
    Gera_Arq_CCB = SUCESSO

    Exit Function

Erro_Gera_Arq_CCB:

    Gera_Arq_CCB = gErr

    Select Case gErr

        Case 133582, 133583, 133584, 133586

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163172)

    End Select

    Exit Function

End Function

Function Upload_DadosCCB() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim colArq As New Collection
Dim vArq As Variant
Dim sDir As String
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String


On Error GoTo Erro_Upload_DadosCCB

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        sArquivo = Dir(gobjTransfECF.sDirDadosCCC & "CC_" & glEmpresa & "_" & "*.ccb")
        
        Do While sArquivo <> ""
        
        
            If sArquivo <> "" Then
            
                Set objFTP = CreateObject("SGEUtil.FTP1")
                
                sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
                If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
            
                lErro = objFTP.Upload_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosCCC, sArquivo, sErro)
                If lErro <> SUCESSO Then gError 214525
                                    
                lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & gobjTransfECF.sDirDadosCCC & sArquivo & " foi transmitido com sucesso.", MARCADO)
                If lErro <> SUCESSO Then gError 133596
        
            End If
        
            sArquivo = Dir
        
        Loop
    
        For Each vArq In colArq
                
           sArquivo = vArq
            
            sDir = Dir(gobjTransfECF.sDirDadosCCC & "back\" & sArquivo)
        
            If Len(sDir) > 0 Then
                Kill gobjTransfECF.sDirDadosCCC & "back\" & sArquivo
            End If
        
            'coloca o arquivo no diretorio DirDadosECF\back
            Name gobjTransfECF.sDirDadosCCC & sArquivo As gobjTransfECF.sDirDadosCCC & "back\" & sArquivo
    
        Next

    End If

    Upload_DadosCCB = SUCESSO

    Exit Function

Erro_Upload_DadosCCB:

    Upload_DadosCCB = gErr

    Select Case gErr

        Case 133582, 133583, 133596

        Case 133588
            Call Rotina_Erro(vbOKOnly, "ERRO_UPLOAD_ARQUIVO_FTP", gErr, gobjTransfECF.sDirDadosCCC & sArquivo, gobjTransfECF.sFTPDiretorio & sArquivo)

        Case 214525
            Call Rotina_Erro(vbOKOnly, "ERRO_UPLOAD_ARQUIVO_FTP1", gErr, gobjTransfECF.sDirDadosCCC & sArquivo, sFTPDiretorio & sArquivo, sErro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163173)

    End Select

    Exit Function

End Function

Public Function Download_DadosCCB() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim objLojaConfig As New ClassLojaConfig
Dim lPos As Long
Dim lPos1 As Long
Dim sMsg1 As String
Dim sFile As String
Dim sDir As String
Dim iResult As VbMsgBoxResult
Dim lErro As Long
Dim iPos1 As Integer
Dim iPos As Integer
Dim sFTPDiretorio As String
Dim objFTP As Object
Dim sErro As String
Dim sRetorno As Long


On Error GoTo Erro_Download_DadosCCB

    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then

        Set objFTP = CreateObject("SGEUtil.FTP1")
        
        sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
        If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
        lErro = objFTP.Download_CC(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosCCC, sErro, ".ccb", sRetorno)
        If lErro <> SUCESSO Then gError 214527
                            
        iPos1 = 1
                            
        If Len(sRetorno) = 0 Then
            iPos1 = 0
        Else
            iPos = InStr(sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
        End If
                            
        Do While iPos1 > 0
        
        
            lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & Mid(sRetorno, iPos1, iPos + 1 - iPos1) & " foi baixado com sucesso.", MARCADO)
            If lErro <> SUCESSO Then gError 133641
            
            iPos1 = iPos + 1
            
            iPos = InStr(sRetorno, ";")
            If iPos = 0 Then
                iPos = Len(sRetorno)
            Else
                iPos = iPos - 1
            End If
            
        Loop

    End If

    Download_DadosCCB = SUCESSO

    Exit Function

Erro_Download_DadosCCB:

    Download_DadosCCB = gErr

    Select Case gErr

        Case 133641
        
        Case 214527
            Call Rotina_Erro(vbOKOnly, "ERRO_DOWNLOAD_CCB_FTP", gErr, sErro)
        

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163174)

    End Select

    Exit Function

End Function

Private Function Carrega_Arq_CCB() As Long

Dim lErro As Long
Dim sArq As String
Dim objBarraProgresso As Object
Dim colFile As New Collection
Dim sDir As String
Dim vFile As Variant
Dim iCarregou As Integer
Dim sFile As String
Dim iCarregouAlgo As Integer

On Error GoTo Erro_Carrega_Arq_CCB

    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then

        Set objBarraProgresso = BarraProgresso
    
        iCarregou = 1
        iCarregouAlgo = 0
    
        Do While iCarregou = 1
    
            iCarregou = 0
    
            sFile = Dir(gobjTransfECF.sDirDadosCCC & "CC_" & glEmpresa & "_" & "*.ccb")
    
            Do While Len(sFile) > 0
            
                sArq = gobjTransfECF.sDirDadosCCC & sFile
        
                'nao para de executar mesmo que tenha dado erro
                lErro = CF("Verifica_Nome_Arquivo", sFile)
                    
                If lErro = SUCESSO Then
                
                    BarraProgresso.Value = 0
                
                    lErro = CF("Rotina_Carga_CC_Back", sArq, objBarraProgresso)
                    If lErro = SUCESSO Then

                        iCarregou = 1
                        iCarregouAlgo = 1
                                   
                        lErro = Arquivo_Log_Grava(CStr(Now) & " - Carregou o arquivo " & sArq, MARCADO)
                        If lErro <> SUCESSO Then gError 133600
                        
                        colFile.Add sFile
                        
                    Else
                        
                        lErro = Arquivo_Log_Grava(CStr(Now) & " - Aconteceu um erro na carga do arquivo " & sArq, MARCADO)
                        If lErro <> SUCESSO Then gError 133601
                        
                    End If
        
                End If
        
                sFile = Dir
        
            Loop
    
        Loop
    
        For Each vFile In colFile
    
            sDir = Dir(gobjTransfECF.sDirDadosCCC & "back\" & vFile)
    
            If Len(sDir) > 0 Then
                Kill gobjTransfECF.sDirDadosCCC & "back\" & vFile
            End If
    
            Name gobjTransfECF.sDirDadosCCC & vFile As gobjTransfECF.sDirDadosCCC & "back\" & vFile
                
        Next
        
        If iCarregouAlgo = 0 Then
        
            lErro = Arquivo_Log_Grava(CStr(Now) & " - Nao ha arquivos dos caixas centrais a serem carregados no momento.", MARCADO)
            If lErro <> SUCESSO Then gError 133603
        
        End If

    End If

    Carrega_Arq_CCB = SUCESSO

    Exit Function

Erro_Carrega_Arq_CCB:

    Carrega_Arq_CCB = gErr

    Select Case gErr

        Case 133600 To 133603

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163175)

    End Select

    Exit Function

End Function

Public Function Upload_DadosBack() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim sDir As String
Dim sDiretorio As String
Dim lPos As Long
Dim objLojaConfig As New ClassLojaConfig
Dim lErro As Long
Dim sFTPDiretorio As String
Dim objFTP As Object
Dim sErro As String

On Error GoTo Erro_Upload_DadosBack

    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then


        objLojaConfig.iFilialEmpresa = EMPRESA_TODA
        objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOSCCBACK
        
        lErro = CF("LojaConfig_Le1", objLojaConfig)
        If lErro <> SUCESSO And lErro <> 126361 Then gError 133627
        
        'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
        If lErro = 126361 Then objLojaConfig.sConteudo = CurDir
        
        sDiretorio = objLojaConfig.sConteudo

        'se o diretorio nao for terminado por \  ===> acrescentar
        If right(sDiretorio, 1) <> "\" Then sDiretorio = sDiretorio & "\"
    
        sArquivo = Dir(sDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK)
        
        If sArquivo <> "" Then
        
            Set objFTP = CreateObject("SGEUtil.FTP1")
            
            sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
            If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
            lErro = objFTP.Upload_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, sDiretorio, sArquivo, sErro)
            If lErro <> SUCESSO Then gError 133611
                                
            lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & sDiretorio & sArquivo & " foi transmitido com sucesso.", MARCADO)
            If lErro <> SUCESSO Then gError 133642
    
        End If

    End If

    Upload_DadosBack = SUCESSO

    Exit Function

Erro_Upload_DadosBack:

    Upload_DadosBack = gErr

    Select Case gErr

        Case 133610
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ARQUIVO_FTP", gErr, gobjTransfECF.sFTPDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK)

        Case 133611
            Call Rotina_Erro(vbOKOnly, "ERRO_UPLOAD_ARQUIVO_FTP1", gErr, sDiretorio & glEmpresa & "_" & giFilialEmpresa & NOME_ARQUIVOBACK, gobjTransfECF.sFTPDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK, sErro)

        Case 133612
            Call Rotina_Erro(vbOKOnly, "ERRO_COMUNICACAO_FTP", gErr)

        Case 133627, 133642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163176)

    End Select

    Exit Function

End Function

Public Function Download_DadosBack() As Long

Dim sArquivo As String
Dim lTeste As Long
Dim objLojaConfig As New ClassLojaConfig
Dim lPos As Long
Dim lPos1 As Long
Dim sMsg1 As String
Dim sFile As String
Dim sDir As String
Dim iResult As VbMsgBoxResult
Dim lErro As Long
Dim objFTP As Object
Dim sFTPDiretorio As String
Dim sErro As String

On Error GoTo Erro_Download_DadosBack

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then


        Set objFTP = CreateObject("SGEUtil.FTP1")
        
        sFTPDiretorio = gobjTransfECF.sFTPURL & "/" & gobjTransfECF.sFTPDiretorio
        If left(UCase(sFTPDiretorio), Len("FTP://")) <> "FTP://" Then sFTPDiretorio = "ftp://" & sFTPDiretorio
        
        sArquivo = Dir(gobjTransfECF.sDirDadosCCC & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK)
    
        lErro = objFTP.Download_Arquivo(sFTPDiretorio, gobjTransfECF.sFTPUserName, gobjTransfECF.sFTPPassword, gobjTransfECF.sDirDadosCCC, sArquivo, sErro)
        If lErro <> SUCESSO Then gError 133619
                            
        lErro = Arquivo_Log_Grava(CStr(Now) & " - O arquivo " & gobjTransfECF.sDirDadosCCC & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK & " foi baixado com sucesso.", MARCADO)
        If lErro <> SUCESSO Then gError 133643

    End If

    Download_DadosBack = SUCESSO

    Exit Function

Erro_Download_DadosBack:

    Download_DadosBack = gErr

    Select Case gErr

        Case 133617
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, gobjTransfECF.sFTPDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK)

        Case 133618
            Call Rotina_Erro(vbOKOnly, "ERRO_COMUNICACAO_FTP", gErr)
        
        Case 133619
            Call Rotina_Erro(vbOKOnly, "ERRO_DOWNLOAD_ARQUIVO_FTP1", gErr, gobjTransfECF.sFTPDiretorio & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK, gobjTransfECF.sDirDadosCCC & glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK, sErro)

        Case 133643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163177)

    End Select

    Exit Function

End Function

