VERSION 5.00
Begin VB.Form AtualizaLivezila 
   Caption         =   "AtualizaLivezilla"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoAtualizar 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   885
      TabIndex        =   0
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label Total 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2220
      TabIndex        =   4
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   1455
      Width           =   915
   End
   Begin VB.Label De 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   975
      TabIndex        =   2
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Registro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   1455
      Width           =   915
   End
End
Attribute VB_Name = "AtualizaLivezila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BotaoAtualizar_Click()

Dim ConStrOut As String * 500
Dim len_out As Integer, lTransacao As Long, sCliAux As String
Dim lConexaoMy As Long, sStringConexaoMy As String, lComandoAux As Long
Dim lConexaoControle As Long, sStringConexaoControle As String
Dim alComandoMy(1 To 4) As Long, lErro As Long, lPos As Long, lSeq As Long
Dim alComando(1 To 7) As Long, iIndice As Integer, iAux As Integer, lCliAux As Long
Dim lTime As Long, lEndTime As Long, sChat_id As String, sFullName As String
Dim sGroup_id, sHtml As String, sEmail As String, sCompany As String, sHost As String, sIP As String, sQuestion As String
Dim dtData As Date, sHora As String, dDuracao As Double, lTanHtml As Long, lParte As Long, lNumPartes As Long
Dim sHtmlAux As String, lDias As Long, lHoras As Long, lMinutos As Long, lSegundos As Long, sIDCli As String
Dim lCliente As Long, lProximo As Long, lContador As Long, sContadorTotal As String, bTransAberta As Boolean
Dim sCodChat As String, lSeqMax As Long, sChatU As String, sChatPU As String, lQtdeChats As Long, sTanHtml As String
    
Const PASTA_HTML = "\\ASP22\c$\Inetpub\Controle\ArqChat\"
    
On Error GoTo Erro_BotaAtualizar_Click

    bTransAberta = False

    Total.Caption = "0"
    De.Caption = "0"
    
    sStringConexaoControle = "DSN=Controle;UID=sa;PWD=SAPWD"
    lConexaoControle = Conexao_Abrir(1, sStringConexaoControle, Len(sStringConexaoControle) + 1, ConStrOut, len_out)
    
    sStringConexaoMy = "DSN=helpdesk_corporator;UID=corporator;PWD=abc123.."
    lConexaoMy = Conexao_Abrir(1, sStringConexaoMy, Len(sStringConexaoMy) + 1, ConStrOut, len_out)
  
    lContador = 0
    
    'Abre os comandos
    For iIndice = LBound(alComandoMy) To UBound(alComandoMy)
        alComandoMy(iIndice) = Comando_AbrirExt(lConexaoMy)
        If alComandoMy(iIndice) = 0 Then Error 20000
    Next
    
    lComandoAux = Comando_AbrirExt(lConexaoControle)
    If lComandoAux = 0 Then Error 20000
    
    'Lê o mais recente
    lErro = Comando_Executar(lComandoAux, "SELECT MAX(Seq) FROM Chat", lSeq)
    If lErro <> AD_SQL_SUCESSO Then Error 20000

    lErro = Comando_BuscarProximo(lComandoAux)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    If lSeq > (7 * 86400) Then '7 * 24 *60 *60
        lSeq = lSeq - (7 * 86400) 'Relê até os da semana anterior (não dá para ter certeza absoluta quando um ticket aparece nessa tabela)
    Else
        lSeq = 0
    End If
    
    Call Comando_Fechar(lComandoAux)
    
    'Vê quantos registros são
    sContadorTotal = String(255, 0)
    lErro = Comando_Executar(alComandoMy(4), "SELECT {fn CAST(COUNT(*) AS CHAR)} FROM livezillachat_archive WHERE time > ?", sContadorTotal, lSeq)
    If lErro <> AD_SQL_SUCESSO Then Error 20000

    lErro = Comando_BuscarPrimeiro(alComandoMy(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000

    Total.Caption = Format(CLng(sContadorTotal), "#,##0")
    
    sChat_id = String(STRING_MAXIMO, 0)
    sFullName = String(STRING_MAXIMO, 0)
    sGroup_id = String(STRING_MAXIMO, 0)
    sEmail = String(STRING_MAXIMO, 0)
    sCompany = String(STRING_MAXIMO, 0)
    sHost = String(STRING_MAXIMO, 0)
    sIP = String(STRING_MAXIMO, 0)
    sQuestion = String(STRING_MAXIMO, 0)
    
    'Lê todas conversas iniciadas após 24 da última registrada
    sTanHtml = String(255, 0)
    lErro = Comando_Executar(alComandoMy(1), "SELECT time, endtime, chat_id, fullname, group_id, {fn CAST(LENGTH(html) AS CHAR) }, email, company, host, ip, question FROM livezillachat_archive WHERE time > ? ORDER BY time ", _
                                        lTime, lEndTime, sChat_id, sFullName, sGroup_id, sTanHtml, sEmail, sCompany, sHost, sIP, sQuestion, lSeq)
    If lErro <> AD_SQL_SUCESSO Then Error 20000
    
    lErro = Comando_BuscarPrimeiro(alComandoMy(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lContador = 0 Then
            lTransacao = Transacao_AbrirExt(lConexaoControle)
            bTransAberta = True
            
            'Abre os comandos
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_AbrirExt(lConexaoControle)
                If alComando(iIndice) = 0 Then Error 20000
            Next
        End If
        
        lTanHtml = CLng(sTanHtml)
        lContador = lContador + 1
       
        'Verifica se ela já foi cadastrada
        lErro = Comando_Executar(alComando(1), "SELECT CliConferido FROM Chat WHERE Codigo = ?", _
            iAux, sChat_id)
        If lErro <> AD_SQL_SUCESSO Then Error 20000

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
        
        If lErro = AD_SQL_SEM_DADOS Then
    
            'O tempo está em segundos com data base de 01/01/1970
            lSegundos = lTime - (3 * 60 * 60) 'Ajuste de 3hs por conta do fuso horário
            lMinutos = CLng((lSegundos / 60) - 0.5)
            lSegundos = lSegundos - (lMinutos * 60)
            lHoras = CLng((lMinutos / 60) - 0.5)
            lMinutos = lMinutos - (lHoras * 60)
            lDias = CLng((lHoras / 24) - 0.5)
            lHoras = lHoras - (lDias * 24)
            
            dtData = DateAdd("d", lDias, #1/1/1970#)
            dDuracao = (lEndTime - lTime) / 60
            sHora = Format(lHoras, "00") & ":" & Format(lMinutos, "00") & ":" & Format(lSegundos, "00")
            
            sHtml = "<HTML><HEAD><STYLE TYPE=""text/css"">BODY{margin:4px;font-family:verdana,arial;font-size:12px;}"
            sHtml = sHtml & vbNewLine & "TABLE{border:0px;text-align:center;width:100%;}"
            sHtml = sHtml & vbNewLine & "TD{margin:0px;padding:0px;text-align:left;font-family:verdana,arial;font-size:12px;}"
            sHtml = sHtml & vbNewLine & "TD A{font-size:12px;color:#787878;text-decoration:underline;}"
            sHtml = sHtml & vbNewLine & "TD A:VISITED{font-size:12px;color:#787878;text-decoration:underline;}"
            sHtml = sHtml & vbNewLine & "TD A:ACTIVE{font-size:12px;color:#787878;text-decoration:underline;}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".lz_message_translation{font-size:11px;border-left:0px solid #99ccff;margin-top:5px;padding-left:5px;font-style:italic;color:#787878;}"
            sHtml = sHtml & vbNewLine & ".TCM{height:20px;width:100%;margin:0px;margin-top:6px;}"
            sHtml = sHtml & vbNewLine & ".TCF{width:100%;margin:0px;margin-bottom:2px;}"
            sHtml = sHtml & vbNewLine & ".TCB{margin:0px;width:auto;margin-top:2px;}"
            sHtml = sHtml & vbNewLine & ".TCB TD{font-size:11px;}"
            sHtml = sHtml & vbNewLine & ".TCB A{font-size:11px;}"
            sHtml = sHtml & vbNewLine & ".TCB A:ACTIVE{font-size:11px;}"
            sHtml = sHtml & vbNewLine & ".TCB A:VISITED{font-size:11px;}"
            sHtml = sHtml & vbNewLine & ".TCBB{margin:0px;width:100%;padding:4px;padding-left:0px;}"
            sHtml = sHtml & vbNewLine & ".TCQM{margin:0px;padding:5px;}"
            sHtml = sHtml & vbNewLine & ".TCQH{margin:0px;padding-left:5px;}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FCM0{margin:0px;padding:0px;height:20px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_gray_left.gif');}"
            sHtml = sHtml & vbNewLine & ".FCM1{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_gray_center.gif');}"
            sHtml = sHtml & vbNewLine & ".FCM2{margin:0px;padding:0px 0px 2px 4px;text-align:left;vertical-align: middle;background-image: url('C:\Program Files\LiveZilla/images/chat_bg_gray_center.gif');font-size:10px;font-weight:bold;font-family:verdana,arial;color:#696969;}"
            sHtml = sHtml & vbNewLine & ".FCM3{margin:0px;padding:0px 4px 1px 0px;text-align:right;vertical-align:middle;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_gray_center.gif');font-size:10px;font-weight:bold;font-family:verdana,arial;color:#8c8c8c;}"
            sHtml = sHtml & vbNewLine & ".FCM4{margin:0px;padding:0px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_gray_right.gif');}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FCMg0{margin:0px;padding:0px;height:20px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_green_left.gif');}"
            sHtml = sHtml & vbNewLine & ".FCMg1{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_green_center.gif');}"
            sHtml = sHtml & vbNewLine & ".FCMg2{margin:0px;padding: 0px 0px 2px 4px;text-align:left;vertical-align: middle;background-image: url('C:\Program Files\LiveZilla/images/chat_bg_green_center.gif');font-size:10px;font-weight:bold;font-family:verdana,arial;color:white}"
            sHtml = sHtml & vbNewLine & ".FCMg3{margin:0px;padding: 0px 4px 1px 0px;text-align:right;vertical-align:middle;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_green_center.gif');font-size:10px;font-weight:bold;font-family:verdana,arial;color:white;}"
            sHtml = sHtml & vbNewLine & ".FCMg4{margin:0px;padding:0px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_green_right.gif');}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FCQF0{margin:0px;padding:0px 0px 8px 10px;text-align:left;}"
            sHtml = sHtml & vbNewLine & ".FCQF1{margin:0px;padding:0px;width:6px;height:6px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_top_left.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF2{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_middle_top.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF3{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_top_right.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF4{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_middle_left.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF5{margin:0px;padding:0px;vertical-align:top;background-image:url('C:\Program Files\LiveZilla/images/quote_bg.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF6{margin:0px;padding:0px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_middle_right.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF7{margin:0px;padding:0px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_bottom_left.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF8{margin:0px;padding:0px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_middle_bottom.gif');}"
            sHtml = sHtml & vbNewLine & ".FCQF9{margin:0px;padding:0px;height:6px;width:6px;background-image:url('C:\Program Files\LiveZilla/images/quote_bg_bottom_right.gif');}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FCCF1{margin:0px;padding:0px;width:3px;height:3px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_continue_left_top.gif');}"
            sHtml = sHtml & vbNewLine & ".FCCF2{margin:0px;padding:0px;width:3px;height:3px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_continue_right_top.gif');}"
            sHtml = sHtml & vbNewLine & ".FCCF3{margin:0px;padding:0px;width:3px;height:3px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_continue_left_bottom.gif');}"
            sHtml = sHtml & vbNewLine & ".FCCF4{margin:0px;padding:0px;width:3px;height:3px;background-image:url('C:\Program Files\LiveZilla/images/chat_bg_continue_right_bottom.gif');}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FCCF {margin:0px;padding: 0px 0px 0px 10px;text-align:left;background:#fbfbfb;}"
            sHtml = sHtml & vbNewLine & ".FCM5 {margin:0px;padding: 2px 0px 0px 10px;text-align: left;}"
            sHtml = sHtml & vbNewLine & ".FCMCA{margin:0px;padding: 0px 0px 0px 10px;text-align: left;}"
            sHtml = sHtml & vbNewLine & ".FCMCB{margin:0px;padding: 0px 0px 0px 5px;text-align: left;}"
            sHtml = sHtml & vbNewLine & ".TCCL{width:25px;border-top: 2px solid white;background-image: url('C:\Program Files\LiveZilla/images/chat_chat.gif');background-repeat: no-repeat; background-position: bottom;}"
            sHtml = sHtml & vbNewLine & ".TCBB{margin:5px;}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".FILEREQUESTLINK_FILE{color:#8A8A8A;text-decoration:none;cursor:pointer;font-weight:bold;padding-left:20px;height:16px;background-image:url('C:\Program Files\LiveZilla/images/chat_file.gif');background-repeat:no-repeat;}"
            sHtml = sHtml & vbNewLine & ".FILEREQUESTLINK_ALLOW{color:#8A8A8A;text-decoration:none;cursor:pointer;font-weight:bold;padding-left:19px;height:16px;background-image:url('C:\Program Files\LiveZilla/images/chat_accept.gif');background-repeat:no-repeat;}"
            sHtml = sHtml & vbNewLine & ".FILEREQUESTLINK_REJECT{color:#8A8A8A;text-decoration:none;cursor:pointer;font-weight:bold;padding-left:19px;height:16px;background-image:url('C:\Program Files\LiveZilla/images/chat_decline.gif');background-repeat:no-repeat;}"
            sHtml = sHtml & vbNewLine & ""
            sHtml = sHtml & vbNewLine & ".SMILEYCOOL{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/cool.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYCRY{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/cry.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYLOL{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/lol.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYNEUTRAL{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/neutral.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYQUESTION{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/question.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYSAD{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/sad.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYSHOCKED{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/shocked.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYSICK{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/sick.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYSLEEP{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/sleep.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYSMILE{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/smile.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYWINK{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/wink.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & ".SMILEYTONGUE{width:22px;height:22px;background-image:url('C:\Program Files\LiveZilla/images/smilies/tongue.gif');background-repeat: no-repeat;}"
            sHtml = sHtml & vbNewLine & "</STYLE></HEAD><BODY>"
    
            lNumPartes = CInt(lTanHtml / 250)
            If lNumPartes * 250 < lTanHtml Then lNumPartes = lNumPartes + 1
            
            For lParte = 1 To lNumPartes
            
                lProximo = ((lParte - 1) * 250) + 1
            
                sHtmlAux = String(STRING_MAXIMO, 0)
            
                'Vai lento as partes da mensagem
                lErro = Comando_Executar(alComandoMy(2), "SELECT MID(html,?,250) FROM livezillachat_archive WHERE chat_id = ?", sHtmlAux, lProximo, sChat_id)
                If lErro <> AD_SQL_SUCESSO Then
                    Error 20000
                End If
                
                lErro = Comando_BuscarPrimeiro(alComandoMy(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO_PARCIAL Then
                    Error 20000
                End If
                
                sHtml = sHtml & sHtmlAux
    
            Next
            
            sHtml = sHtml & "</BODY></HTML>"
            
            'Ajusta a localização das imagens
            sHtml = Replace(sHtml, "C:\Program Files\LiveZilla/images/", "./img/")
            
            'Grava o html
            Open PASTA_HTML & sChat_id & ".html" For Output As #1
            Print #1, sHtml
            Close #1
            
            'Tenta achar o cliente
            
            lCliente = 0
            
            '1 Parte da mensagem com o ID
            lPos = InStr(1, sHtml, "C_ID=")
            If lPos <> 0 Then
            
                sIDCli = Mid(lPos + 5, 10)
                lPos = InStr(1, sIDCli, " ")
                If lPos <> 0 Then sIDCli = left(sIDCli, lPos)
                sIDCli = Trim(sIDCli)
                
                If IsNumeric(sIDCli) Then
                
                    lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Clientes WHERE Codigo = ?", lCliAux, StrParaLong(sIDCli))
                    If lErro <> AD_SQL_SUCESSO Then Error 20000
            
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                    
                    If lErro = AD_SQL_SUCESSO Then
                        lCliente = lCliAux
                    End If
                    
                End If
            
            End If
            
            '2 Nome exato
            
            If lCliente = 0 Then
            
                lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM Clientes WHERE NomeReduzido = ? OR RazaoSocial = ?", lCliAux, sCompany, sCompany)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
        
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                
                If lErro = AD_SQL_SUCESSO Then
                    lCliente = lCliAux
                End If
            
            End If
            
            '3 IP
            
            If lCliente = 0 Then
            
                lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM Clientes WHERE IPServidorExt = ?", lCliAux, sIP)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
        
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                
                If lErro = AD_SQL_SUCESSO Then
                    lCliente = lCliAux
                End If
            
            End If
            
            '4 Parte do nome
            
            If lCliente = 0 Then
            
                sCliAux = sCompany
                lPos = InStr(1, sCliAux, " ")
                If lPos <> 0 Then sCliAux = Trim(left(sCliAux, lPos))
            
                lErro = Comando_Executar(alComando(6), "SELECT Codigo FROM Clientes WHERE NomeReduzido LIKE ? OR RazaoSocial LIKE ?", lCliAux, sCliAux, sCliAux)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
        
                lErro = Comando_BuscarPrimeiro(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                
                If lErro = AD_SQL_SUCESSO Then
                    lCliente = lCliAux
                End If
            
            End If
    
            '5 pelo email
            
            If lCliente = 0 Then
            
                sCliAux = sEmail
                lPos = InStr(1, sCliAux, "@")
                If lPos <> 0 Then
                    sCliAux = Trim(Mid(sCliAux, lPos + 1))
                    lPos = InStr(1, sCliAux, ".")
                    If lPos <> 0 Then sCliAux = left(sCliAux, lPos - 1)
                End If
            
                lErro = Comando_Executar(alComando(7), "SELECT Codigo FROM Clientes WHERE NomeReduzido LIKE ? OR RazaoSocial LIKE ?", lCliAux, sCliAux, sCliAux)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
        
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                
                If lErro = AD_SQL_SUCESSO Then
                    lCliente = lCliAux
                End If
            
            End If
        
            lErro = Comando_Executar(alComando(2), "INSERT INTO Chat (Codigo,Seq,CliLivezilla,CliConferido,Cliente,Assunto,Data,Hora,Duracao,Contato,Email,Host,IP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
            sChat_id, lTime, sCompany, 0, lCliente, sQuestion, dtData, sHora, dDuracao, sFullName, sEmail, sHost, sIP)
            If lErro <> AD_SQL_SUCESSO Then Error 20000
                               
        End If
        
        De.Caption = Format(StrParaLong(De.Caption) + 1, "#,##0")
        
        DoEvents
    
        lErro = Comando_BuscarProximo(alComandoMy(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
                
        'Se já tem mais de 100 faz o commit
        If lContador >= 100 And lErro <> AD_SQL_SEM_DADOS Then
    
            'Fecha os comandos
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
                    
            Call Transacao_CommitExt(lTransacao)
            bTransAberta = False
            
            lContador = 0
            
        End If
        
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComandoMy) To UBound(alComandoMy)
        Call Comando_Fechar(alComandoMy(iIndice))
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
            
    If bTransAberta Then Call Transacao_CommitExt(lTransacao)
    bTransAberta = False
    
    lTransacao = Transacao_AbrirExt(lConexaoControle)
    bTransAberta = True
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexaoControle)
        If alComando(iIndice) = 0 Then Error 20000
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Clientes ORDER BY Codigo ", 0, lCliAux)
    If lErro <> AD_SQL_SUCESSO Then Error 20000

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sCodChat = String(STRING_MAXIMO, 0)
        sChatU = ""
        sChatPU = ""
    
        lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Chat WHERE Cliente = ? ORDER BY Seq DESC ", sCodChat, lCliAux)
        If lErro <> AD_SQL_SUCESSO Then Error 20000
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
        
        If lErro = AD_SQL_SUCESSO Then
        
            sChatU = sCodChat
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
            
            If lErro = AD_SQL_SUCESSO Then
                sChatPU = sCodChat
            End If
        
        End If
    
        lErro = Comando_Executar(alComando(3), "SELECT COUNT(*) FROM Chat WHERE Cliente = ? ", lQtdeChats, lCliAux)
        If lErro <> AD_SQL_SUCESSO Then Error 20000
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
        
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE Clientes SET UltChat = ?, PenultChat = ?, QtdeChats = ? ", alComando(1), sChatU, sChatPU, lQtdeChats)
        If lErro <> AD_SQL_SUCESSO Then Error 20000
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
            
    If bTransAberta Then Call Transacao_CommitExt(lTransacao)
    
    MsgBox ("Atualizacao efetuada com sucesso")
    
    Exit Sub
    
Erro_BotaAtualizar_Click:

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComandoMy) To UBound(alComandoMy)
        Call Comando_Fechar(alComandoMy(iIndice))
    Next

    Call Transacao_RollbackExt(lTransacao)

    Call MsgBox("Deu erro. VB: " & CStr(Err) & "-" & Err.Description)
    
    'Resume Next
End Sub


