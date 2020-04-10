VERSION 5.00
Begin VB.Form AtualizahelpDesk 
   Caption         =   "AtualizaHelpDesk"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "atualizar novo"
      Height          =   735
      Left            =   780
      TabIndex        =   1
      Top             =   1950
      Width           =   1635
   End
   Begin VB.CommandButton BotaAtualizar 
      Caption         =   "Atualizar antigo"
      Height          =   600
      Left            =   915
      TabIndex        =   0
      Top             =   960
      Width           =   1200
   End
End
Attribute VB_Name = "AtualizahelpDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BotaAtualizar_Click()
    Dim ConStrOut As String * 500
    Dim len_out As Integer
    Dim lConexaoMy As Long, sStringConexaoMy As String
    Dim lConexaoControle As Long, sStringConexaoControle As String
    Dim lComandoMy As Long, lErro As Long
    Dim lComando1 As Long, lComando2 As Long, dHora As Double, semail As String
    Dim strackid As String, sname As String, dt As Date, ssubject As String, smessage As String, sstatus As String, scustom1 As String, scustom2 As String
    Dim iStatus As Integer, iTipoAssunto As Integer, iTipoContato As Integer, iCat As Integer, iStatusAnt As Integer
    Dim lTransacao As Long
    
    On Error GoTo Erro_BotaAtualizar_Click
    
    'sStringConexaoMy = "DSN=helpdesk_asp33;UID=root;PWD=SAPWD"
    sStringConexaoMy = "DSN=helpdesk_forw3;UID=forw3;PWD=abc1234"
    
    lConexaoMy = Conexao_Abrir(1, sStringConexaoMy, Len(sStringConexaoMy) + 1, ConStrOut, len_out)

    sStringConexaoControle = "DSN=Controle;UID=sa;PWD=SAPWD"
    
    lConexaoControle = Conexao_Abrir(1, sStringConexaoControle, Len(sStringConexaoControle) + 1, ConStrOut, len_out)

    lTransacao = Transacao_AbrirExt(lConexaoControle)

    lComandoMy = Comando_AbrirExt(lConexaoMy)
    lComando1 = Comando_AbrirExt(lConexaoControle)
    lComando2 = Comando_AbrirExt(lConexaoControle)
    
    strackid = String(255, 0)
    sname = String(255, 0)
    ssubject = String(255, 0)
    smessage = String(65536, 0)
    scustom1 = String(65536, 0)
    scustom2 = String(65536, 0)
    sstatus = String(255, 0)
    semail = String(255, 0)
    
    lErro = Comando_Executar(lComandoMy, "SELECT trackid, name,email,category, dt, subject, LEFT({fn CONVERT(message, CHAR)},255), {fn CONVERT(status,CHAR)}, LEFT({fn CONVERT(custom1, CHAR)},255), LEFT({fn CONVERT(custom2, CHAR)},255) FROM tickets ORDER BY trackid", _
        strackid, sname, semail, iCat, dt, ssubject, smessage, sstatus, scustom1, scustom2)
    If lErro <> AD_SQL_SUCESSO Then Error 20000
    
    lErro = Comando_BuscarProximo(lComandoMy)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'O controle foi compatibilizado com os status do helpdesk
        '0 = Novo, 1 = Aguardando Resposta, 2 = Respondido e 3 = Resolvido
        iStatus = StrParaInt(sstatus)
        
        'Também houve um ajuste no controle para inserir
        '9 = Comercial e 10 - Administrativo
        '1 - Passou de Dúvidas para suporte e os demais permaneceram iguais
        '2 = Erro e 3 = Orçamento
        iTipoAssunto = iCat
        
        iTipoContato = 4 'HelpDesk
    
        lErro = Comando_ExecutarPos(lComando1, "SELECT Status FROM Contatos WHERE Codigo = ?", 0, _
            iStatusAnt, strackid)
        If lErro <> AD_SQL_SUCESSO Then Error 20000

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
        
        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(lComando2, "INSERT INTO Contatos (Codigo,Atendente,Data,Cliente,NomeContato,TipoContato,InforRetorno, " & _
            "Assunto,TipoAssunto,DetAssunto,Status) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
            strackid, scustom2, dt, left(scustom1, 50), sname, iTipoContato, semail, ssubject, iTipoAssunto, left(smessage, 800), iStatus)
            If lErro <> AD_SQL_SUCESSO Then Error 20000
            
        Else
        
            If iStatus <> iStatusAnt Then
            
                lErro = Comando_ExecutarPos(lComando2, "UPDATE Contatos SET Status = ? ", lComando1, iStatus)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
                
            End If
        
        End If
    
        lErro = Comando_BuscarProximo(lComandoMy)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Loop
    
    Call Comando_Fechar(lComandoMy)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Call Transacao_CommitExt(lTransacao)
    
    MsgBox ("Atualizacao efetuada com sucesso")
    
    Exit Sub
    
Erro_BotaAtualizar_Click:

    Call Transacao_RollbackExt(lTransacao)

    MsgBox ("deu erro")
End Sub

Private Sub Command1_Click()

    Dim ConStrOut As String * 500
    Dim len_out As Integer
    Dim lConexaoMy As Long, sStringConexaoMy As String
    Dim lConexaoControle As Long, sStringConexaoControle As String
    Dim lComandoMy As Long, lErro As Long
    Dim lComando1 As Long, lComando2 As Long, dHora As Double, semail As String
    Dim strackid As String, sname As String, dt As Date, ssubject As String, smessage As String, sstatus As String, scustom1 As String, scustom2 As String
    Dim iStatus As Integer, iTipoAssunto As Integer, iTipoContato As Integer, iCat As Integer, iStatusAnt As Integer
    Dim lTransacao As Long
    
    On Error GoTo Erro_BotaAtualizar_Click
    
    'sStringConexaoMy = "DSN=helpdesk_asp33;UID=root;PWD=SAPWD"
    sStringConexaoMy = "DSN=helpdesk_corporator;UID=corporator;PWD=abc123.."
    
    lConexaoMy = Conexao_Abrir(1, sStringConexaoMy, Len(sStringConexaoMy) + 1, ConStrOut, len_out)

    sStringConexaoControle = "DSN=Controle;UID=sa;PWD=SAPWD"
    
    lConexaoControle = Conexao_Abrir(1, sStringConexaoControle, Len(sStringConexaoControle) + 1, ConStrOut, len_out)

    lTransacao = Transacao_AbrirExt(lConexaoControle)

    lComandoMy = Comando_AbrirExt(lConexaoMy)
    lComando1 = Comando_AbrirExt(lConexaoControle)
    lComando2 = Comando_AbrirExt(lConexaoControle)
    
    strackid = String(255, 0)
    sname = String(255, 0)
    ssubject = String(255, 0)
    smessage = String(65536, 0)
    scustom1 = String(65536, 0)
    scustom2 = String(65536, 0)
    sstatus = String(255, 0)
    semail = String(255, 0)
    
    lErro = Comando_Executar(lComandoMy, "SELECT trackid, name,email,category, dt, subject, LEFT({fn CONVERT(message, CHAR)},255), {fn CONVERT(status,CHAR)}, LEFT({fn CONVERT(custom1, CHAR)},255), LEFT({fn CONVERT(custom2, CHAR)},255) FROM tickets ORDER BY trackid", _
        strackid, sname, semail, iCat, dt, ssubject, smessage, sstatus, scustom1, scustom2)
    If lErro <> AD_SQL_SUCESSO Then Error 20000
    
    lErro = Comando_BuscarProximo(lComandoMy)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'O controle foi compatibilizado com os status do helpdesk
        '0 = Novo, 1 = Aguardando Resposta, 2 = Respondido e 3 = Resolvido
        iStatus = StrParaInt(sstatus)
        
        'Também houve um ajuste no controle para inserir
        '9 = Comercial e 10 - Administrativo
        '1 - Passou de Dúvidas para suporte e os demais permaneceram iguais
        '2 = Erro e 3 = Orçamento
        iTipoAssunto = iCat
        
        iTipoContato = 4 'HelpDesk
    
        lErro = Comando_ExecutarPos(lComando1, "SELECT Status FROM Contatos WHERE Codigo = ?", 0, _
            iStatusAnt, strackid)
        If lErro <> AD_SQL_SUCESSO Then Error 20000

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
        
        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(lComando2, "INSERT INTO Contatos (Codigo,Atendente,Data,Cliente,NomeContato,TipoContato,InforRetorno, " & _
            "Assunto,TipoAssunto,DetAssunto,Status) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
            strackid, scustom2, dt, left(scustom1, 50), sname, iTipoContato, semail, ssubject, iTipoAssunto, left(smessage, 800), iStatus)
            If lErro <> AD_SQL_SUCESSO Then Error 20000
            
        Else
        
            If iStatus <> iStatusAnt Then
            
                lErro = Comando_ExecutarPos(lComando2, "UPDATE Contatos SET Status = ? ", lComando1, iStatus)
                If lErro <> AD_SQL_SUCESSO Then Error 20000
                
            End If
        
        End If
    
        lErro = Comando_BuscarProximo(lComandoMy)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20000
    
    Loop
    
    Call Comando_Fechar(lComandoMy)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Call Transacao_CommitExt(lTransacao)
    
    MsgBox ("Atualizacao efetuada com sucesso")
    
    Exit Sub
    
Erro_BotaAtualizar_Click:

    Call Transacao_RollbackExt(lTransacao)

    MsgBox ("deu erro")
End Sub


