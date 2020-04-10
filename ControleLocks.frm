VERSION 5.00
Begin VB.Form ControleLocks 
   Caption         =   "Controle de Bloqueios"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2775
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   2505
   End
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
      Height          =   495
      Left            =   5925
      TabIndex        =   4
      Top             =   3390
      Width           =   2040
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   255
      TabIndex        =   3
      Top             =   360
      Width           =   7770
   End
   Begin VB.CommandButton BotaoSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5955
      TabIndex        =   2
      Top             =   4095
      Width           =   2040
   End
   Begin VB.CommandButton BotaoExcluirTodos 
      Caption         =   "Excluir Todos os Locks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4170
      Width           =   2040
   End
   Begin VB.CommandButton BotaoExcluirUsu 
      Caption         =   "Excluir Locks do usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3495
      Width           =   2040
   End
End
Attribute VB_Name = "ControleLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AD_Conexao_ResetarLocks Lib "ADSQLMN.DLL" (ByVal lpCon As Long, ByVal lpusername As String) As Long
Private Declare Function AD_Conexao_ObterRegLock Lib "ADSQLMN.DLL" (ByVal lpCon As Long, ByVal l_reg As Long, contador As Integer, prox_reg_vago As Long, tipo As Integer, l_conexao_id As Long, ByVal lpusername As String, ByVal lptabela As String, tam_chave As Integer, ByVal lpchave As String) As Long

Private Sub BotaoAtualizar_Click()

    Call Atualiza

End Sub

Private Sub BotaoExcluirTodos_Click()

    Call AD_Conexao_ResetarLocks(GL_lConexaoDic, "")
    Call BotaoAtualizar_Click

End Sub

Private Sub BotaoExcluirUsu_Click()
Dim sUsu As String

    If List1.ListIndex <> -1 Then
    
        sUsu = Mid(List1.Text, 1, List1.ItemData(List1.ListIndex))
        'MsgBox (sUsu)
        Call AD_Conexao_ResetarLocks(GL_lConexaoDic, sUsu)
        Call BotaoAtualizar_Click
        
    End If
End Sub

Private Sub BotaoSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = Atualiza
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Exit Sub

End Sub

Function Atualiza() As Long

Dim lReg As Long, lErro As Long
Dim iContador As Integer, lProxRegVago As Long, iTipo As Integer, lConexaoId As Long, sUserName As String, sTabela As String, iTamChave As Integer, sChave As String

On Error GoTo Erro_Atualiza

    List1.Clear
    List2.Clear
    
    sUserName = String(255, 0)
    sTabela = String(255, 0)
    sChave = String(255, 0)
    lReg = 1
    
    lErro = AD_Conexao_ObterRegLock(GL_lConexaoDic, lReg, iContador, lProxRegVago, iTipo, lConexaoId, sUserName, sTabela, iTamChave, sChave)
    If lErro <> 0 And lErro <> 100 Then Error 1
    
    Do While lErro = 0
    
        If iContador <> 0 Then
        
            List1.AddItem StringZ(sUserName) & " - " & CStr(lReg) & " - " & StringZ(sTabela) & " - " & IIf(iTipo = 1, "Shared", "Exclusive") & " - " & CStr(iContador) & " - " & CStr(lProxRegVago) & " - " & CStr(lConexaoId) & " - " & CStr(iTamChave) & " - " & sChave
            List1.ItemData(List1.NewIndex) = Len(StringZ(sUserName))
            List2.AddItem CStr(100000 + lConexaoId) & " - " & StringZ(sUserName)
            
        End If
        
        lReg = lReg + 1
        
        lErro = AD_Conexao_ObterRegLock(GL_lConexaoDic, lReg, iContador, lProxRegVago, iTipo, lConexaoId, sUserName, sTabela, iTamChave, sChave)
        If lErro <> 0 And lErro <> 100 Then Error 1
    
    Loop

    Atualiza = SUCESSO
    
    Exit Function
    
Erro_Atualiza:

    Atualiza = Err

    Exit Function

End Function
