VERSION 5.00
Begin VB.Form UsuarioEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corporator"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   ControlBox      =   0   'False
   Icon            =   "UsuarioEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BotaoCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2340
      Picture         =   "UsuarioEmpresa.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3390
      Width           =   975
   End
   Begin VB.Frame FrameEmpresa 
      Caption         =   "Empresa"
      Height          =   1515
      Left            =   105
      TabIndex        =   7
      Top             =   1695
      Width           =   3990
      Begin VB.ComboBox ComboEmpresa 
         Height          =   315
         Left            =   810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   3000
      End
      Begin VB.ComboBox ComboFilial 
         Height          =   315
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   3000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   435
         Width           =   555
      End
      Begin VB.Label FilialLabel 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   1005
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1515
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   3990
      Begin VB.TextBox TextSenha 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   810
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   945
         Width           =   2985
      End
      Begin VB.ComboBox ComboUsuario 
         Height          =   315
         Left            =   810
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   390
         Width           =   3000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   990
         Width           =   615
      End
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   915
      Picture         =   "UsuarioEmpresa.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3390
      Width           =   975
   End
End
Attribute VB_Name = "UsuarioEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjUsuarioEmpresa As New ClassUsuarioEmpresa

Function Verifica_Senha(sCodUsuario As String, sSenha As String) As Long
'Verifica a senha do usuario na tabela

Dim lErro As Long, objUsuario As New ClassDicUsuario

On Error GoTo Erro_Verifica_Senha

    'Preenche a chave de objUsuarios
    objUsuario.sCodUsuario = sCodUsuario

    lErro = DicRotinas.DicUsuario_Le(objUsuario)
    If lErro <> SUCESSO Then Error 50164
    
    If objUsuario.iAtivo <> USUARIO_ATIVO Then Error 41660
    
    'Senha nao esta cadastrada
    If objUsuario.sSenha <> sSenha Then Error 50165
    
    'Verifica se a Data da Senha não está expirada.
    If objUsuario.dtDataValidade <> DATA_NULA And objUsuario.dtDataValidade < Date Then Error 50173
        
    Verifica_Senha = SUCESSO
    
    Exit Function
    
Erro_Verifica_Senha:

    Verifica_Senha = Err

    Select Case Err
    
        Case 50164, 41660
        
        Case 50165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SENHA_INVALIDA", Err)
            
        Case 50173 'Data invalida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SENHA_EXPIRADA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175652)
        
    End Select
    
    Exit Function

End Function

Private Sub Limpar_Tela_UsuarioEmpresa()

    'Limpar os Campos
    TextSenha.Text = ""
    ComboEmpresa.Clear
    ComboFilial.Clear
    
End Sub

Function Trata_Parametros(objUsuarioEmpresa1 As ClassUsuarioEmpresa) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjUsuarioEmpresa = objUsuarioEmpresa1
    gobjUsuarioEmpresa.iTelaOK = False
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175653)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoOk_Click
    
    'Verificar se o campo Usuario esta preenchido
    If Len(ComboUsuario) = 0 Then Error 50158
    
    'Verificar se a senha esta preenchida
    If Len(TextSenha) = 0 Then Error 50159
    
    'Verificar se a Empresa esta preenchida
    If Len(ComboEmpresa) = 0 Then Error 50160
    
    'Verificar se a Filial esta preenchida
    If Len(ComboFilial) = 0 Then Error 50161
    
    'Tranferir os dados da tela para gobjUsuarioEmpresa
    gobjUsuarioEmpresa.sNomeEmpresa = ComboEmpresa.Text
    gobjUsuarioEmpresa.sNomeFilial = ComboFilial.Text
    gobjUsuarioEmpresa.sSenha = TextSenha.Text
    gobjUsuarioEmpresa.lCodEmpresa = ComboEmpresa.ItemData(ComboEmpresa.ListIndex)
    gobjUsuarioEmpresa.iCodFilial = ComboFilial.ItemData(ComboFilial.ListIndex)
    gobjUsuarioEmpresa.sCodUsuario = ComboUsuario.Text
    gobjUsuarioEmpresa.iTelaOK = True
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOk_Click:

    Select Case Err
    
        Case 50158 'Usuario nao preenchido
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)
            
        Case 50159 'Senha nao preenchida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", Err)
        
        Case 50160 'Empresa nao preenchida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", Err)
        
        Case 50161 'Filial nao preenchida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175654)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ComboEmpresa_Click()

Dim lErro As Long
Dim colFilialEmpresa As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim lCodEmpresa As Long

On Error GoTo Erro_ComboEmpresa_Click

    If ComboEmpresa.ListIndex = -1 Then Exit Sub

    'Limpar a ComboFilial
    ComboFilial.Clear
    
    'Ler o Codigo da Empresa
    lCodEmpresa = ComboEmpresa.ItemData(ComboEmpresa.ListIndex)
    
    'Carregar todas as filiais da empresa selecionada para os quais o usuário está autorizado a acessar
    lErro = FiliaisEmpresa_Le_Usuario(ComboUsuario.Text, lCodEmpresa, colFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 50172 Then Error 50176

    'Se não houverem filiais para empresa/usuário em questão ==> erro
    If lErro = 50172 Then Error 50177

    If giTipoVersao = VERSAO_FULL Then
        
        For Each objUsuarioEmpresa In colFilialEmpresa
            
            If objUsuarioEmpresa.iCodFilial = EMPRESA_TODA Then
                objUsuarioEmpresa.sNomeFilial = EMPRESA_TODA_NOME
            End If
            
            ComboFilial.AddItem objUsuarioEmpresa.sNomeFilial
            ComboFilial.ItemData(ComboFilial.NewIndex) = objUsuarioEmpresa.iCodFilial
        
        Next
    
        If ComboFilial.ListCount >= 2 Then
            ComboFilial.ListIndex = 1
        Else
            If ComboFilial.ListCount >= 1 Then ComboFilial.ListIndex = 0
        End If
    
    ElseIf giTipoVersao = VERSAO_LIGHT Then
    
        For Each objUsuarioEmpresa In colFilialEmpresa
        
            If objUsuarioEmpresa.iCodFilial <> EMPRESA_TODA Then
                ComboFilial.AddItem objUsuarioEmpresa.sNomeFilial
                ComboFilial.ItemData(ComboFilial.NewIndex) = objUsuarioEmpresa.iCodFilial
            End If
        
        Next

        If ComboFilial.ListCount >= 1 Then ComboFilial.ListIndex = 0
    
    End If
    
    Exit Sub
        
Erro_ComboEmpresa_Click:

    Select Case Err
    
        Case 50176
        
        Case 50177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_SEM_FILIAIS", Err, ComboEmpresa.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175655)
    
    End Select
    
    Exit Sub
        
End Sub

Private Sub ComboUsuario_Click()

    Call Limpar_Tela_UsuarioEmpresa
    
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As ClassDicUsuario
Dim lAltura As Long

On Error GoTo Erro_Form_Load

    Me.WindowState = 0 'normal

'    If giTipoVersao = VERSAO_LIGHT Then
'
'        ComboFilial.left = POSICAO_FORA_TELA
'        ComboFilial.TabStop = False
'        lAltura = FrameEmpresa.Height - ComboFilial.top
'        FrameEmpresa.Height = ComboFilial.top
'        BotaoOk.top = BotaoOk.top - lAltura
'        BotaoCancela.top = BotaoCancela.top - lAltura
'        UsuarioEmpresa.Height = UsuarioEmpresa.Height - lAltura
'
'    End If

    If giTipoVersao = VERSAO_LIGHT Then
    
        FrameEmpresa.Visible = False
        lAltura = FrameEmpresa.Height
        BotaoOk.top = BotaoOk.top - lAltura
        BotaoCancela.top = BotaoCancela.top - lAltura
        UsuarioEmpresa.Height = UsuarioEmpresa.Height - lAltura
        
    End If

    'Le todos os usuarios da tabela usuarios e coloca na colecao
    lErro = Usuarios_Le_Todos1(colUsuarios)
    If lErro <> SUCESSO Then Error 50153

    'Coloca todos os Usuarios na ComboUsuario
    For Each objUsuarios In colUsuarios
        ComboUsuario.AddItem objUsuarios.sCodUsuario
    Next

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
    
        Case 50153
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175656)
    
    End Select
    
    Exit Sub

End Sub

Private Sub TextSenha_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colEmpresas As New Collection
Dim objEmpresa As ClassDicEmpresa

On Error GoTo Erro_TextSenha_LostFocus
    
    'Se a senha e a comboUsuario estiverem preenchidos
    If Len(ComboUsuario) <> 0 And Len(TextSenha) <> 0 Then
        
        'Se a senha esta cadastrada e se é valida
        lErro = Verifica_Senha(ComboUsuario.Text, TextSenha.Text)
        If lErro <> SUCESSO Then Error 50155
    
        'Limpar a ComboEmpresa
        ComboEmpresa.Clear
        
        'Carregar as Empresas que o usuario esta autorizado a acessar
        lErro = Empresas_Le_Usuario(ComboUsuario.Text, colEmpresas)
        If lErro <> SUCESSO And lErro <> 50183 Then Error 50168
     
        'não há empresa cadastrada para o usuário
        If lErro = 50183 Then Error 50175
     
        For Each objEmpresa In colEmpresas
            
            ComboEmpresa.AddItem objEmpresa.sNome
            ComboEmpresa.ItemData(ComboEmpresa.NewIndex) = objEmpresa.lCodigo
            
        Next
            
        If ComboEmpresa.ListCount >= 1 Then ComboEmpresa.ListIndex = 0
                    
    End If
        
    Exit Sub
        
Erro_TextSenha_LostFocus:

    Cancel = True
    
    Select Case Err
    
        Case 50168
        
        Case 50155 'Senha nao cadastrada
        
        Case 50175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_SEM_EMPRESA", Err, ComboUsuario.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175657)
    
    End Select
    
    Exit Sub
        
End Sub
