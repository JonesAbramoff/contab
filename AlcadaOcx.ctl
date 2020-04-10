VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.UserControl AlcadaOcx 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   8475
   Begin VB.ListBox UsuariosComAlcada 
      Height          =   2400
      Left            =   5970
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1890
      Left            =   300
      TabIndex        =   14
      Top             =   525
      Width           =   5325
      Begin VB.ComboBox Usuario 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   345
         Width           =   3180
      End
      Begin VB.Label UsuariosLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   390
         Width           =   1410
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label Nome 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   945
         TabIndex        =   2
         Top             =   1365
         Width           =   4065
      End
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   975
         TabIndex        =   1
         Top             =   870
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   255
         TabIndex        =   15
         Top             =   900
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Limites"
      Height          =   1275
      Left            =   270
      TabIndex        =   11
      Top             =   2670
      Width           =   5325
      Begin MSMask.MaskEdBox LimiteOperacao 
         Height          =   300
         Left            =   2655
         TabIndex        =   3
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LimiteMensal 
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   750
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Limite por Operação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   750
         TabIndex        =   13
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Limite Mensal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1305
         TabIndex        =   12
         Top             =   780
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AlcadaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AlcadaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AlcadaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AlcadaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Usuários com alçada"
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
      Left            =   5985
      TabIndex        =   18
      Top             =   1305
      Width           =   1785
   End
End
Attribute VB_Name = "AlcadaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoUsuario As AdmEvento
Attribute objEventoUsuario.VB_VarHelpID = -1

Public iAlterado As Integer

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuario com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuario As New ClassUsuario
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Usuarios

    'Le nomes reduzidos de todos os usuarios e coloca em colUsuarios
    lErro = CF("Codigos_Le","Usuario", "NomeReduzido", TIPO_STR, colUsuarios, STRING_USUARIO_NOMEREDUZIDO)
    If lErro <> SUCESSO Then Error 49209

    'Adiciona na comboUsuario os nomes reduzidos dos Usuarios
    For iIndice = 1 To colUsuarios.Count
        Usuario.AddItem colUsuarios.Item(iIndice)
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = Err

    Select Case Err

        Case 49209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142652)

    End Select

    Exit Function

End Function

Private Function Carrega_UsuariosComAlcada() As Long
'Carrega a Lista de usuários com habilitação para autorizar crédito

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim colUsuarios As New Collection

On Error GoTo Erro_Carrega_UsuariosComAlcada

    'Le os usuarios que possuem alcada
    lErro = Usuarios_Alcada_Le(colUsuarios)
    If lErro <> SUCESSO Then Error 57273

    'Carrega nomes reduzidos de usuários com alçada
    For Each objUsuario In colUsuarios
        'Adiciona nome reduzido do usuario na listbox
        UsuariosComAlcada.AddItem objUsuario.sNomeReduzido
    Next

    Carrega_UsuariosComAlcada = SUCESSO

    Exit Function

Erro_Carrega_UsuariosComAlcada:

    Carrega_UsuariosComAlcada = Err

    Select Case Err

        Case 57273

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142653)

    End Select

    Exit Function

End Function

Private Function Traz_Alcada_Tela(objAlcada As ClassAlcada) As Long
'Traz os dados da alcada para tela

Dim lErro As Long
Dim iIndice As Integer
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_Traz_Alcada_Tela

    'Limpa a tela
    Call Limpa_Tela_Alcada

    objUsuario.sCodUsuario = objAlcada.sCodUsuario
    'Le os dados do usuario
    lErro = CF("Usuario_Le",objUsuario)
    If lErro <> SUCESSO And lErro <> 36347 Then Error 57274
    If lErro = 36347 Then Error 57275

   'Preenche a tela com os dados de objAlcada
    For iIndice = 0 To Usuario.ListCount - 1
        If objUsuario.sNomeReduzido = Usuario.List(iIndice) Then
            Usuario.ListIndex = iIndice
            Exit For
        End If
    Next

    'Preenche LimiteMensal e LimiteOperacao do Usuario em questao
    If objAlcada.dLimiteMensal > 0 Then LimiteMensal.Text = objAlcada.dLimiteMensal
    If objAlcada.dLimiteOperacao > 0 Then LimiteOperacao.Text = objAlcada.dLimiteOperacao

    iAlterado = 0

    Exit Function

Erro_Traz_Alcada_Tela:

    Select Case Err

        Case 57274
            'Erro tratado na rotina chamada

        Case 57275
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142654)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objAlcada As New ClassAlcada

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Alcada"

    'Le os dados da tela
    lErro = Move_Tela_Memoria(objAlcada)
    If lErro <> SUCESSO Then Error 49203

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodUsuario", objAlcada.sCodUsuario, STRING_ALCADA_CODUSUARIO, "CodUsuario"
    colCampoValor.Add "LimiteOperacao", objAlcada.dLimiteOperacao, 0, "LimiteMensal"
    colCampoValor.Add "LimiteMensal", objAlcada.dLimiteMensal, 0, "LimiteOperacao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 49203
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142655)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objAlcada As New ClassAlcada
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objAlcada com os dados passados em colCampoValor
    objAlcada.sCodUsuario = colCampoValor.Item("CodUsuario").vValor
    objAlcada.dLimiteMensal = colCampoValor.Item("LimiteMensal").vValor
    objAlcada.dLimiteOperacao = colCampoValor.Item("LimiteOperacao").vValor
    
    'Verifica se o Código do Usuário está preenchido
    If Len(Trim(objAlcada.sCodUsuario)) <> 0 Then

        'Traz os dados da alcada para tela
        lErro = Traz_Alcada_Tela(objAlcada)
        If lErro <> SUCESSO Then Error 49158

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 49158 'Usuario sem alcada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142656)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objAlcada As ClassAlcada) As Long
'Recolhe os dados da tela

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objAlcada
    objAlcada.sCodUsuario = Codigo.Caption
    objAlcada.dLimiteMensal = StrParaDbl(LimiteMensal.Text)
    objAlcada.dLimiteOperacao = StrParaDbl(LimiteOperacao.Text)
    objAlcada.sNomeUsuario = Nome.Caption

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142657)

    End Select

    Exit Function

End Function

Function Usuarios_Alcada_Le(colUsuarios2 As Collection) As Long
'Le os usuários que tem alçada

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objAlcada As New ClassAlcada
Dim colUsuarios As New Collection
Dim colAlcada As New Collection

On Error GoTo Erro_Usuarios_Alcada_Le

    'Le todos os usuários contidos na tabela de Usuario e coloca os dados em colUsuarios
    lErro = CF("Usuario_Le_Todos",colUsuarios)
    If lErro <> SUCESSO Then Error 57272

    'Guarda em colAlcada todas as alçadas cadastradas
    lErro = CF("Alcada_Le_Todas",colAlcada)
    If lErro <> SUCESSO Then Error 49210

    For Each objAlcada In colAlcada

        For Each objUsuario In colUsuarios

            'Verifica se CodUsuario em Alcada é igual ao CodUsuario em Usuario
            If UCase(objAlcada.sCodUsuario) = UCase(objUsuario.sCodUsuario) Then

                'Adiciona o Usuario que tem alcada em colUsuarios2
                colUsuarios2.Add objUsuario
            
            End If
        Next
    Next

    Usuarios_Alcada_Le = SUCESSO

    Exit Function

Erro_Usuarios_Alcada_Le:

    Usuarios_Alcada_Le = Err

    Select Case Err

        Case 49210, 57272

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142658)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objAlcada As ClassAlcada) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de Alcada

Dim lErro As Long
Dim iIndice As Integer
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_Trata_Parametros

    If Not (objAlcada Is Nothing) Then

        'Le alcada no BD
        lErro = CF("Alcada_Le",objAlcada)
        If lErro <> SUCESSO And lErro <> 49208 Then Error 49159
        If lErro = SUCESSO Then

            'Traz os dados da alcada para a tela
            lErro = Traz_Alcada_Tela(objAlcada)
            If lErro <> SUCESSO Then Error 49160

        Else

            objUsuario.sCodUsuario = objAlcada.sCodUsuario
            'Le Usuario cujo codigo é igual a objAlcada.sCodUsuario
            lErro = CF("Usuario_Le",objUsuario)
            If lErro <> SUCESSO And lErro <> 36347 Then Error 57276
            If lErro = 36347 Then Error 57277

            'Exibe apenas o usuario passado
            For iIndice = 0 To Usuario.ListCount
                If Usuario.List(iIndice) = objUsuario.sNomeReduzido Then
                    Usuario.ListIndex = iIndice
                    Exit For
                End If
            Next

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 49159, 49160, 57276

        Case 57277
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", Err, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142659)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Exclui alçada

Dim lErro As Long
Dim iIndice As Integer
Dim objAlcada As New ClassAlcada
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_BotaoExcluir_Click

    If Len(Trim(Usuario.Text)) = 0 Then Error 49174

    objAlcada.sCodUsuario = Codigo.Caption
    objUsuario.sCodUsuario = Codigo.Caption

    'Verifica se o usuario tem alcada
    lErro = CF("Alcada_Le",objAlcada)
    If lErro <> SUCESSO And lErro <> 49208 Then Error 49175
    If lErro = 49208 Then Error 49272

    'Pede a confirmação da exclusão da alçada do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ALCADA", objAlcada.sCodUsuario)
    If vbMsgRes = vbYes Then

        'Exlcui a alcada do usuario
        lErro = CF("Alcada_Exclui",objAlcada)
        If lErro <> SUCESSO Then Error 49204

        'Limpa a tela
        Call Limpa_Tela_Alcada

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Le o usuario
        lErro = CF("Usuario_Le",objUsuario)
        If lErro <> SUCESSO And lErro <> 36347 Then Error 57278
        If lErro = 36347 Then Error 57279

        'Remove o usuário da lista de usuários com alçada
        Call Exclui_Lista(objUsuario)

        iAlterado = 0
        
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 49174
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO2", Err)

        Case 49175, 49204, 57278
                'Erro tratado na rotina chamada

        Case 49272
            Call Rotina_Erro(vbOKOnly, "ERRO_ALCADA_NAO_CADASTRADA2", Err, objAlcada.sCodUsuario)

        Case 57279
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", Err, objUsuario.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142660)

    End Select

    Exit Sub

End Sub

Private Sub Exclui_Lista(objUsuario As ClassUsuario)
'Remove o usuário da lista de usuários com alçada

Dim iIndice As Integer

    For iIndice = 0 To UsuariosComAlcada.ListCount - 1
        If UsuariosComAlcada.List(iIndice) = objUsuario.sNomeReduzido Then
            UsuariosComAlcada.RemoveItem iIndice
            Exit For
        End If
    Next

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'Grava uma alçada

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava uma alçada
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 49155

    'Limpa a tela
    Call Limpa_Tela_Alcada

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 49155
            ' Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142661)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 49152

    'Limpa a tela
    Call Limpa_Tela_Alcada

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 49152
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142662)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Alcada()

    Call Limpa_Tela(Me)

    'Limpa o restante da tela
    Codigo.Caption = ""
    Nome.Caption = ""
    Usuario.ListIndex = -1
    
    Exit Sub

End Sub

Private Sub LimiteMensal_GotFocus()

    Call MaskEdBox_TrataGotFocus(LimiteMensal, iAlterado)
        
End Sub

Private Sub LimiteOperacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(LimiteOperacao, iAlterado)
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Usuario Then
            Call UsuariosLabel_Click
        End If
    End If

End Sub

Private Sub Usuario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Usuario_Click()

Dim lErro As Long
Dim objAlcada As New ClassAlcada
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_CodUsuario_Click

    'Verifica se algum usuario esta selecionado
    If Usuario.ListIndex = -1 Then Exit Sub

    'Coloca o nome reduzido selecionado nos obj's
    objUsuario.sNomeReduzido = Usuario.List(Usuario.ListIndex)

    'Le o nome do Usuário
    lErro = CF("Usuario_Le_NomeRed",objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then Error 49165
    If lErro = 57269 Then Error 49166

    'Coloca o nome do usuário na tela
    Nome.Caption = objUsuario.sNome

    'Coloca o código do usuário na tela
    Codigo.Caption = objUsuario.sCodUsuario
    objAlcada.sCodUsuario = objUsuario.sCodUsuario

    lErro = CF("Alcada_Le",objAlcada)
    If lErro <> SUCESSO And lErro <> 49208 Then Error 49167
    
    'Se encontrar alçada para o usuario, coloca os dados da alcada na tela
    If lErro = SUCESSO Then

        LimiteMensal.Text = objAlcada.dLimiteMensal
        LimiteOperacao.Text = objAlcada.dLimiteOperacao
        
    Else
    
        LimiteMensal.Text = ""
        LimiteOperacao.Text = ""
        
    End If

    Exit Sub

Erro_CodUsuario_Click:

    Select Case Err

        Case 49165, 49167

        Case 49166
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", Err, objUsuario.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142663)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava uma alcada

Dim lErro As Long
Dim objAlcada As New ClassAlcada
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Usuario esta preenchido
    If Len(Trim(Usuario.Text)) = 0 Then Error 49169

    'Verifica se o Limite de Operacao esta preenchido
    If Len(Trim(LimiteOperacao.Text)) = 0 Then Error 49170

    'Verifica se o Limite Mensal esta preenchido
    If Len(Trim(LimiteMensal.Text)) = 0 Then Error 49171

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objAlcada)
    If lErro <> SUCESSO Then Error 49172

    If objAlcada.dLimiteOperacao > objAlcada.dLimiteMensal Then Error 51346
    
    objUsuario.sNomeReduzido = Usuario.List(Usuario.ListIndex)

    lErro = Trata_Alteracao(objAlcada, objAlcada.sCodUsuario)
    If lErro <> SUCESSO Then Error 32293

    'Grava uma alçada
    lErro = CF("Alcada_Grava",objAlcada)
    If lErro <> SUCESSO Then Error 49173

    Call Exclui_Lista(objUsuario)

    'Adiciona na listbox se necessário
    Call Adiciona_Lista(objUsuario)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err
    
        Case 32293

        Case 49169
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO2", Err)

        Case 49170
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEOPERACAO_NAO_INFORMADO", Err)

        Case 49171
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEMENSAL_NAO_INFORMADO", Err)

        Case 49172, 49173
            'Erros tratados nas rotinas chamadas

        Case 51346
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_MENSAL_MENOR_LIMITE_OPERACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142664)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub Adiciona_Lista(objUsuario As ClassUsuario)

    UsuariosComAlcada.AddItem objUsuario.sNomeReduzido

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoUsuario = New AdmEvento

    'Carrega a listbox com usuários que possuem alçada
    lErro = Carrega_UsuariosComAlcada()
    If lErro <> SUCESSO Then Error 49153

    'Carrega a combobox todos os usuários
    lErro = Carrega_Usuarios()
    If lErro <> SUCESSO Then Error 49154

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 49153, 49154

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142665)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoUsuario = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Exit Sub

End Sub

Private Sub LimiteMensal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LimiteMensal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteMensal_Validate

    If Len(Trim(LimiteMensal.Text)) = 0 Then Exit Sub

    'Faz a critica do valor de Limite Mensal
    lErro = Valor_Positivo_Critica(LimiteMensal.Text)
    If lErro <> SUCESSO Then Error 49162

    'Coloca o valor formatado na tela
    LimiteMensal.Text = Format(LimiteMensal.Text, "Standard")

    Exit Sub

Erro_LimiteMensal_Validate:

    Cancel = True

    Select Case Err

        Case 49162
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142666)

    End Select

    Exit Sub

End Sub

Private Sub LimiteOperacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LimiteOperacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteOperacao_Validate

    If Len(Trim(LimiteOperacao.Text)) = 0 Then Exit Sub

    'Faz a critica do valor de LimiteOperacao
    lErro = Valor_Positivo_Critica(LimiteOperacao.Text)
    If lErro <> SUCESSO Then Error 49161

    'Coloca o valor no formato 'standard'da tela
    LimiteOperacao.Text = Format(LimiteOperacao.Text, "Standard")

    Exit Sub

Erro_LimiteOperacao_Validate:

    Cancel = True

    Select Case Err

        Case 49161
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142667)

    End Select

    Exit Sub

End Sub

Private Sub objEventoUsuario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAlcada As New ClassAlcada
Dim objUsuario As ClassUsuarios

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    Call Limpa_Tela_Alcada

    Set objUsuario = obj1

    objAlcada.sCodUsuario = objUsuario.sCodUsuario

    'Verifica se o usuario tem alcada
    lErro = CF("Alcada_Le",objAlcada)
    If lErro <> SUCESSO And lErro <> 49208 Then Error 49157
    
    'Coloca os dados do usuário na tela
    lErro = Traz_Alcada_Tela(objAlcada)
    If lErro <> SUCESSO Then Error 49275
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_ObjEventoUsuario_evSelecao:

    Select Case Err

        Case 49157, 49275

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142668)

    End Select

    Exit Sub

End Sub

Private Sub UsuariosComAlcada_DblClick()

Dim lErro As Long
Dim objAlcada As New ClassAlcada
Dim objUsuario As New ClassUsuario

On Error GoTo Erro_UsuariosComAlcada_DblClick

    objUsuario.sNomeReduzido = UsuariosComAlcada.List(UsuariosComAlcada.ListIndex)

    lErro = CF("Usuario_Le_NomeRed",objUsuario)
    If lErro <> SUCESSO And lErro <> 57269 Then Error 57270

    'Se nao encontrou usuario => erro
    If lErro = 57269 Then Error 57271

    objAlcada.sCodUsuario = objUsuario.sCodUsuario

    'Le a alcada do usuario selecionado
    lErro = CF("Alcada_Le",objAlcada)
    If lErro <> SUCESSO And lErro <> 49208 Then Error 49163
    
    'Se nao encontrou => erro
    If lErro = 49208 Then Error 49276

    'Traz para a tela os dados da alcada do usuario selecionado
    lErro = Traz_Alcada_Tela(objAlcada)
    If lErro <> SUCESSO Then Error 49164

    Exit Sub

Erro_UsuariosComAlcada_DblClick:

    Select Case Err

        Case 49163, 49164

        Case 49276
            Call Rotina_Erro(vbOKOnly, "ERRO_ALCADA_NAO_CADASTRADA", Err, objUsuario.sNomeReduzido)

        Case 57270

        Case 57271
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", Err, objUsuario.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142669)

    End Select

    Exit Sub

End Sub
Private Sub UsuariosLabel_Click()

Dim objUsuario As New ClassUsuarios
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_UsuariosLabel_Click

    'Preenche o codigo do usuario com o codigo da tela
    objUsuario.sCodUsuario = Codigo.Caption

    'Chama a tela UsuarioLista
    Call Chama_Tela("UsuarioLista", colSelecao, objUsuario, objEventoUsuario)

    Exit Sub

Erro_UsuariosLabel_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142670)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Alçada"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Alcada"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****
Private Sub UsuariosLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UsuariosLabel, Source, X, Y)
End Sub

Private Sub UsuariosLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UsuariosLabel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Nome_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nome, Source, X, Y)
End Sub

Private Sub Nome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nome, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

