VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form UsuarioTela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuário"
   ClientHeight    =   5070
   ClientLeft      =   945
   ClientTop       =   2100
   ClientWidth     =   9000
   Icon            =   "Usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3870
      Index           =   1
      Left            =   165
      TabIndex        =   19
      Top             =   960
      Width           =   8520
      Begin VB.Frame SSFrame1 
         Caption         =   "Usuário"
         Height          =   3105
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   5505
         Begin VB.TextBox Email 
            Height          =   345
            Left            =   1920
            TabIndex        =   10
            Top             =   2565
            Width           =   3510
         End
         Begin VB.TextBox Senha 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   1635
            Width           =   1095
         End
         Begin VB.CheckBox Ativo 
            Caption         =   "Ativo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3540
            TabIndex        =   11
            Top             =   270
            Width           =   825
         End
         Begin MSMask.MaskEdBox Nome 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   720
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   315
            Left            =   3165
            TabIndex        =   26
            Top             =   2115
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataValidade 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Top             =   2115
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   300
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Top             =   1170
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
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
            Index           =   13
            Left            =   1200
            TabIndex        =   32
            Top             =   2655
            Width           =   585
         End
         Begin VB.Label Label2 
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
            Left            =   390
            TabIndex        =   31
            Top             =   1230
            Width           =   1425
         End
         Begin VB.Label Label6 
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
            Left            =   1125
            TabIndex        =   30
            Top             =   375
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data Validade:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   525
            TabIndex        =   29
            Top             =   2160
            Width           =   1275
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1125
            TabIndex        =   28
            Top             =   1695
            Width           =   645
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1200
            TabIndex        =   27
            Top             =   795
            Width           =   585
         End
      End
      Begin MSComctlLib.TreeView Usuarios 
         Height          =   3090
         Left            =   5805
         TabIndex        =   13
         Top             =   645
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5450
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox Grupo 
         Height          =   315
         Left            =   2055
         TabIndex        =   4
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo-Usuário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5805
         TabIndex        =   20
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
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
         Left            =   1365
         TabIndex        =   21
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3870
      Index           =   2
      Left            =   150
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton DesselecionarTodas 
         Caption         =   "Desmarcar Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4305
         Picture         =   "Usuario.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3240
         Width           =   1830
      End
      Begin VB.CommandButton SelecionaTodas 
         Caption         =   "Marcar Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2100
         Picture         =   "Usuario.frx":132C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3240
         Width           =   1830
      End
      Begin VB.ListBox FiliaisEmpresa 
         Height          =   2085
         Left            =   1485
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   960
         Width           =   5625
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   615
         TabIndex        =   22
         Top             =   270
         Width           =   720
      End
      Begin VB.Label LabelFiliaisEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Filiais-Empresa"
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
         Left            =   1485
         TabIndex        =   23
         Top             =   750
         Width           =   1275
      End
      Begin VB.Label UsuarioLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1485
         TabIndex        =   24
         Top             =   240
         Width           =   5625
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6615
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Usuario.frx":2346
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Usuario.frx":24C4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Usuario.frx":29F6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Usuario.frx":2B80
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4455
      Left            =   90
      TabIndex        =   14
      Top             =   510
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Acesso a Filiais/Empresas"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "UsuarioTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constante da tela
Const TAB_Identificacao = 1
Const TAB_ACESSO = 2

Dim iFrameAtual As Integer
Private gcolFilEmp As Collection 'contem objetos do tipo objUsuarioEmpresa com todas as empresas-filiais em ordem alfabetica

Dim iAlterado As Integer

Public bSenhaAlterada As Boolean

Private Sub TreeView_Modifica(ByVal sGrupo As String, ByVal sCodUsuario As String, Usuarios As TreeView)
'Modifica posição de Usuário na TreeView se Grupo mudou

Dim objNode As Node
Dim objNode1 As Node
Dim iIndice As Integer
Dim bFilho As Boolean

    'Flag que determina se sCodUsuario está como filho de sGrupo
    bFilho = False

    'Tenta localizar o nó sCodUsuario como filho de sGrupo
    'Varre os nós procurando o nó sGrupo
    For Each objNode In Usuarios.Nodes

        'Se encontrou nó sGrupo
        If objNode.Parent Is Nothing And objNode.Text = sGrupo Then

            'Se sGrupo tem filhos analisa os filhos
            If Not objNode.Child Is Nothing Then

                'Referencia o primeiro filho de sGrupo
                Set objNode1 = objNode.Child

                If objNode1.Text = sCodUsuario Then

                    bFilho = True
                    Exit For

                End If

                'Pesquisa os outros filhos de sGrupo (se houver)
                If objNode.Children > 1 Then

                    For iIndice = 1 To objNode.Children - 1

                        Set objNode1 = objNode1.Next

                        If objNode1.Text = sCodUsuario Then

                            bFilho = True
                            Exit For

                        End If

                    Next

                    If bFilho Then Exit For

                End If

            End If

        End If

    Next

    'Se não encontrou sCodUsuario como filho de sGrupo
    If Not bFilho Then

        'Remove sCodUsuario da TreeView
        Usuarios.Nodes.Remove (KEY_CARACTER2 & sCodUsuario)

        'Adiciona sCodUsuario como filho de sGrupo na TreeView
        Call TreeView_Adiciona(sGrupo, sCodUsuario, Usuarios)

    End If

End Sub

Private Sub TreeView_Adiciona(ByVal sGrupo As String, ByVal sCodUsuario As String, Usuarios As TreeView)
'Adiciona sCodUsuario como nó filho de sGrupo na TreeView

Dim objNode As Node
Dim objNode1 As Node
Dim bAdicionado As Boolean

    'Flag que determina se sCodUsuario foi adicionado
    bAdicionado = False

    'Varre os nós procurando o nó sGrupo
    For Each objNode In Usuarios.Nodes

        'Se nó é sGrupo
        If objNode.Parent Is Nothing And objNode.Text = sGrupo Then

            'Cria nó sCodUsuario
            Set objNode1 = Usuarios.Nodes.Add(KEY_CARACTER & sGrupo, tvwChild)
            objNode1.Key = KEY_CARACTER2 & sCodUsuario
            objNode1.Text = sCodUsuario
            'Ordena o nó alfabeticamente na TreeView
            objNode1.Parent.Sorted = True

            bAdicionado = True
            Exit For

        End If

    Next

    'Se Usuário não foi adicionado,
    If Not bAdicionado Then

        'Cria nó sGrupo
        Set objNode = Usuarios.Nodes.Add()
        objNode.Key = KEY_CARACTER & sGrupo
        objNode.Text = sGrupo
        'Ordena o nó alfabeticamente na TreeView
        Usuarios.Sorted = True

        'Cria nó sCodUsuario
        Set objNode = Usuarios.Nodes.Add(KEY_CARACTER & sGrupo, tvwChild)
        objNode.Key = KEY_CARACTER2 & sCodUsuario
        objNode.Text = sCodUsuario

    End If

End Sub

Private Function TreeView_Grupo_Modifica(sGrupo As String, NodeUsuario As Node) As Long
'Modifica a TreeView de modo a incluir NodeUsuario como filho de sGrupo (sGrupo pode não existir na TreeView)
'Rotina chamada por nodeclick na TreeView quando o Grupo do Usuário no BD não coincide com o Grupo na TreeView

Dim lErro As Long
Dim objNode As Node
Dim bModificado As Boolean
Dim colUsuario As New Collection
Dim vUsuario As Variant

On Error GoTo Erro_TreeView_Grupo_Modifica

    bModificado = False 'Flag que controla se TreeView foi modificada

    For Each objNode In Usuarios.Nodes

        'Testa se nó é tipo Grupo e coincide com sGrupo
        If objNode.Parent Is Nothing And objNode.Text = sGrupo Then

            'Adiciona usuário da TreeView (nova posição)
            Usuarios.Nodes.Add objNode.Key, tvwChild, KEY_CARACTER2 & NodeUsuario.Text, NodeUsuario.Text
            'Ordena o nó alfabeticamente na TreeView
            objNode.Sorted = True
            'Remove usuário da TreeView (posição antiga)
            Usuarios.Nodes.Remove (NodeUsuario.Index)
            bModificado = True

            Exit For

        End If

    Next

    If Not bModificado Then  'Nao há sGrupo na TreeView

        'Remove Usuario da TreeView
        Usuarios.Nodes.Remove (NodeUsuario.Index)

        'Adiciona sGrupo na TreeView
        Set objNode = Usuarios.Nodes.Add()
        objNode.Key = KEY_CARACTER & sGrupo
        objNode.Text = sGrupo
        'Ordena o nó alfabeticamente na TreeView
        Usuarios.Sorted = True

        'Le os Usuarios de sGrupo
        lErro = Usuarios_Le_Grupo(sGrupo, colUsuario)
        If lErro Then Error 6241

        'Acrescenta Usuários de sGrupo na TreeView
        For Each vUsuario In colUsuario
            Set objNode = Usuarios.Nodes.Add(KEY_CARACTER & sGrupo, tvwChild)
            objNode.Key = KEY_CARACTER2 & vUsuario
            objNode.Text = vUsuario
        Next

        'Ordena os Usuários alfabeticamente
        Usuarios.Nodes.Item(KEY_CARACTER & sGrupo).Sorted = True

    End If

    TreeView_Grupo_Modifica = SUCESSO

    Exit Function

Erro_TreeView_Grupo_Modifica:

    TreeView_Grupo_Modifica = Err

    Select Case Err

        Case 6241

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175633)

    End Select

    Exit Function


End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objUsuario As New ClassDicUsuario
'Dim sGrupo As String
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se código do Usuário foi informado
    If Len(Codigo.Text) = 0 Then Error 6272

    'Preenche o código de Usuário em objUsuario
    objUsuario.sCodUsuario = Codigo.Text

    lErro = DicUsuario_Le(objUsuario)
    If lErro = 6223 Then Error 6273  'Usuario não cadastrado
    If lErro <> SUCESSO Then Error 6274

    'Pede confirmação para exclusão
    vbMsgRet = Rotina_Aviso(vbYesNo, "EXCLUSAO_USUARIO", objUsuario.sCodUsuario, objUsuario.sNome)

    If vbMsgRet = vbYes Then

        'Exclui o Usuário do BD
        lErro = Usuario_Exclui(objUsuario.sCodUsuario)
        If lErro Then Error 6275

        'Exclui o Usuário da TreeView
        Usuarios.Nodes.Remove (KEY_CARACTER2 & objUsuario.sCodUsuario)

        'Limpa a Tela
        lErro = LimpaTelaUsuario()
        If lErro <> SUCESSO Then Error 25941

        'Exibe data default
        Call DateParaMasked(DataValidade, DATA_NULA)
        
        iAlterado = 0

    End If

Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 6272
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_USUARIO_NAO_INFORMADO", Err)
            Codigo.SetFocus

        Case 6273
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, objUsuario.sCodUsuario)
            Codigo.SetFocus

        Case 6274, 6275, 25941   'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175634)

     End Select

     Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload UsuarioTela
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objUsuario As New ClassDicUsuario
Dim sGrupo As String
Dim lCodigo As Long
Dim iOperacao As Integer
Dim colUsuFilEmp As New Collection

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se Grupo foi informado
    If Len(Grupo.Text) = 0 Then gError 6242

    'Verifica se dados do Usuário foram informados
    If Len(Trim(Codigo.Text)) = 0 Then gError 6243
    If Len(Trim(Nome.Text)) = 0 Then gError 6244
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 32203
    If Len(Senha.Text) = 0 Then gError 6245
    
    If bSenhaAlterada Then gError 134272

    'Preenche sGrupo
    sGrupo = Grupo.Text

    'Preenche objUsuario
    objUsuario.sCodUsuario = Trim(Codigo.Text)
    objUsuario.sCodGrupo = sGrupo
    objUsuario.sNome = Trim(Nome.Text)
    objUsuario.sNomeReduzido = Trim(NomeReduzido.Text)
    objUsuario.sSenha = Senha.Text
    If DataValidade.Text = "  /  /  " Then
        objUsuario.dtDataValidade = DATA_NULA
    Else
        objUsuario.dtDataValidade = CDate(DataValidade.Text)
    End If
    objUsuario.iAtivo = IIf(Ativo.Value, ATIVIDADE, INATIVIDADE)
    objUsuario.sEmail = Email.Text

    'preencher colUsuFilEmp com as filiais selecionadas
    lErro = Preenche_UsuFilEmp(objUsuario.sCodUsuario, colUsuFilEmp)
    If lErro <> SUCESSO Then gError 32039
    
    'grava o Usuário no banco de dados
    lErro = DicUsuario_Grava(objUsuario, colUsuFilEmp, iOperacao)
    If lErro Then gError 6246

    Select Case iOperacao

        Case GRAVACAO    'GRAVACAO ou REPLICACAO de Usuário

            'Adiciona o Usuário a TreeView
            Call TreeView_Adiciona(sGrupo, objUsuario.sCodUsuario, Usuarios)

        Case MODIFICACAO

            'Modifica o Usuário na TreeView, se grupo mudou
            Call TreeView_Modifica(sGrupo, objUsuario.sCodUsuario, Usuarios)

    End Select

    'Limpa a Tela
    lErro = LimpaTelaUsuario()
    If lErro <> SUCESSO Then gError 25939

    'Exibe data default
    Call DateParaMasked(DataValidade, DATA_NULA)
    
    iAlterado = 0

Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 6242
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_GRUPO_NAO_INFORMADO", gErr)
            
        Case 6243
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_USUARIO_NAO_INFORMADO", gErr)

        Case 6244
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_USUARIO_NAO_INFORMADO", gErr)

        Case 6245
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_USUARIO_NAO_INFORMADA", gErr)

        Case 32203
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_USUARIO_NAO_INFORMADO", gErr)
        
        Case 134272
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_ALTERADA_NAO_CONFIRMADA", gErr)

        Case 6246, 25939  'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175635)

     End Select

     Exit Sub

End Sub
Private Function LimpaTelaUsuario()

Dim iItem As Integer

On Error GoTo Erro_LimpaTelaUsuario

    If bSenhaAlterada Then gError 134266
    
    iAlterado = 0
    
    Call Limpa_Tela(UsuarioTela)
    
    If FiliaisEmpresa.ListCount > 0 Then
        
        For iItem = 0 To FiliaisEmpresa.ListCount - 1
        
            FiliaisEmpresa.Selected(iItem) = False
        
        Next
        
    End If
    
    bSenhaAlterada = False

    LimpaTelaUsuario = SUCESSO
    
    Exit Function

Erro_LimpaTelaUsuario:

    LimpaTelaUsuario = gErr
    
    Select Case gErr
    
        Case 134266
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_ALTERADA_NAO_CONFIRMADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175636)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_LimpaTelaUsuario
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134271
    
    lErro = LimpaTelaUsuario()
    If lErro <> SUCESSO Then gError 25940

    'Exibe data default
    Call DateParaMasked(DataValidade, DATA_NULA)
    
    iAlterado = 0

    Exit Sub

Erro_LimpaTelaUsuario:

    Select Case gErr

        Case 25940, 134271 'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175637)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

Dim lErro As Long

On Error GoTo Erro_Codigo_Change

    iAlterado = REGISTRO_ALTERADO

    lErro = Preenche_UsuarioLabel(Codigo.Text, Nome.Text)
    If lErro <> SUCESSO Then Error 25937

    Exit Sub

Erro_Codigo_Change:

    Select Case Err

        Case 25937  'Tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175638)

    End Select

    Exit Sub
        
End Sub
Private Function Preenche_UsuarioLabel(sCodigo As String, sNome As String) As Long

Dim lErro As Long

On Error GoTo Erro_Preenche_UsuarioLabel

    If Len(Trim(sCodigo) & Trim(sNome)) > 0 Then
        UsuarioLabel.Caption = Trim(sCodigo) & " - " & Trim(sNome)
    Else
        UsuarioLabel.Caption = ""
    End If

    Preenche_UsuarioLabel = SUCESSO

    Exit Function

Erro_Preenche_UsuarioLabel:

    Preenche_UsuarioLabel = Err
    
    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175639)

    End Select

    Exit Function
        
End Function



Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DataValidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataValidade)

End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_DataValidade_Validate

    If DataValidade.Text <> "  /  /  " Then

        lErro = Data_Critica(DataValidade.Text)
        If lErro Then Error 6233

        dtData = CDate(DataValidade.Text)
        'Compara DataAtual com DataValidade
        If DateDiff("d", Now, dtData) <= 0 Then Error 6234

    End If

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True
    
    Select Case Err

        Case 6233
    
        Case 6234
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, DataValidade.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175640)

    End Select

    Exit Sub

End Sub

Private Sub DesselecionarTodas_Click()

Dim iItem As Integer

    For iItem = 0 To FiliaisEmpresa.ListCount - 1
    
        FiliaisEmpresa.Selected(iItem) = False
    
    Next
    
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim objUsuario As New ClassDicUsuario
Dim colGrupo As New Collection
Dim colUsuario As Collection
Dim vGrupo As Variant
'Dim sGrupo As String
Dim vUsuario As Variant
Dim objNode As Node
Dim objGrupo As New ClassDicGrupo

On Error GoTo Erro_Usuario_Form_Load

    Me.HelpContextID = TAB_Identificacao
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Opcao.Tabs(TAB_ACESSO).Caption = "Acesso a Empresas"
'        LabelFiliaisEmpresa.Caption = "Empresas"
'
'    End If
    
    Email.MaxLength = STRING_EMAIL
   
    iFrameAtual = 1
    
    'carregar todas as filiais/empresa possiveis na listbox do frame2
    Call Carrega_ListaFilEmp
    
    'Preenche a coleção colGrupo com Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 6366 'Não existem Grupos
    If lErro Then Error 6215

    'Acrescenta os Grupos na TreeView
    For Each vGrupo In colGrupo
        Set objNode = Usuarios.Nodes.Add()
        objNode.Key = KEY_CARACTER & vGrupo
        objNode.Text = vGrupo
        objNode.Expanded = True
    Next

    'Ordena alfabeticamente os Grupos na TreeView
    Usuarios.Sorted = True

    'Para cada grupo preenche a TreeView com Usuários do Grupo do BD.
    For Each vGrupo In colGrupo

        'Cria nova coleção. Libera a anterior
        Set colUsuario = New Collection

        'Preenche a coleção colUsuario com Usuários de vGrupo.
        lErro = Usuarios_Le_Grupo(vGrupo, colUsuario)
        If lErro Then Error 6216

        'Acrescenta Usuários do Grupo na TreeView
        For Each vUsuario In colUsuario
            Set objNode = Usuarios.Nodes.Add(KEY_CARACTER & vGrupo, tvwChild)
            objNode.Key = KEY_CARACTER2 & vUsuario
            objNode.Text = vUsuario
        Next

        'Ordena os Usuários do Grupo alfabeticamente na TreeView
        Usuarios.Nodes.Item(KEY_CARACTER & vGrupo).Sorted = True

    Next

    'Se há um Usuário selecionado, exibir seus dados
    If Len(gsUsuario) > 0 Then

        objUsuario.sCodUsuario = gsUsuario

        'Verifica se o Usuario existe.
        lErro = DicUsuario_Le(objUsuario)

        If lErro <> 6223 And lErro <> SUCESSO Then Error 6217

        If lErro = SUCESSO Then     'Usuário está cadastrado

            'Coloca dados na Tela.
            Call Traz_Usuario_Tela(objUsuario)

            'Seleciona Usuario na TreeView
            Usuarios.Nodes(KEY_CARACTER2 & Codigo.Text).Selected = True

        Else  'Usuário não está cadastrado

             If objUsuario.sCodUsuario <> "0" Then
                Codigo.Text = objUsuario.sCodUsuario
             Else
                Codigo.Text = ""
             End If
             
             Call DateParaMasked(DataValidade, DATA_NULA)

        End If

        gsUsuario = ""
        gsGrupo = ""

    ElseIf Len(gsGrupo) > 0 Then
        
        objGrupo.sCodGrupo = gsGrupo

        'Verifica se o Usuario existe.
        lErro = Grupo_Le(objGrupo)

        If lErro <> 6230 And lErro <> SUCESSO Then Error 25942
        If lErro = SUCESSO Then     'Grupo está cadastrado

            'Coloca dado na Tela.
            Grupo.Text = objGrupo.sCodGrupo

            'Seleciona Usuario na TreeView
            Usuarios.Nodes(KEY_CARACTER & Grupo.Text).Selected = True

        Else  'Grupo não está cadastrado

             If objGrupo.sCodGrupo <> "0" Then
                Grupo.Text = objGrupo.sCodGrupo
             Else
                Grupo.Text = ""
             End If
             
             Call DateParaMasked(DataValidade, DATA_NULA)

        End If

        gsGrupo = ""

    Else    'Não há usuário nem grupo selecionado. Exibe apenas data default.

        Call DateParaMasked(DataValidade, DATA_NULA)

    End If
    
    bSenhaAlterada = False
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Usuario_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 6215, 6216, 6217, 25942   'Tratados na rotina chamada

        Case 6366
            Call Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175641)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click
    
    If bSenhaAlterada Then gError 134267

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134268

    Set gcolFilEmp = Nothing

    Exit Sub
    
Erro_BotaoFechar_Click:

    Select Case gErr
        
        Case 134267
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_ALTERADA_NAO_CONFIRMADA", gErr)
       
        Case 134268 ' erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175642)

    End Select

    Cancel = True

    Exit Sub

End Sub

Private Sub Grupo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Grupo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objGrupo As New ClassDicGrupo
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_Grupo_Validate

    If Len(Grupo.Text) > 0 Then

        objGrupo.sCodGrupo = Grupo.Text

        lErro = Grupo_Le(objGrupo)
        If lErro <> 6230 And lErro <> SUCESSO Then Error 6232
        If lErro = 6230 Then   'Não existe o Grupo.

            'Pergunta se deseja criar.
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_GRUPO_INEXISTENTE", Grupo.Text)

            If vbMsgRet = vbYes Then

                'Ativa a tela de Grupo
                gsGrupo = Grupo.Text
                GrupoForm.Show

            Else  'vbMsgRet = vbNo

                Cancel = True
                
            End If

        End If

    End If

    Exit Sub

Erro_Grupo_Validate:

    Select Case Err

        Case 6232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175643)

    End Select

    Exit Sub

End Sub

Private Sub Nome_Change()

Dim lErro As Long

On Error GoTo Erro_Nome_Change

    iAlterado = REGISTRO_ALTERADO

    lErro = Preenche_UsuarioLabel(Codigo.Text, Nome.Text)
    If lErro <> SUCESSO Then Error 25938

    Exit Sub

Erro_Nome_Change:

    Select Case Err

        Case 25938  'Tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175644)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57827

    End If
                
    Exit Sub

Erro_NomeReduzido_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 57827
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175645)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me, 0) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Me.HelpContextID = IDH_USUARIO_ID
                
            Case TAB_ACESSO
                Me.HelpContextID = IDH_USUARIO_ACESSO_FILIAIS_EMPRESA
                        
        End Select
    
    End If
    
End Sub

Private Sub SelecionaTodas_Click()

Dim iItem As Integer

    For iItem = 0 To FiliaisEmpresa.ListCount - 1
    
        FiliaisEmpresa.Selected(iItem) = True
    
    Next
    
End Sub

Private Sub Senha_Change()

    iAlterado = REGISTRO_ALTERADO
    
    bSenhaAlterada = True

End Sub

Private Sub Senha_Validate(Cancel As Boolean)

    If Len(Senha.Text) = 0 Then Exit Sub
    
    If bSenhaAlterada Then

       ConfirmacaoDeSenha.Show vbModal
       
    End If

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub
    
    sData = DataValidade.Text

    lErro = Data_Diminui(sData)
    If lErro Then Error 6237

    dtData = CDate(sData)
    'Compara DataAtual com DataValidade
    If DateDiff("d", Now, dtData) <= 0 Then Error 6295

    DataValidade.Text = sData

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case Err

        Case 6237  'Já foi tratado na rotina chamada

        Case 6295
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, sData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175646)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub
    
    sData = DataValidade.Text

    lErro = Data_Aumenta(sData)
    If lErro Then Error 6238

    dtData = CDate(sData)
    'Compara DataAtual com DataValidade
    If DateDiff("d", Now, dtData) <= 0 Then Error 6296

    DataValidade.Text = sData

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case Err

        Case 6238  'Já foi tratado na rotina chamada

        Case 6296
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, sData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175647)

    End Select

    Exit Sub

End Sub

Private Sub Usuarios_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objPai As Node
Dim objGrupo As New ClassDicGrupo
Dim objUsuario As New ClassDicUsuario
Dim sGrupo As String

On Error GoTo Erro_Usuarios_NodeClick

    'Seta Nó pai
    Set objPai = Node.Parent

    If objPai Is Nothing Then  'Nó é Grupo

        objGrupo.sCodGrupo = Node.Text
        lErro = Grupo_Le(objGrupo)
        If lErro <> SUCESSO And lErro <> 6230 Then gError 6239

        If lErro = SUCESSO Then    'Grupo está cadastrado.

            'Coloca na Tela
            Grupo.Text = Node.Text

        Else    'Grupo não está cadastrado.

            'Remove da TreeView
            Usuarios.Nodes.Remove (Node.Index)

        End If

    Else   'Nó é Usuário. Nó pai é Grupo.

        objUsuario.sCodUsuario = Node.Text

        'Le Usuario e Grupo que o contém no BD
        lErro = DicUsuario_Le(objUsuario)
        If lErro <> SUCESSO And lErro <> 6223 Then gError 6240

        If lErro = SUCESSO Then     'Usuário está cadastrado

            If bSenhaAlterada Then gError 134269
            
            lErro = Teste_Salva(Me, iAlterado)
            If lErro <> SUCESSO Then gError 134270

            'Coloca dados na Tela
            Call Traz_Usuario_Tela(objUsuario)

            'Se Grupo do Usuário BD <> Grupo na TreeView modifica TreeView
            If objPai.Text <> objUsuario.sCodGrupo Then

                'Modifica TreeView
                lErro = TreeView_Grupo_Modifica(objUsuario.sCodGrupo, Node)

            End If
            
            bSenhaAlterada = False

        Else    'Usuário não está cadastrado

            'Remove Usuário da TreeView
            Usuarios.Nodes.Remove (Node.Index)

        End If

    End If
    
    Exit Sub

Erro_Usuarios_NodeClick:

    Select Case gErr

        Case 6239, 6240, 134270 'Já foi tratado na rotina chamada
    
        Case 134269
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_ALTERADA_NAO_CONFIRMADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175648)

    End Select

    Exit Sub

End Sub

Private Function Preenche_UsuFilEmp(sCodUsuario As String, colUsuFilEmp As Collection) As Long
'preencher colUsuFilEmp com as filiais selecionadas

Dim lErro As Long, objUsuFilEmp As ClassUsuFilEmp, iItem As Integer

On Error GoTo Erro_Preenche_UsuFilEmp

    For iItem = 0 To FiliaisEmpresa.ListCount - 1
    
        If FiliaisEmpresa.Selected(iItem) = True Then
        
        Set objUsuFilEmp = New ClassUsuFilEmp
        
        objUsuFilEmp.sCodUsuario = sCodUsuario
        objUsuFilEmp.lCodEmpresa = gcolFilEmp.Item(iItem + 1).lCodEmpresa
        objUsuFilEmp.iCodFilial = gcolFilEmp.Item(iItem + 1).iCodFilial
        
        Call colUsuFilEmp.Add(objUsuFilEmp)
        
        End If
    
    Next
    
    Preenche_UsuFilEmp = SUCESSO
     
    Exit Function
    
Erro_Preenche_UsuFilEmp:

    Preenche_UsuFilEmp = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175649)
     
    End Select
     
    Exit Function

End Function

Private Function Traz_Usuario_Tela(objUsuario As ClassDicUsuario) As Long
'Preenche tela

Dim lErro As Long, iItem As Integer
Dim colUsuFilEmp As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa

On Error GoTo Erro_Traz_Usuario_Tela

    Grupo.Text = objUsuario.sCodGrupo
    Codigo.Text = objUsuario.sCodUsuario
    Nome.Text = objUsuario.sNome
    NomeReduzido.Text = objUsuario.sNomeReduzido
    Senha.Text = objUsuario.sSenha
    Call DateParaMasked(DataValidade, objUsuario.dtDataValidade)
    Ativo.Value = IIf(objUsuario.iAtivo = ATIVIDADE, vbChecked, vbUnchecked)
    Email.Text = objUsuario.sEmail

    'ler as filiais/empresas a que o usuario tem acesso
    lErro = UsuFilEmp_Le_Usuario(objUsuario.sCodUsuario, colUsuFilEmp)
    If lErro <> SUCESSO Then Error 32045
    
    'marcar na listbox as filiais a que o usuario tem permissao de acesso
    iItem = 0
    For Each objUsuarioEmpresa In gcolFilEmp
            
        If FilEmp_Obtem_Item_Lista(objUsuarioEmpresa.lCodEmpresa, objUsuarioEmpresa.iCodFilial, colUsuFilEmp) <> -1 Then
            FiliaisEmpresa.Selected(iItem) = True
        Else
            FiliaisEmpresa.Selected(iItem) = False
        End If
        iItem = iItem + 1
        
    Next
    
    iAlterado = 0

    Traz_Usuario_Tela = SUCESSO
     
    Exit Function
    
Erro_Traz_Usuario_Tela:

    Traz_Usuario_Tela = Err
     
    Select Case Err
          
        Case 32045
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175650)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_ColFilEmp(colFilEmp As Collection) As Long

Dim lErro As Long, objUsuarioEmpresa As ClassUsuarioEmpresa
Dim lEmpresa As Long, iItem As Integer
Dim objUsuarioEmpresaAux As ClassUsuarioEmpresa

On Error GoTo Erro_Carrega_ColFilEmp

    lErro = FiliaisEmpresas_Le_Todas(colFilEmp)
    If lErro <> SUCESSO Then Error 32040
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'para cada empresa incluir uma linha p/testar acesso "empresa toda"
        lEmpresa = 0
        iItem = 0
        For Each objUsuarioEmpresa In colFilEmp
        
            iItem = iItem + 1
            If objUsuarioEmpresa.lCodEmpresa <> lEmpresa Then
                
                Set objUsuarioEmpresaAux = New ClassUsuarioEmpresa
                objUsuarioEmpresaAux.lCodEmpresa = objUsuarioEmpresa.lCodEmpresa
                objUsuarioEmpresaAux.sNomeEmpresa = objUsuarioEmpresa.sNomeEmpresa
                objUsuarioEmpresaAux.iCodFilial = 0
                objUsuarioEmpresaAux.sNomeFilial = "<Empresa Toda>"
                
                Call colFilEmp.Add(objUsuarioEmpresaAux, Before:=iItem)
                
                lEmpresa = objUsuarioEmpresa.lCodEmpresa
                iItem = iItem + 1
                
            End If
        
        Next
    
'    End If
    
    Carrega_ColFilEmp = SUCESSO
     
    Exit Function
    
Erro_Carrega_ColFilEmp:

    Carrega_ColFilEmp = Err
     
    Select Case Err
          
        Case 32040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175651)
     
    End Select
     
    Exit Function
    
End Function

Private Sub Carrega_ListaFilEmp()

Dim objUsuarioEmpresa As ClassUsuarioEmpresa

    Set gcolFilEmp = New Collection
        
    Call Carrega_ColFilEmp(gcolFilEmp)
    
'    If giTipoVersao = VERSAO_FULL Then
        
        For Each objUsuarioEmpresa In gcolFilEmp
            FiliaisEmpresa.AddItem objUsuarioEmpresa.sNomeEmpresa & " - " & objUsuarioEmpresa.sNomeFilial
        Next
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        For Each objUsuarioEmpresa In gcolFilEmp
'            FiliaisEmpresa.AddItem objUsuarioEmpresa.sNomeEmpresa
'        Next
'
'    End If

End Sub

Private Function FilEmp_Obtem_Item_Lista(lCodEmpresa, iCodFilial, colUsuFilEmp As Collection)
'retorna o indice da Empresa-Filial na colecao colUsuFilEmp
'retorna -1 p/"nao encontrado"

Dim iItem As Integer, objUsuFilEmp As ClassUsuFilEmp

    iItem = 0
    
    For Each objUsuFilEmp In colUsuFilEmp
    
        If objUsuFilEmp.lCodEmpresa = lCodEmpresa And objUsuFilEmp.iCodFilial = iCodFilial Then
        
            FilEmp_Obtem_Item_Lista = iItem
            Exit Function
            
        End If
        
        iItem = iItem + 1
        
    Next

    FilEmp_Obtem_Item_Lista = -1
    
End Function

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function Gravar_Registro() As Long

    Call BotaoGravar_Click

End Function

Private Sub Email_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
