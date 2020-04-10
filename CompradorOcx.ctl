VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CompradorOcx 
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   8565
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6270
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "CompradorOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CompradorOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CompradorOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CompradorOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Compradores 
      Height          =   2595
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1275
      Width           =   2295
   End
   Begin VB.CommandButton BotaoUsuarios 
      Caption         =   "Usuários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   165
      Picture         =   "CompradorOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3345
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1395
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Width           =   5595
      Begin VB.Label Nome 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   4
         Top             =   840
         Width           =   3900
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
         Left            =   600
         TabIndex        =   3
         Top             =   870
         Width           =   555
      End
      Begin VB.Label CodUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   2
         Top             =   330
         Width           =   1080
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
         Left            =   495
         TabIndex        =   1
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprador"
      Height          =   1710
      Left            =   135
      TabIndex        =   5
      Top             =   1575
      Width           =   5595
      Begin VB.TextBox Email 
         Height          =   345
         Left            =   1635
         TabIndex        =   10
         Top             =   1170
         Width           =   3870
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2280
         Picture         =   "CompradorOcx.ctx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Numeração Automática"
         Top             =   330
         Width           =   300
      End
      Begin VB.CheckBox AumentaQuant 
         Caption         =   "Pode aumentar quantidades requisitadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   870
         TabIndex        =   9
         Top             =   780
         Width           =   3855
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1635
         TabIndex        =   7
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
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
         Left            =   915
         TabIndex        =   19
         Top             =   1245
         Width           =   585
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   870
         TabIndex        =   6
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Compradores"
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
      Left            =   6120
      TabIndex        =   18
      Top             =   990
      Width           =   1110
   End
End
Attribute VB_Name = "CompradorOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer

Private WithEvents objEventoUsuario As AdmEvento
Attribute objEventoUsuario.VB_VarHelpID = -1

Private Function Move_Tela_Memoria(objComprador As ClassComprador) As Long
'Lê os dados que estão na tela Comprador e coloca em objComprador

On Error GoTo Erro_Move_Tela_Memoria

    'Se o codigo não estiver vazio coloca-o no objComprador
    objComprador.iCodigo = StrParaInt(Codigo.Text)
    
    objComprador.iAumentaQuant = AumentaQuant.Value

    'Se o CodUsuario não estiver vazio coloca-o no objComprador
    If Len(CodUsuario.Caption) > 0 Then
        objComprador.sCodUsuario = CodUsuario.Caption
        objComprador.sNome = Nome.Caption
    End If

    objComprador.iFilialEmpresa = giFilialEmpresa
    objComprador.sEmail = Email.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154489)

    End Select

    Exit Function

End Function

Private Sub Adiciona_Lista_Comprador(objComprador As ClassComprador)
'Adiciona um comprador na ListBox

Dim lErro As Long
Dim colComprador As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Adiciona_Lista_Comprador
    
    If Len(Trim(objComprador.sNomeReduzido)) > 0 Then
        'Se ele é novo adiciona-o na lista
        Compradores.AddItem objComprador.sNomeReduzido
    End If
    
    Exit Sub

Erro_Adiciona_Lista_Comprador:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154490)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava um registro no bd

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos estão preenchidos
    If Len(CodUsuario.Caption) = 0 Then Error 50070
    If Len(Nome.Caption) = 0 Then Error 50071
    If Len(Codigo.ClipText) = 0 Then Error 50072


    'Transfere os dados da tela para os obj's
    objComprador.sCodUsuario = CodUsuario.Caption
    objComprador.sNome = Nome.Caption
    objComprador.iCodigo = StrParaInt(Codigo.ClipText)
    objComprador.iAumentaQuant = AumentaQuant.Value
    objComprador.iFilialEmpresa = giFilialEmpresa

    'Le o comprador com o usuario da tela
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then Error 50074

    'Se encontrar
    If lErro = SUCESSO Then

        'Verifica se o codigo e o mesmo que o codigo da tela
        If (objComprador.iCodigo <> StrParaInt(Codigo.ClipText)) Then Error 50075

    End If

    'Chama função que armazena os dados da tela no objComprador
    lErro = Move_Tela_Memoria(objComprador)
    If lErro <> SUCESSO Then Error 50073

    lErro = Trata_Alteracao(objComprador, objComprador.iFilialEmpresa, objComprador.iCodigo)
    If lErro <> SUCESSO Then Error 32292

    'Grava a comprador no BD
    lErro = CF("Comprador_Grava", objComprador)
    If lErro <> SUCESSO Then Error 50076

    'Adiciona na listbox se necessário
    Call Adiciona_Lista_Comprador(objComprador)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 32292, 50073, 50074, 50076, 50141

        Case 50070
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)

        Case 50071
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)

        Case 50072
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 50075
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_USUARIO", Err, objComprador.iCodigo, objComprador.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154491)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub Limpa_Tela_Comprador()
'Limpa a tela

    Call Limpa_Tela(Me)

    Nome.Caption = ""
    CodUsuario.Caption = ""
    Codigo.Text = ""
    AumentaQuant.Value = gobjCOM.iCompradorAumentaQuant

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

End Sub

Function Traz_Comprador_Tela(objComprador As ClassComprador) As Long
'Traz os dados do comprador para a tela

Dim iIndice As Integer
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Traz_Comprador_Tela

    'Le o CodUsuario
    objUsuarios.sCodUsuario = objComprador.sCodUsuario

    'Le o Usuario na tabela
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then Error 50123
    If lErro <> SUCESSO Then Error 50135

    'Preenche a tela com os dados de objComprador
    CodUsuario.Caption = objComprador.sCodUsuario
    Nome.Caption = objUsuarios.sNome

    Codigo.Text = objComprador.iCodigo
    AumentaQuant.Value = objComprador.iAumentaQuant
    Email.Text = objComprador.sEmail

    iAlterado = 0

    Traz_Comprador_Tela = SUCESSO

    Exit Function

Erro_Traz_Comprador_Tela:

    Traz_Comprador_Tela = Err

    Select Case Err

        Case 50060, 50134

        Case 50135
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154492)

    End Select

    Exit Function

End Function

Function Comprador_Automatico(iCodigo As Integer) As Long
'Gera o próximo comprador

Dim lErro As Long

On Error GoTo Erro_Comprador_Automatico

    lErro = CF("Config_Obter_Inteiro_Automatico", "ComprasConfig", "NUM_PROX_COMPRADOR", "Comprador", "Codigo", iCodigo)
    If lErro <> SUCESSO Then Error 50122

    Comprador_Automatico = SUCESSO

    Exit Function

Erro_Comprador_Automatico:

    Comprador_Automatico = Err

    Select Case Err

        Case 50122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154493)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objComprador As ClassComprador) As Long
'Trata os parametros

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Comprador preenchido
    If Not (objComprador Is Nothing) Then

        'Se objComprador.codigo > 0
        If objComprador.iCodigo > 0 Then

            'Verifica se o Comprador existe, lendo no BD a partir do  codigo
            lErro = CF("Comprador_Le", objComprador)
            If lErro <> SUCESSO And lErro <> 50064 Then Error 50049

            'Se o Comprador existe
            If lErro = SUCESSO Then
                lErro = Traz_Comprador_Tela(objComprador)
                If lErro <> SUCESSO Then Error 50050

            'Se o Comprador não existe
            ElseIf objComprador.iCodigo > 0 Then

                'Mantém o Código do Comprador na tela
                Codigo.Text = CStr(objComprador.iCodigo)

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 50049, 50050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154494)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Compradores_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colComprador As New Collection

On Error GoTo Erro_Compradores_Carrega

    'Le todos os Usuarios da Colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then Error 50051

    'Le todos os Compradores da Filial Empresa
    lErro = CF("Comprador_Le_Todos", colComprador)
    If lErro <> SUCESSO And lErro <> 50126 Then Error 50122

    For Each objComprador In colComprador
        For Each objUsuarios In colUsuarios
            If objComprador.sCodUsuario = objUsuarios.sCodUsuario Then
                Compradores.AddItem objUsuarios.sNomeReduzido
            End If
        Next
    Next

    Compradores_Carrega = SUCESSO

    Exit Function

Erro_Compradores_Carrega:

    Compradores_Carrega = Err

    Select Case Err

        Case 50051, 50122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154495)

    End Select

    Exit Function

End Function

Private Sub AumentaQuant_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o comprador está preenchido
    If Len(Codigo.Text) = 0 Then Error 50077

    objComprador.iCodigo = CInt(Codigo.Text)

    'Verifica se o usuário tem comprador
    lErro = CF("Comprador_Le", objComprador)
    If lErro <> SUCESSO And lErro <> 50064 Then Error 50078
    If lErro <> SUCESSO Then Error 50079

    'Pede a confirmação da exclusão do comprador do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COMPRADOR", objComprador.sCodUsuario)
    If vbMsgRes = vbYes Then

        lErro = Move_Tela_Memoria(objComprador)
        If lErro <> SUCESSO Then Error 50080
    
        lErro = CF("Comprador_Exclui", objComprador)
        If lErro <> SUCESSO Then Error 50081
    
        Call Limpa_Tela_Comprador
    
        objUsuarios.sCodUsuario = objComprador.sCodUsuario
    
        'Lê o nome do usuário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO Then Error 50082
    
        'Procura o índice da comprador
        For iIndice = 0 To Compradores.ListCount - 1
            If Compradores.List(iIndice) = objUsuarios.sNomeReduzido Then
                Compradores.RemoveItem iIndice
                Exit For
            End If
        Next

        iAlterado = 0
    
    End If
    
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 50077
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 50078, 50080, 50081, 50082

        Case 50079
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO", Err, objComprador.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154496)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_BotaoGravar_Click

    'Grava os registros na tabela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 50069

    Call Limpa_Tela_Comprador

    iAlterado = 0

    Exit Sub
Erro_BotaoGravar_Click:

    Select Case Err

        Case 50069

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154497)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 50087

    Call Limpa_Tela_Comprador

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 50087

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154498)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
        
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
        
    'Gera Código da proximo Comprador
    lErro = Comprador_Automatico(iCodigo)
    If lErro <> SUCESSO Then Error 50050

    Codigo.Text = iCodigo
    
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case Err
    
        Case 50050
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154499)
    
    End Select
        
    Exit Sub

End Sub

Private Sub BotaoUsuarios_Click()

Dim objUsuarios As New ClassUsuarios
Dim colSelecao As Collection

On Error GoTo Erro_BotaoUsuarios_Click

    'Guarda o Codigo do Usuario
    objUsuarios.sCodUsuario = CodUsuario.Caption

    'Chama a tela UsuarioLista
    Call Chama_Tela("UsuarioLista", colSelecao, objUsuarios, objEventoUsuario)

    Exit Sub

Erro_BotaoUsuarios_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154500)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_Codigo_Validate

    'Verifica se esta preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    objComprador.iCodigo = Codigo.Text
    objComprador.iFilialEmpresa = giFilialEmpresa

    'Seleciona no bd o comprador com o codigo informado
    lErro = CF("Comprador_Le", objComprador)
    If lErro <> SUCESSO And lErro <> 50064 Then Error 50067

    'Se existir, traz para a tela
    If lErro = SUCESSO Then
        lErro = Traz_Comprador_Tela(objComprador)
        If lErro <> SUCESSO Then Error 50068
    End If

    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 50066, 50067, 50068, 50140, 50144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154501)

    End Select

    Exit Sub

End Sub

Private Sub Compradores_DblClick()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Compradores_Click

    'Coloca o nome do Usuario selecionado no objUsuarios
    objUsuarios.sNomeReduzido = Compradores.List(Compradores.ListIndex)

    'Le o Usuario
    lErro = CF("Usuarios_Le_NomeRed", objUsuarios)
    If lErro <> SUCESSO And lErro <> 50132 Then Error 50133
    If lErro <> SUCESSO Then Error 53206

    'Preenche o objComprador com o Usuario lido
    objComprador.sCodUsuario = objUsuarios.sCodUsuario
    objComprador.iFilialEmpresa = giFilialEmpresa

    'Verifica o Usuario na tabela de Comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO Then Error 50128

    'Traz o comprador para tela
    lErro = Traz_Comprador_Tela(objComprador)
    If lErro <> SUCESSO Then Error 50065

    Exit Sub

Erro_Compradores_Click:

    Select Case Err

        Case 50065, 50128, 50133

        Case 53206
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", Err, objUsuarios.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154502)

    End Select

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

On Error GoTo Erro_Compradores_Form_Load

    iAlterado = 0

    Set objEventoUsuario = New AdmEvento
    
    'CArrega a listbox de compradores
    lErro = Compradores_Carrega()
    If lErro <> SUCESSO Then Error 50052

    AumentaQuant.Value = gobjCOM.iCompradorAumentaQuant
    
    Email.MaxLength = STRING_EMAIL

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Compradores_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 50052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154503)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_Tela_Extrai

    sTabela = "Comprador"

    'Armazena os dados presentes na tela em objComprador
    lErro = Move_Tela_Memoria(objComprador)
    If lErro <> SUCESSO Then Error 50053

    'Preenche a colecao de campos-valores com os dados de objComprador
    objComprador.sCodUsuario = CodUsuario.Caption

    colCampoValor.Add "CodUsuario", objComprador.sCodUsuario, STRING_USUARIO_CODIGO, "CodUsuario"
    colCampoValor.Add "Codigo", objComprador.iCodigo, 0, "Codigo"
    colCampoValor.Add "AumentaQuant", objComprador.iAumentaQuant, 0, "AumentaQuant"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 50053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154504)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim objComprador As New ClassComprador
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objComprador
    objComprador.sCodUsuario = colCampoValor.Item("CodUsuario").vValor
    objComprador.iCodigo = colCampoValor.Item("Codigo").vValor
    objComprador.iAumentaQuant = colCampoValor.Item("AumentaQuant").vValor

    If objComprador.iCodigo <> 0 Then
    
        'Verifica o Usuario na tabela de Comprador
        lErro = CF("Comprador_Le", objComprador)
        If lErro <> SUCESSO And lErro <> 50064 Then Error 50128

        'Se o Codigo do Comprador nao for nulo Traz o Comprador para a tela
        lErro = Traz_Comprador_Tela(objComprador)
        If lErro <> SUCESSO Then Error 50054

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 50054, 50137

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154505)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoUsuario = Nothing
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub objEventoUsuario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As ClassUsuarios
Dim iCodigo As Integer

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    Call Limpa_Tela(Me)

    Set objUsuarios = obj1

    objComprador.sCodUsuario = objUsuarios.sCodUsuario
    objComprador.iFilialEmpresa = giFilialEmpresa
    
    'Coloca os dados do usuário na tela
    CodUsuario.Caption = objUsuarios.sCodUsuario
    Nome.Caption = objUsuarios.sNome
    
    'Ler comprador correspondente ao usuario
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then Error 50055

    If lErro = SUCESSO Then
        Codigo.Text = objComprador.iCodigo
        AumentaQuant.Value = objComprador.iAumentaQuant
        iAlterado = 0
    End If


    'Fecha o comando de setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_ObjEventoUsuario_evSelecao:

    Select Case Err

        Case 50055, 50139

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154506)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Comprador"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Comprador"
    
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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

End Sub

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


Private Sub Nome_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nome, Source, X, Y)
End Sub

Private Sub Nome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nome, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub CodUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodUsuario, Source, X, Y)
End Sub

Private Sub CodUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodUsuario, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Email_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
