VERSION 5.00
Begin VB.UserControl UsuProdArtlux 
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   5880
   Begin VB.Frame Frame2 
      Caption         =   "Acesso as Etapas"
      Height          =   705
      Left            =   135
      TabIndex        =   14
      Top             =   1275
      Width           =   3390
      Begin VB.CheckBox AcessoMontagem 
         Caption         =   "Montagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         TabIndex        =   2
         Top             =   315
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox AcessoForro 
         Caption         =   "Forro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   315
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CheckBox AcessoCorte 
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   0
         Top             =   300
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3540
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "UsuProdArtlux.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "UsuProdArtlux.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "UsuProdArtlux.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "UsuProdArtlux.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Usuarios 
      Height          =   2010
      Left            =   3570
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   975
      Width           =   2100
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
      Picture         =   "UsuProdArtlux.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2145
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   690
      Left            =   135
      TabIndex        =   9
      Top             =   570
      Width           =   3390
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
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   660
      End
      Begin VB.Label CodUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   855
         TabIndex        =   11
         Top             =   255
         Width           =   1080
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Usuários da Produção"
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
      Left            =   3570
      TabIndex        =   13
      Top             =   690
      Width           =   1875
   End
End
Attribute VB_Name = "UsuProdArtlux"
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

Private Function Move_Tela_Memoria(ByVal objUsu As ClassUsuProdArtlux) As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se o codigo não estiver vazio coloca-o no objComprador
    objUsu.sCodUsuario = CodUsuario.Caption
    objUsu.iFilialEmpresa = giFilialEmpresa
    
    If AcessoCorte.Value = vbChecked Then
        objUsu.iAcessoCorte = MARCADO
    Else
        objUsu.iAcessoCorte = DESMARCADO
    End If
    
    If AcessoForro.Value = vbChecked Then
        objUsu.iAcessoForro = MARCADO
    Else
        objUsu.iAcessoForro = DESMARCADO
    End If
    
    If AcessoMontagem.Value = vbChecked Then
        objUsu.iAcessoMontagem = MARCADO
    Else
        objUsu.iAcessoMontagem = DESMARCADO
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206679)

    End Select

    Exit Function

End Function

Private Sub Adiciona_Lista_Usuario(ByVal objUsu As ClassUsuProdArtlux)
'Adiciona um comprador na ListBox

Dim lErro As Long
Dim colComprador As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Adiciona_Lista_Usuario
    
    If Len(Trim(objUsu.sCodUsuario)) > 0 Then
        'Se ele é novo adiciona-o na lista
        Usuarios.AddItem objUsu.sCodUsuario
    End If
    
    Exit Sub

Erro_Adiciona_Lista_Usuario:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206680)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava um registro no bd

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos estão preenchidos
    If Len(CodUsuario.Caption) = 0 Then gError 206681
                                                                
    lErro = Move_Tela_Memoria(objUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Trata_Alteracao(objUsu, objUsu.iFilialEmpresa, objUsu.sCodUsuario)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava
    lErro = CF("UsuProdArtlux_Grava", objUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Adiciona na listbox se necessário
    lErro = Usuarios_Carrega()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 206681
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206682)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub Limpa_Tela_Usuario()
'Limpa a tela

    Call Limpa_Tela(Me)

    CodUsuario.Caption = ""
    AcessoCorte.Value = vbUnchecked
    AcessoForro.Value = vbUnchecked
    AcessoMontagem.Value = vbUnchecked
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

End Sub

Function Traz_Usuario_Tela(ByVal objUsu As ClassUsuProdArtlux) As Long
'Traz os dados do comprador para a tela

Dim iIndice As Integer
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Traz_Usuario_Tela

    'Le o CodUsuario
    objUsuarios.sCodUsuario = objUsu.sCodUsuario

    'Le o Usuario na tabela
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
    If lErro <> SUCESSO Then gError 206683
    
    objUsu.iFilialEmpresa = giFilialEmpresa

    lErro = CF("UsuProdArtlux_Le", objUsu)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    'Preenche a tela com os dados de objComprador
    CodUsuario.Caption = objUsu.sCodUsuario

    If objUsu.iAcessoCorte = MARCADO Then
        AcessoCorte.Value = vbChecked
    Else
        AcessoCorte.Value = vbUnchecked
    End If
    
    If objUsu.iAcessoForro = MARCADO Then
        AcessoForro.Value = vbChecked
    Else
        AcessoForro.Value = vbUnchecked
    End If
    
    If objUsu.iAcessoMontagem = MARCADO Then
        AcessoMontagem.Value = vbChecked
    Else
        AcessoMontagem.Value = vbUnchecked
    End If
    
    iAlterado = 0

    Traz_Usuario_Tela = SUCESSO

    Exit Function

Erro_Traz_Usuario_Tela:

    Traz_Usuario_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206683
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206684)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional ByVal objUsu As ClassUsuProdArtlux) As Long
'Trata os parametros

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Comprador preenchido
    If Not (objUsu Is Nothing) Then
                
        lErro = Traz_Usuario_Tela(objUsu)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206685)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Usuarios_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colUsu As New Collection

On Error GoTo Erro_Usuarios_Carrega

    Usuarios.Clear

    'Le todos os Usuarios da Colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Le todos os Compradores da Filial Empresa
    lErro = CF("UsuProdArtlux_Le_Todos", colUsu)
    If lErro <> SUCESSO And lErro <> 50126 Then gError ERRO_SEM_MENSAGEM

    For Each objUsu In colUsu
        For Each objUsuarios In colUsuarios
            If objUsu.sCodUsuario = objUsuarios.sCodUsuario Then
                Usuarios.AddItem objUsu.sCodUsuario
            End If
        Next
    Next

    Usuarios_Carrega = SUCESSO

    Exit Function

Erro_Usuarios_Carrega:

    Usuarios_Carrega = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206686)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o comprador está preenchido
    If Len(CodUsuario.Caption) = 0 Then gError 206687

    objUsu.iFilialEmpresa = giFilialEmpresa
    objUsu.sCodUsuario = CodUsuario.Caption

    'Pede a confirmação da exclusão do comprador do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_USUPRODARTLUX")
    If vbMsgRes = vbYes Then
    
        lErro = CF("UsuProdArtlux_Exclui", objUsu)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Call Limpa_Tela_Usuario
   
        'Procura o índice da comprador
        For iIndice = 0 To Usuarios.ListCount - 1
            If Usuarios.List(iIndice) = objUsu.sCodUsuario Then
                Usuarios.RemoveItem iIndice
                Exit For
            End If
        Next

        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 206687
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206682)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Usuario

    iAlterado = 0

    Exit Sub
Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206683)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Usuario

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206684)

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

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206685)

    End Select

    Exit Sub

End Sub

Private Sub Usuarios_DblClick()

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux

On Error GoTo Erro_Usuarios_DblClick

    'Coloca o nome do Usuario selecionado no objUsuarios
    objUsu.sCodUsuario = Usuarios.List(Usuarios.ListIndex)

    'Traz o comprador para tela
    lErro = Traz_Usuario_Tela(objUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Usuarios_DblClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206686)

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

On Error GoTo Erro_Form_Load

    iAlterado = 0

    Set objEventoUsuario = New AdmEvento
    
    'CArrega a listbox de compradores
    lErro = Usuarios_Carrega()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206687)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    sTabela = "UsuProdArtlux"

    colCampoValor.Add "CodUsuario", CodUsuario.Caption, STRING_USUARIO_CODIGO, "CodUsuario"

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206688)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objComprador
    objUsu.sCodUsuario = colCampoValor.Item("CodUsuario").vValor

    'Se o Codigo do Comprador nao for nulo Traz o Comprador para a tela
    lErro = Traz_Usuario_Tela(objUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206689)

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
Dim objUsuarios As ClassUsuarios

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    Call Limpa_Tela(Me)

    Set objUsuarios = obj1

    'Coloca os dados do usuário na tela
    CodUsuario.Caption = objUsuarios.sCodUsuario
    
    'Fecha o comando de setas, se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_ObjEventoUsuario_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206690)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Usuário da Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "UsuProdArtlux"
    
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

Private Sub CodUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodUsuario, Source, X, Y)
End Sub

Private Sub CodUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodUsuario, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


