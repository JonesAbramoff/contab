VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Gerente 
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   8640
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6360
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Gerente.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Gerente.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Gerente.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "Gerente.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Gerentes 
      Height          =   1815
      Left            =   6090
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1245
      Width           =   2415
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
      Height          =   405
      Left            =   1635
      TabIndex        =   1
      Top             =   150
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1395
      Left            =   90
      TabIndex        =   10
      Top             =   765
      Width           =   5595
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
         TabIndex        =   14
         Top             =   360
         Width           =   660
      End
      Begin VB.Label CodUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   13
         Top             =   360
         Width           =   1080
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
         TabIndex        =   12
         Top             =   870
         Width           =   555
      End
      Begin VB.Label Nome 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1245
         TabIndex        =   11
         Top             =   840
         Width           =   3900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gerente"
      Height          =   885
      Left            =   120
      TabIndex        =   0
      Top             =   2325
      Width           =   5595
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1860
         Picture         =   "Gerente.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Numeração Automática"
         Top             =   345
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   315
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
         TabIndex        =   9
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Gerentes"
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
      Index           =   0
      Left            =   6105
      TabIndex        =   16
      Top             =   960
      Width           =   780
   End
End
Attribute VB_Name = "Gerente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declarações Globais
Dim iAlterado As Integer
Private WithEvents objEventoUsuario As AdmEvento
Attribute objEventoUsuario.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Gerentes_Form_Load

    iAlterado = 0

    Set objEventoUsuario = New AdmEvento

    'Carrega a listbox com os gerentes da Filial Empresa.
    lErro = Gerentes_Carrega()
    If lErro <> SUCESSO Then gError 81072

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Gerentes_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 81072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161592)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Gerentes_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objGerente As ClassGerente
Dim colUsuarios As New Collection
Dim objUsuarios As ClassUsuarios
Dim colGerente As New Collection

On Error GoTo Erro_Gerentes_Carrega

    'Le todos os Usuarios da Colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 81083

    'Le todos os Gerentes da Filial Empresa
    lErro = CF("Gerente_Le_Todos", colGerente)
    If lErro <> SUCESSO And lErro <> 81076 Then gError 81078

    For Each objGerente In colGerente
        
        For Each objUsuarios In colUsuarios
            
            If objGerente.sCodUsuario = objUsuarios.sCodUsuario Then
               Gerentes.AddItem objUsuarios.sNomeReduzido
            End If
            
        Next
        
    Next
        
    Gerentes_Carrega = SUCESSO

    Exit Function

Erro_Gerentes_Carrega:

    Gerentes_Carrega = gErr

    Select Case gErr

        Case 81078, 81083

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161593)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objGerente As ClassGerente) As Long
'Trata os parametros

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um gerente preenchido
    If Not (objGerente Is Nothing) Then

        'Se objGerente.iCodigo > 0
        If objGerente.iCodigo > 0 Then

            'Verifica se o Gerente existe, lendo no BD a partir do código
            lErro = CF("Gerente_Le", objGerente)
            If lErro <> SUCESSO And lErro <> 81026 Then gError 81188
          
            'Se o Gerente existe
            If lErro = SUCESSO Then
                lErro = Traz_Gerente_Tela(objGerente)
                If lErro <> SUCESSO Then gError 81189
                
                'Se o gerente não existe
            Else

                'Mantém o Código do gerente na tela
                Codigo.Text = CStr(objGerente.iCodigo)

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 81188, 81189
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161594)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_Gerente_Tela(objGerente As ClassGerente) As Long
'Traz os dados do Gerente para a tela

Dim iIndice As Integer
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Traz_Gerente_Tela

    'Le o CodUsuario
    objUsuarios.sCodUsuario = objGerente.sCodUsuario

    'Le o Usuario na tabela
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 81081
    If lErro <> SUCESSO Then gError 81082

    'Preenche a tela com os dados de objGerente
    CodUsuario.Caption = objGerente.sCodUsuario
    Nome.Caption = objUsuarios.sNome

    Codigo.Text = objGerente.iCodigo

    iAlterado = 0

    Traz_Gerente_Tela = SUCESSO

    Exit Function

Erro_Traz_Gerente_Tela:

    Traz_Gerente_Tela = gErr

    Select Case gErr

        Case 81081

        Case 81082
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161595)

    End Select

    Exit Function
    
End Function
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objGerente As New ClassGerente

On Error GoTo Erro_Tela_Extrai

    sTabela = "Gerente"

    'Armazena os dados presentes na tela em objGerente
    lErro = Move_Tela_Memoria(objGerente)
    If lErro <> SUCESSO Then gError 81079

    'Preenche a colecao de campos-valores com os dados de objGerente
    objGerente.sCodUsuario = CodUsuario.Caption

    colCampoValor.Add "CodUsuario", objGerente.sCodUsuario, STRING_USUARIO_CODIGO, "CodUsuario"
    colCampoValor.Add "Codigo", objGerente.iCodigo, 0, "Codigo"
   

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 81079

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161596)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objGerente As ClassGerente) As Long
'Lê os dados que estão na tela Gerente e coloca em objGerente

On Error GoTo Erro_Move_Tela_Memoria

    'Se o codigo não estiver vazio coloca-o no objGerente
     objGerente.iCodigo = StrParaInt(Codigo.ClipText)
     objGerente.sCodUsuario = CodUsuario.Caption
    
    objGerente.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161597)

    End Select

    Exit Function

End Function

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim objGerente As New ClassGerente
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objGerente
    objGerente.sCodUsuario = colCampoValor.Item("CodUsuario").vValor
    objGerente.iCodigo = colCampoValor.Item("Codigo").vValor
  
    If objGerente.iCodigo <> 0 Then

        'Se o Codigo do Gerente nao for nulo Traz o Gerente para a tela
        lErro = Traz_Gerente_Tela(objGerente)
        If lErro <> SUCESSO Then gError 81080

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 81080

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161598)

    End Select

    Exit Sub

End Sub
Private Sub BotaoUsuarios_Click()

Dim objUsuarios As New ClassUsuarios
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_BotaoUsuarios_Click

    'Guarda o Codigo do Usuario
    objUsuarios.sCodUsuario = CodUsuario.Caption

    'Chama a tela UsuarioLista
    Call Chama_Tela("UsuarioLista", colSelecao, objUsuarios, objEventoUsuario)

    Exit Sub

Erro_BotaoUsuarios_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161599)

    End Select

    Exit Sub

End Sub
Private Sub objEventoUsuario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objGerente As New ClassGerente
Dim objUsuarios As ClassUsuarios
Dim iCodigo As Integer

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    Call Limpa_Tela(Me)

    Set objUsuarios = obj1

    objGerente.sCodUsuario = objUsuarios.sCodUsuario
    objGerente.iFilialEmpresa = giFilialEmpresa

    'Ler Gerente correspondente ao usuario
    lErro = CF("Gerente_Le_Usuario", objGerente)
    If lErro <> SUCESSO And lErro <> 81084 Then gError 81085

    If lErro = SUCESSO Then
    
        lErro = Traz_Gerente_Tela(objGerente)
        If lErro <> SUCESSO Then gError 81143
        
    End If

    'Coloca os dados do usuário na tela
    CodUsuario.Caption = objUsuarios.sCodUsuario
    Nome.Caption = objUsuarios.sNome

    'Fecha o comando de setas, se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_ObjEventoUsuario_evSelecao:

    Select Case gErr

        Case 81085, 81143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161600)

    End Select

    Exit Sub

End Sub
Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código da proximo gerente
    lErro = Gerente_Automatico(iCodigo)
    If lErro <> SUCESSO Then gError 81099

    Codigo.Text = iCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 81099

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161601)

    End Select

    Exit Sub

End Sub
Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objGerente As New ClassGerente

On Error GoTo Erro_BotaoGravar_Click

    'Grava os registros na tabela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 81101

    Call Limpa_Tela_Gerente

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 81101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161602)

    End Select

    Exit Sub

End Sub
Public Function Gravar_Registro() As Long
'Grava um registro no bd

Dim lErro As Long
Dim objGerente As New ClassGerente
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos estão preenchidos
    If Len(CodUsuario.Caption) = 0 Then gError 81102
    If Len(Nome.Caption) = 0 Then gError 81103
    If Len(Codigo.ClipText) = 0 Then gError 81104


    'Transfere os dados da tela para os obj's
    objGerente.sCodUsuario = CodUsuario.Caption
    objGerente.iCodigo = StrParaInt(Codigo.ClipText)
    objGerente.iFilialEmpresa = giFilialEmpresa

    'Le o Gerente com o usuario da tela
    lErro = CF("Gerente_Le_Usuario", objGerente)
    If lErro <> SUCESSO And lErro <> 81084 Then gError 81105

    'Se encontrar
    If lErro = SUCESSO Then

        'Verifica se o codigo e o mesmo que o codigo da tela
        If (objGerente.iCodigo <> StrParaInt(Codigo.ClipText)) Then gError 81106

    End If

    'Chama função que armazena os dados da tela no objGerente
    lErro = Move_Tela_Memoria(objGerente)
    If lErro <> SUCESSO Then gError 81107

    lErro = Trata_Alteracao(objGerente, objGerente.iFilialEmpresa, objGerente.iCodigo)
    If lErro <> SUCESSO Then gError 32330

    'Grava a Gerente no BD
    lErro = CF("Gerente_Grava", objGerente)
    If lErro <> SUCESSO Then gError 81108
    

    'Adiciona na listbox se necessário
     Call Adiciona_Lista_Gerente(objGerente)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32330, 81105, 81107, 81108

        Case 81194
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)
                    
        Case 81102, 81103
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)
     
        Case 81104
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 81106
            Call Rotina_Erro(vbOKOnly, "ERRO_GERENTE_USUARIO", gErr, objGerente.iCodigo, objGerente.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161603)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function
Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objGerente As New ClassGerente
Dim objGerente1 As ClassGerente
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Gerente está preenchido
    If Len(Codigo.Text) = 0 Then gError 81120

    objGerente.iCodigo = CInt(Codigo.Text)

    'Verifica se o usuário tem Gerente
    lErro = CF("Gerente_Le", objGerente)
    If lErro <> SUCESSO And lErro <> 81026 Then gError 81121
    If lErro <> SUCESSO Then gError 81122

    'Pede a confirmação da exclusão do Gerente do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_GERENTE", objGerente.sCodUsuario)
    If vbMsgRes = vbYes Then

        lErro = Move_Tela_Memoria(objGerente)
        If lErro <> SUCESSO Then gError 81123

        lErro = CF("Gerente_Exclui", objGerente)
        If lErro <> SUCESSO Then gError 81124

        Call Limpa_Tela_Gerente

        objUsuarios.sCodUsuario = objGerente.sCodUsuario
        
          'Lê o nome do usuário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO Then gError 81192

        'Procura o índice do Gerente
        For iIndice = 0 To Gerentes.ListCount - 1
            If Gerentes.List(iIndice) = objUsuarios.sNomeReduzido Then
                Gerentes.RemoveItem iIndice
                Exit For
            End If
        Next

        iAlterado = 0

    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 81120
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 81121, 81123, 81124, 81192

        Case 81122
            Call Rotina_Erro(vbOKOnly, "ERRO_GERENTE_NAO_CADASTRADO", gErr, objGerente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161604)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Dim lErro As Long
    
    Set objEventoUsuario = Nothing
    
    lErro = ComandoSeta_Liberar(Me.Name)
    
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Gerente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Gerente"
    
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


Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 81138
    
    Call Limpa_Tela_Gerente

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 81138

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161605)

    End Select

    Exit Sub

End Sub


Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
Dim lErro As Long
Dim objGerente As New ClassGerente

On Error GoTo Erro_Codigo_Validate

    'Verifica se esta preenchido
    If Len(Codigo.ClipText) = 0 Then Exit Sub

    objGerente.iCodigo = StrParaInt(Codigo.Text)
    
    objGerente.iFilialEmpresa = giFilialEmpresa

    'Seleciona no bd o Gerente com o codigo informado
    lErro = CF("Gerente_Le", objGerente)
    If lErro <> SUCESSO And lErro <> 81096 Then gError 81097
    'Se existir, traz para a tela
    If lErro = SUCESSO Then
        lErro = Traz_Gerente_Tela(objGerente)
        If lErro <> SUCESSO Then gError 81098
    End If

    'Fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 81097, 81098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161606)

    End Select

    Exit Sub

End Sub

Private Sub CodUsuario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Gerentes_DblClick()

Dim lErro As Long
Dim objGerente As New ClassGerente
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gerentes_DblClick

    'Coloca o nome do Usuario selecionado no objUsuarios
    objUsuarios.sNomeReduzido = Gerentes.List(Gerentes.ListIndex)

    'Le o Usuario
    lErro = CF("Usuarios_Le_NomeRed", objUsuarios)
    If lErro <> SUCESSO And lErro <> 50132 Then gError 81089
    If lErro <> SUCESSO Then gError 81090

    'Preenche o objGerente com o Usuario lido
    objGerente.sCodUsuario = objUsuarios.sCodUsuario
    objGerente.iFilialEmpresa = giFilialEmpresa

    'Verifica o Usuario na tabela de Gerente
    lErro = CF("Gerente_Le_Usuario", objGerente)
    If lErro <> SUCESSO Then gError 81091

    'Traz o gerente para tela
    lErro = Traz_Gerente_Tela(objGerente)
    If lErro <> SUCESSO Then gError 81092

    Exit Sub

Erro_Gerentes_DblClick:

    Select Case gErr

        Case 81089, 81091, 81092

        Case 81090
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuarios.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161607)

    End Select

    Exit Sub

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

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
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Habilita o reconhecimento das teclas F2 e F3

    Select Case KeyCode
    
        'Se o usuário pressiona a tecla F2 => dispara o botão próximo número
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
    
           
    End Select
    
End Sub
Function Gerente_Automatico(iCodigo As Integer) As Long
'Gera o próximo gerente

Dim lErro As Long

On Error GoTo Erro_Gerente_Automatico

    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "COD_PROX_GERENTE", "Gerente", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 81100

    Gerente_Automatico = SUCESSO

    Exit Function

Erro_Gerente_Automatico:

    Gerente_Automatico = gErr

    Select Case gErr

        Case 81100

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161608)
    
    End Select

    Exit Function

End Function
Private Sub Adiciona_Lista_Gerente(objGerente As ClassGerente)
'Adiciona um gerente na ListBox

Dim lErro As Long

On Error GoTo Erro_Adiciona_Lista_Gerente

    'verifica se o nome reduzido foi preenchido
    If Len(Trim(objGerente.sNomeReduzido)) > 0 Then Gerentes.AddItem objGerente.sNomeReduzido
    'Se ele é novo adiciona-o na lista
   
    Exit Sub

Erro_Adiciona_Lista_Gerente:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161609)

    End Select

    Exit Sub

End Sub
Private Sub Limpa_Tela_Gerente()
'Limpa a tela

Dim lErro As Long

    Call Limpa_Tela(Me)

    Nome.Caption = ""
    CodUsuario.Caption = ""
    Codigo.Text = ""
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

End Sub
Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub
Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1
    
End Sub
