VERSION 5.00
Begin VB.UserControl OperadorLogin 
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6150
   Begin VB.Frame FrameOperador 
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         Picture         =   "OperadorLogin.ctx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton BotaoCancelar 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   3720
         Picture         =   "OperadorLogin.ctx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1935
         Width           =   1815
      End
      Begin VB.CommandButton BotaoOK 
         Default         =   -1  'True
         Height          =   615
         Left            =   1575
         Picture         =   "OperadorLogin.ctx":32D4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1935
         Width           =   1815
      End
      Begin VB.ComboBox Operadores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   525
         Width           =   3960
      End
      Begin VB.TextBox Senha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1575
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   3945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label LabelNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   450
         TabIndex        =   5
         Top             =   570
         Width           =   780
      End
   End
End
Attribute VB_Name = "OperadorLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Constantes Relacionadas a Tela de Operador
Const LOGIN_APENAS_GERENTE = 1
Const LOGIN_OPERADOR_SUSPENSO = 2
Const LOGIN_TODOS_OPERADORES = 3
Const OPERADOR_GERENTE = 1

'Declaração de um objGlobal do Tipo ClassOperador
Dim gobjOperador As New ClassOperador

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Login"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OperadorLogin"

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

'***** fim do trecho a ser copiado ******

'Inicio Tela de OperadorLogin dia 11/07/02 Sergio Ricardo
Public Function Trata_Parametros(Optional objOperador As ClassOperador, Optional iTipoCarregamento As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Fazendo com que o objGlobal aponte para a mesma posição de memória do obj passado como parâmetro
    Set gobjOperador = objOperador

    'Função que Vai Carregar a Combo com todos os Operadores
    lErro = Carrega_Operador_Combo(iTipoCarregamento)
    If lErro <> SUCESSO Then gError 107534

    'Se LOGIN_OPERADOR_SUSPENSO desabilitar botao cancel, só pode retornar ao sistema com a Senha
    If iTipoCarregamento = LOGIN_OPERADOR_SUSPENSO Then BotaoCancelar.Enabled = False

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 107534
                'Erro Tratado Dentro da Função que foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163686)

    End Select

    Exit Function

End Function

Function Carrega_Operador_Combo(iTipoCarregamento) As Long
'Função que Vai carregar a Combo com os Operadores em Questão selecionado pelo tipo passado como parâmetro

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_Carrega_Operador_Combo

    If iTipoCarregamento = LOGIN_OPERADOR_SUSPENSO And giCodOperador = 0 Then iTipoCarregamento = LOGIN_TODOS_OPERADORES

    For Each objOperador In gcolOperadores
        'Apenas Gerentes podem ser operadores
        If iTipoCarregamento = LOGIN_APENAS_GERENTE Then

            If objOperador.iGerente = OPERADOR_GERENTE Then

                Operadores.AddItem objOperador.sNome
                Operadores.ItemData(Operadores.NewIndex) = objOperador.iCodigo

            End If
        'Carrega Todos os Gerentes eo operador suspenso
        ElseIf iTipoCarregamento = LOGIN_OPERADOR_SUSPENSO Then

            If objOperador.iGerente = OPERADOR_GERENTE Or objOperador.iCodigo = giCodOperador Then

                Operadores.AddItem objOperador.sNome
                Operadores.ItemData(Operadores.NewIndex) = objOperador.iCodigo

            End If

        'Carrga todos os operadores sem distinção
        ElseIf iTipoCarregamento = LOGIN_TODOS_OPERADORES Then

            Operadores.AddItem objOperador.sNome
            Operadores.ItemData(Operadores.NewIndex) = objOperador.iCodigo

        End If

    Next

    Carrega_Operador_Combo = SUCESSO

    Exit Function

Erro_Carrega_Operador_Combo:

    Carrega_Operador_Combo = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163687)

    End Select

    Exit Function

End Function

Private Sub BotaoOk_Click()
'Função que Verifica se o Operador está Selecionado, e Verifica se senha está Certa

Dim lErro As Long
Dim objOperador As New ClassOperador
Dim iCodigo As Integer

On Error GoTo Erro_BotaoOk_Click

    'Verifica se Algum operador foi selecionado
    If Operadores.ListIndex = -1 Then gError 107535

    'Guardar na Variável iCodigo o Código do Operador Selecionado
    iCodigo = Operadores.ItemData(Operadores.ListIndex)

    'Verifica se a Senha está foi digitada
    If Len(Trim(Senha.Text)) = 0 Then gError 107536

    'Verificar dentro da Coleção se a Senha Confere
    For Each objOperador In gcolOperadores

        If objOperador.iCodigo = iCodigo Then

            'Se a Senha não for igual então sai por erro
            If Not (objOperador.sSenha = Senha.Text) Then
                gError 107537

            Else
                'Carregar o objOperador Globlal com as propriedades da Coleção Global
                gobjOperador.iCodigo = objOperador.iCodigo
                gobjOperador.iCodigoVendedor = objOperador.iCodigoVendedor
                gobjOperador.iDesconto = objOperador.iDesconto
                gobjOperador.iFilialEmpresa = objOperador.iFilialEmpresa
                gobjOperador.iGerente = objOperador.iGerente
                gobjOperador.iLimiteDesconto = objOperador.iLimiteDesconto
                gobjOperador.sNome = objOperador.sNome
                gobjOperador.sSenha = objOperador.sSenha

                'se a Senha for encontrada, e o obj Global for carregado então sai do for
                Exit For

            End If

        End If

    Next

    'Retorno da Tela
    giRetornoTela = vbOK

    'Fechar a Tela
    Unload Me

    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 107535
            Call Rotina_ErroECF(vbOKOnly, ERRO_OPERADOR_NAO_SELECIONADO, gErr)

        Case 107536
            Call Rotina_ErroECF(vbOKOnly, ERRO_SENHA_NAO_PREENCHIDA1, gErr)

        Case 107537
            Call Rotina_ErroECF(vbOKOnly, ERRO_SENHA_INVALIDA1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163688)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCancelar_Click()

    'Se a Saida for pelo Botão Cancela o retorno de tela Será Cancel
    giRetornoTela = vbCancel

    'Fecha a Tela
    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'Verifica se a Saida da Tela foi pelo X se foi Erro
 
Dim lErro As Long

On Error GoTo Erro_Form_QueryUnload

    If UnloadMode = vbFormControlMenu Then
        
        Cancel = 1
        gError 107720
    
    End If
    
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
    Exit Sub
    
Erro_Form_QueryUnload:

    Select Case gErr

        Case 107720

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163689)

    End Select

    Exit Sub

End Sub




