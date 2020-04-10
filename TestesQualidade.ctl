VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl TestesQualidade 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   ScaleHeight     =   4425
   ScaleWidth      =   8835
   Begin VB.TextBox NomeReduzido 
      Height          =   315
      Left            =   1785
      MaxLength       =   100
      TabIndex        =   2
      Top             =   660
      Width           =   4485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Padrões para o cadastramento dos produtos"
      Height          =   3090
      Left            =   195
      TabIndex        =   17
      Top             =   1185
      Width           =   8490
      Begin VB.TextBox MetodoUsado 
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   3
         Top             =   300
         Width           =   3285
      End
      Begin VB.CheckBox NoCertificado 
         Caption         =   "O resultado deste teste deve aparecer no certificado"
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
         Left            =   180
         TabIndex        =   9
         Top             =   2685
         Value           =   1  'Checked
         Width           =   4890
      End
      Begin VB.ComboBox TipoResultado 
         Height          =   315
         ItemData        =   "TestesQualidade.ctx":0000
         Left            =   6180
         List            =   "TestesQualidade.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   1830
      End
      Begin VB.Frame Frame5 
         Caption         =   "Limites"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   165
         TabIndex        =   20
         Top             =   765
         Width           =   4125
         Begin MSMask.MaskEdBox LimiteDe 
            Height          =   315
            Left            =   825
            TabIndex        =   5
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LimiteAte 
            Height          =   315
            Left            =   2730
            TabIndex        =   6
            Top             =   255
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            Caption         =   "De:"
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
            Left            =   375
            TabIndex        =   22
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Até:"
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
            Left            =   2250
            TabIndex        =   21
            Top             =   315
            Width           =   375
         End
      End
      Begin VB.TextBox Observacao 
         Height          =   780
         Left            =   4410
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1770
         Width           =   3855
      End
      Begin VB.TextBox Especificacao 
         Height          =   780
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1770
         Width           =   4035
      End
      Begin VB.Label Label4 
         Caption         =   "Método:"
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
         Left            =   225
         TabIndex        =   24
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Resultado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4455
         TabIndex        =   23
         Top             =   330
         Width           =   1650
      End
      Begin VB.Label Label10 
         Caption         =   "Observação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4410
         TabIndex        =   19
         Top             =   1545
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Especificação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   1515
         Width           =   1305
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2325
      Picture         =   "TestesQualidade.ctx":001F
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   195
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6555
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TestesQualidade.ctx":0109
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TestesQualidade.ctx":0263
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TestesQualidade.ctx":03ED
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TestesQualidade.ctx":091F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   180
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   1035
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   210
      Width           =   660
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
      Index           =   0
      Left            =   270
      TabIndex        =   15
      Top             =   705
      Width           =   1410
   End
End
Attribute VB_Name = "TestesQualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Testes de Qualidade"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TestesQualidade"

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
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174653)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174654)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTestesQualidade As ClassTestesQualidade) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTestesQualidade Is Nothing) Then

        lErro = Traz_TestesQualidade_Tela(objTestesQualidade)
        If lErro <> SUCESSO Then gError 130128

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130128

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174655)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTestesQualidade As ClassTestesQualidade) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTestesQualidade.iCodigo = StrParaInt(Codigo.Text)
    objTestesQualidade.sNomeReduzido = NomeReduzido.Text
    objTestesQualidade.sEspecificacao = Especificacao.Text
    objTestesQualidade.iTipoResultado = TipoResultado.ListIndex
    objTestesQualidade.dLimiteDe = StrParaDbl(LimiteDe.Text)
    objTestesQualidade.dLimiteAte = StrParaDbl(LimiteAte.Text)
    objTestesQualidade.sMetodoUsado = MetodoUsado.Text
    objTestesQualidade.sObservacao = Observacao.Text
    objTestesQualidade.iNoCertificado = NoCertificado.Value

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174656)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TestesQualidade"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTestesQualidade)
    If lErro <> SUCESSO Then gError 130129

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTestesQualidade.iCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130129

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174657)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade

On Error GoTo Erro_Tela_Preenche

    objTestesQualidade.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTestesQualidade.iCodigo <> 0 Then
        lErro = Traz_TestesQualidade_Tela(objTestesQualidade)
        If lErro <> SUCESSO Then gError 130130
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130130

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174658)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 130131
    '#####################

    'Preenche o objTestesQualidade
    lErro = Move_Tela_Memoria(objTestesQualidade)
    If lErro <> SUCESSO Then gError 130132

    lErro = Trata_Alteracao(objTestesQualidade, objTestesQualidade.iCodigo)
    If lErro <> SUCESSO Then gError 130258

    'Grava o/a TestesQualidade no Banco de Dados
    lErro = CF("TestesQualidade_Grava", objTestesQualidade)
    If lErro <> SUCESSO Then gError 130133

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130131
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTESQUALIDADE_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 130132, 130133, 130258

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174659)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TestesQualidade() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TestesQualidade
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_TestesQualidade = SUCESSO

    Exit Function

Erro_Limpa_Tela_TestesQualidade:

    Limpa_Tela_TestesQualidade = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174660)

    End Select

    Exit Function

End Function

Function Traz_TestesQualidade_Tela(objTestesQualidade As ClassTestesQualidade) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TestesQualidade_Tela

    'Lê o TestesQualidade que está sendo Passado
    lErro = CF("TestesQualidade_Le", objTestesQualidade)
    If lErro <> SUCESSO And lErro <> 130109 Then gError 130134

    If lErro = SUCESSO Then

        If objTestesQualidade.iCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objTestesQualidade.iCodigo)
            Codigo.PromptInclude = True
        Else
            Codigo.Text = ""
        End If
        NomeReduzido.Text = objTestesQualidade.sNomeReduzido
        Especificacao.Text = objTestesQualidade.sEspecificacao
        TipoResultado.ListIndex = objTestesQualidade.iTipoResultado
        If objTestesQualidade.dLimiteDe <> 0 Then
            LimiteDe.Text = CStr(objTestesQualidade.dLimiteDe)
        Else
            LimiteDe.Text = ""
        End If
        If objTestesQualidade.dLimiteAte <> 0 Then
            LimiteAte.Text = CStr(objTestesQualidade.dLimiteAte)
        Else
            LimiteAte.Text = ""
        End If
        MetodoUsado.Text = objTestesQualidade.sMetodoUsado
        Observacao.Text = objTestesQualidade.sObservacao
        NoCertificado.Value = objTestesQualidade.iNoCertificado

    End If

    Traz_TestesQualidade_Tela = SUCESSO

    Exit Function

Erro_Traz_TestesQualidade_Tela:

    Traz_TestesQualidade_Tela = gErr

    Select Case gErr

        Case 130134

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174661)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130135

    'Limpa Tela
    Call Limpa_Tela_TestesQualidade

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174662)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174663)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130136

    Call Limpa_Tela_TestesQualidade

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130136

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174664)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 130137
    '#####################

    objTestesQualidade.iCodigo = StrParaInt(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TESTESQUALIDADE", objTestesQualidade.iCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a requisição de consumo
    lErro = CF("TestesQualidade_Exclui", objTestesQualidade)
    If lErro <> SUCESSO Then gError 130138

    'Limpa Tela
    Call Limpa_Tela_TestesQualidade

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130137
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTESQUALIDADE_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 130137
        Case 130138

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174665)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Inteiro_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 130139

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 130139

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174666)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Veifica se NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

       '#######################################
       'CRITICA NomeReduzido
       '#######################################

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr

        Case 130141

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174667)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Especificacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Especificacao_Validate

    'Veifica se Especificacao está preenchida
    If Len(Trim(Especificacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Especificacao
       '#######################################

    End If

    Exit Sub

Erro_Especificacao_Validate:

    Cancel = True

    Select Case gErr

        Case 130142

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174668)

    End Select

    Exit Sub

End Sub

Private Sub Especificacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoResultado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TipoResultado_Validate

    'Verifica se TipoResultado está preenchida
    If Len(Trim(TipoResultado.Text)) <> 0 Then

'       'Critica a TipoResultado
'       lErro = Inteiro_Critica(TipoResultado.Text)
'       If lErro <> SUCESSO Then gError 130143

    End If

    Exit Sub

Erro_TipoResultado_Validate:

    Cancel = True

    Select Case gErr

        Case 130143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174669)

    End Select

    Exit Sub

End Sub

Private Sub TipoResultado_GotFocus()
    
'    Call MaskEdBox_TrataGotFocus(TipoResultado, iAlterado)
    
End Sub

Private Sub TipoResultado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteDe_Validate

    'Veifica se LimiteDe está preenchida
    If Len(Trim(LimiteDe.Text)) <> 0 Then

       'Critica a LimiteDe
       lErro = Valor_Double_Critica(LimiteDe.Text)
       If lErro <> SUCESSO Then gError 130144

    End If

    Exit Sub

Erro_LimiteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 130144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174670)

    End Select

    Exit Sub

End Sub

Private Sub LimiteDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LimiteDe, iAlterado)
    
End Sub

Private Sub LimiteDe_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteAte_Validate

    'Veifica se LimiteAte está preenchida
    If Len(Trim(LimiteAte.Text)) <> 0 Then

       'Critica a LimiteAte
       lErro = Valor_Double_Critica(LimiteAte.Text)
       If lErro <> SUCESSO Then gError 130145

    End If

    Exit Sub

Erro_LimiteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 130145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174671)

    End Select

    Exit Sub

End Sub

Private Sub LimiteAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LimiteAte, iAlterado)
    
End Sub

Private Sub LimiteAte_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MetodoUsado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MetodoUsado_Validate

    'Veifica se MetodoUsado está preenchida
    If Len(Trim(MetodoUsado.Text)) <> 0 Then

       '#######################################
       'CRITICA MetodoUsado
       '#######################################

    End If

    Exit Sub

Erro_MetodoUsado_Validate:

    Cancel = True

    Select Case gErr

        Case 130146

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174672)

    End Select

    Exit Sub

End Sub

Private Sub MetodoUsado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Veifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case 130147

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174673)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NoCertificado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NoCertificado_Validate

    Exit Sub

Erro_NoCertificado_Validate:

    Cancel = True

    Select Case gErr

        Case 130148

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174674)

    End Select

    Exit Sub

End Sub

Private Sub NoCertificado_GotFocus()
    
'    Call MaskEdBox_TrataGotFocus(NoCertificado, iAlterado)
    
End Sub

Private Sub NoCertificado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTestesQualidade As ClassTestesQualidade

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTestesQualidade = obj1

    'Mostra os dados do TestesQualidade na tela
    lErro = Traz_TestesQualidade_Tela(objTestesQualidade)
    If lErro <> SUCESSO Then gError 130149

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 130149


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174675)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTestesQualidade As New ClassTestesQualidade
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTestesQualidade.iCodigo = Codigo.Text

    End If

    Call Chama_Tela("TestesQualidadeLista", colSelecao, objTestesQualidade, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174676)

    End Select

    Exit Sub

End Sub

'##################################
'Inserido por Wagner
Private Sub BotaoProxNum_Click()
'Numeração automática

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("TestesQualidade_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 138301

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 138301
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174677)
    
    End Select

    Exit Sub
    
End Sub
'##################################

