VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NOME_LOGICO_FISICO 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7320
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "NOME_LOGICO_FISICO.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "NOME_LOGICO_FISICO.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "NOME_LOGICO_FISICO.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "NOME_LOGICO_FISICO.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5670
      Left            =   75
      TabIndex        =   5
      Top             =   240
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10001
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NOME_TAB_1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NOME_TAB_2"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Height          =   5220
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   660
      Width           =   9195
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   2000
         TabIndex        =   6
         Top             =   850
         Width           =   880
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   2000
         TabIndex        =   8
         Top             =   1300
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   2000
         TabIndex        =   10
         Top             =   1750
         Width           =   5500
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   315
         Left            =   2000
         TabIndex        =   12
         Top             =   2200
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   2000
         TabIndex        =   14
         Top             =   2650
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorProduto 
         Height          =   315
         Left            =   2000
         TabIndex        =   16
         Top             =   3100
         Width           =   880
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigo 
         Caption         =   "Codigo:"
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
         Height          =   315
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   875
         Width           =   1500
      End
      Begin VB.Label LabelNomeReduzido 
         Caption         =   "NomeReduzido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   9
         Top             =   1325
         Width           =   1500
      End
      Begin VB.Label LabelDescricao 
         Caption         =   "Descricao:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   11
         Top             =   1775
         Width           =   1500
      End
      Begin VB.Label LabelCcl 
         Caption         =   "Ccl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   13
         Top             =   2225
         Width           =   1500
      End
      Begin VB.Label LabelContaContabil 
         Caption         =   "ContaContabil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   15
         Top             =   2675
         Width           =   1500
      End
      Begin VB.Label LabelValorProduto 
         Caption         =   "ValorProduto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   17
         Top             =   3125
         Width           =   1500
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Height          =   5220
      Index           =   2
      Left            =   135
      TabIndex        =   28
      Top             =   660
      Width           =   9195
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   2000
         TabIndex        =   18
         Top             =   3550
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   3310
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3550
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Hora 
         Height          =   315
         Left            =   2000
         TabIndex        =   21
         Top             =   4000
         Width           =   880
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   2000
         TabIndex        =   23
         Top             =   4450
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   2000
         MaxLength       =   255
         TabIndex        =   25
         Top             =   4900
         Width           =   5500
      End
      Begin VB.Label LabelData 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   20
         Top             =   3575
         Width           =   1500
      End
      Begin VB.Label LabelHora 
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   22
         Top             =   4025
         Width           =   1500
      End
      Begin VB.Label LabelProduto 
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   24
         Top             =   4475
         Width           =   1500
      End
      Begin VB.Label LabelObservacao 
         Caption         =   "Observacao:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   26
         Top             =   4925
         Width           =   1500
      End
   End
End
Attribute VB_Name = "NOME_LOGICO_FISICO"
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
    Caption = "Teste"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Teste"

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

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodigo = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163585)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163586)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTeste As ClassTeste) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTeste Is Nothing) Then

        lErro = Traz_Teste_Tela(objTeste)
        If lErro <> SUCESSO Then gError 100024

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 100024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163587)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTeste As ClassTeste) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTeste.lCodigo = StrParaLong(Codigo.Text)
    objTeste.iFilialEmpresa = giFilialEmpresa
    objTeste.sNomeReduzido = NomeReduzido.Text
    objTeste.sDescricao = Descricao.Text
    objTeste.sCcl = Ccl.Text
    objTeste.sContaContabil = ContaContabil.Text
    objTeste.dValorProduto = StrParaDbl(ValorProduto.Text)
    If Len(Trim(Data.ClipText)) <> 0 Then objTeste.dtData = Format(Data.Text, Data.Format)
    objTeste.dHora = StrParaDbl(HORA.Text)
    objTeste.sProduto = Produto.Text
    objTeste.sObservacao = Observacao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163588)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTeste As New ClassTeste

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Teste"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTeste)
    If lErro <> SUCESSO Then gError 100025

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTeste.lCodigo, 0, "Codigo"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 100025

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163589)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTeste As New ClassTeste

On Error GoTo Erro_Tela_Preenche

    objTeste.lCodigo = colCampoValor.Item("Codigo").vValor

    objTeste.iFilialEmpresa = giFilialEmpresa

    If objTeste.lCodigo <> 0 And objTeste.iFilialEmpresa <> 0 Then
        lErro = Traz_Teste_Tela(objTeste)
        If lErro <> SUCESSO Then gError 100026
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 100026

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163590)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTeste As New ClassTeste

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 100027
    '#####################

    'Preenche o objTeste
    lErro = Move_Tela_Memoria(objTeste)
    If lErro <> SUCESSO Then gError 100028

    lErro = Trata_Alteracao(objTeste, , objTeste.lCodigo, objTeste.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 100029


    'Grava o/a Teste no Banco de Dados
    lErro = CF("Teste_Grava", objTeste)
    If lErro <> SUCESSO Then gError 100030

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 100027
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODIGO_TESTE_NAO_PREENCHIDO">, gErr)
            Codigo.SetFocus

        Case 100028, 100029, 100030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163591)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Teste() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Teste
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Teste = SUCESSO

    Exit Function

Erro_Limpa_Tela_Teste:

    Limpa_Tela_Teste = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163592)

    End Select

    Exit Function

End Function

Function Traz_Teste_Tela(objTeste As ClassTeste) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Teste_Tela

    'Lê o Teste que está sendo Passado
    lErro = CF("Teste_Le", objTeste)
    If lErro <> SUCESSO And lErro <> 100004 Then gError 100031

    If lErro = SUCESSO Then

        If objTeste.lCodigo <> 0 Then Codigo.Text = CStr(objTeste.lCodigo)
        NomeReduzido.Text = objTeste.sNomeReduzido
        Descricao.Text = objTeste.sDescricao
        Ccl.Text = objTeste.sCcl
        ContaContabil.Text = objTeste.sContaContabil
        If objTeste.dValorProduto <> 0 Then ValorProduto.Text = CStr(objTeste.dValorProduto)

        If objTeste.dtData <> 0 Then
            Data.PromptInclude = False
            Data.Text = Format(objTeste.dtData, "dd/mm/yy")
            Data.PromptInclude = True
        End If

        If objTeste.dHora <> 0 Then HORA.Text = CStr(objTeste.dHora)
        Produto.Text = objTeste.sProduto
        Observacao.Text = objTeste.sObservacao

    End If

    Traz_Teste_Tela = SUCESSO

    Exit Function

Erro_Traz_Teste_Tela:

    Traz_Teste_Tela = gErr

    Select Case gErr

        Case 100031

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163593)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 100032

    'Limpa Tela
    Call Limpa_Tela_Teste

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 100032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163594)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163595)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 100033

    Call Limpa_Tela_Teste

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 100033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163596)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTeste As New ClassTeste
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 100034
    '#####################

    objTeste.lCodigo = StrParaLong(Codigo.Text)
    objTeste.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TESTE", objTeste.lCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a requisição de consumo
    lErro = CF("Teste_Exclui", objTeste)
    If lErro <> SUCESSO Then gError 100035

    'Limpa Tela
    Call Limpa_Tela_Teste

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 100034
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODIGO_TESTE_NAO_PREENCHIDO">, gErr)
            Codigo.SetFocus

        Case 100034
        Case 100035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163597)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Veifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 100036

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 100036

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163598)

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

        Case 100037

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163599)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Veifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case 100038

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163600)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Ccl_Validate

    'Veifica se Ccl está preenchida
    If Len(Trim(Ccl.Text)) <> 0 Then

       '#######################################
       'CRITICA Ccl
       '#######################################

    End If

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    Select Case gErr

        Case 100039

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163601)

    End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaContabil_Validate

    'Veifica se ContaContabil está preenchida
    If Len(Trim(ContaContabil.Text)) <> 0 Then

       '#######################################
       'CRITICA ContaContabil
       '#######################################

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case gErr

        Case 100040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163602)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorProduto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorProduto_Validate

    'Veifica se ValorProduto está preenchida
    If Len(Trim(ValorProduto.Text)) <> 0 Then

       'Critica a ValorProduto
       lErro = Valor_Positivo_Critica(ValorProduto.Text)
       If lErro <> SUCESSO Then gError 100041

    End If

    Exit Sub

Erro_ValorProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 100041

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163603)

    End Select

    Exit Sub

End Sub

Private Sub ValorProduto_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorProduto, iAlterado)
    
End Sub

Private Sub ValorProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 100042

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 100042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163604)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 100043

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 100043

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163605)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 100044

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 100044

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163606)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Veifica se Hora está preenchida
    If Len(Trim(HORA.Text)) <> 0 Then

       'Critica a Hora
       lErro = Valor_Positivo_Critica(HORA.Text)
       If lErro <> SUCESSO Then gError 100045

    End If

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 100045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163607)

    End Select

    Exit Sub

End Sub

Private Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HORA, iAlterado)
    
End Sub

Private Sub Hora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate

    'Veifica se Produto está preenchida
    If Len(Trim(Produto.Text)) <> 0 Then

       '#######################################
       'CRITICA Produto
       '#######################################

    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 100046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163608)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

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

        Case 100047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163609)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTeste As ClassTeste

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTeste = obj1

    'Mostra os dados do Teste na tela
    lErro = Traz_Teste_Tela(objTeste)
    If lErro <> SUCESSO Then gError 100048

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 100048


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163610)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTeste As New ClassTeste
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTeste.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("TesteLista", colSelecao, objTeste, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163611)

    End Select

    Exit Sub

End Sub
