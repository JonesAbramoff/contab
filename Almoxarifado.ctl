VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Almoxarifado 
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   9195
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   795
      Visible         =   0   'False
      Width           =   8850
      Begin TelasEst.TabEndereco TabEnd 
         Height          =   3255
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   390
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   5741
      End
      Begin VB.Frame SSFrame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   135
         TabIndex        =   14
         Top             =   -120
         Width           =   8535
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            Left            =   210
            TabIndex        =   21
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label AlmoxarifadoLabel 
            Height          =   210
            Left            =   1500
            TabIndex        =   22
            Top             =   195
            Width           =   6765
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3330
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   990
      Width           =   8850
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2295
         Picture         =   "Almoxarifado.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   150
         Width           =   300
      End
      Begin VB.ListBox AlmoxarifadoList 
         Height          =   2400
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   2940
      End
      Begin VB.CommandButton BotaoEstoque 
         Caption         =   "Estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6120
         TabIndex        =   7
         Top             =   2850
         Width           =   2355
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   135
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1740
         TabIndex        =   3
         Top             =   810
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1740
         TabIndex        =   4
         Top             =   1485
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Top             =   2190
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label13 
         Caption         =   "Almoxarifados"
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
         Left            =   5850
         TabIndex        =   19
         Top             =   45
         Width           =   2040
      End
      Begin VB.Label ContaContabilLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta Contábil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         ToolTipText     =   "Conta contábil de estoque"
         Top             =   2235
         Width           =   1320
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
         Left            =   240
         TabIndex        =   18
         Top             =   1530
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
         Index           =   1
         Left            =   1125
         TabIndex        =   17
         Top             =   840
         Width           =   555
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
         Left            =   1005
         TabIndex        =   16
         Top             =   165
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6855
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Almoxarifado.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Almoxarifado.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Almoxarifado.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Almoxarifado.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4185
      Left            =   105
      TabIndex        =   15
      Top             =   315
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7382
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
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
Attribute VB_Name = "Almoxarifado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

Dim iFrameAtual As Integer
Public iAlterado As Integer
Public gobjTabEnd As New ClassTabEndereco
Dim objEndFilEMp As New ClassEndereco

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Endereco = 2

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo voucher(documento) disponível
    lErro = CF("Almoxarifado_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57510
    
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57510
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142680)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEndereco As New ClassEndereco
Dim colSelecao As New Collection
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_BotaoEstoque_Click

    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then Error 22253

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then Error 22254

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 22255

    'Preenche os objetos com os dados da tela
    lErro = Move_Tela_Memoria(objAlmoxarifado, objEndereco)
    If lErro <> SUCESSO Then Error 22256

    'Adiciona em ColSelecao
    colSelecao.Add objAlmoxarifado.iCodigo

    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    'Chama tela Estoque
    Call Chama_Tela("EstoqueAlmoxarifadoLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case Err

        Case 22253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 22254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)

        Case 22255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 22256

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142681)

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

On Error GoTo Erro_Codigo_Validate

    'Verifica se o Código foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é do tipo inteiro positivo
    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 22228

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 22228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142682)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 43572

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 43573

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 43574

        'Conta não cadastrada
        If lErro = 5700 Then Error 43575

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case Err

        Case 43572, 43574
    
        Case 43573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            
        Case 43575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142683)
    
    End Select

    Exit Sub

End Sub

Private Sub ContaContabilLabel_Click()
'Chama o browser de plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 43570

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

        Case 43570
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142684)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colAlmoxarifados As New Collection
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim sMascaraConta As String
Dim objFilialEmpresa As New AdmFiliais
Dim colEnderecos As New Collection
Dim objTela As Object

On Error GoTo Erro_Form_Load
    
    Set objEventoNumero = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoContaContabil = New AdmEvento

    iFrameAtual = 1

    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 28591

    ContaContabil.Mask = sMascaraConta

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 22229

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        AlmoxarifadoList.AddItem objAlmoxarifado.sNomeReduzido
        AlmoxarifadoList.ItemData(AlmoxarifadoList.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then Error 22230

    'Preenche objFilialEmpresa
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    'Lê a Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 22321
    
    'Se não encontrou a Filial Empresa --> Erro
    If lErro <> SUCESSO Then Error 22328
    
    'Se encontrou traz o Endereço p/ tela
    colEnderecos.Add objFilialEmpresa.objEndereco
    Set objEndFilEMp = objFilialEmpresa.objEndereco
            
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 22230
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22229, 22230, 28591
        
        Case 22321
        
        Case 22328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142685)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objAlmoxarifado As ClassAlmoxarifado) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um Almoxarifado selecionado
    If Not (objAlmoxarifado Is Nothing) Then

        'Verifica se o Almoxarifado existe, lendo no BD a partir do Codigo
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 22236

        'Se o Almoxarifado existe
        If lErro = SUCESSO Then
           
           lErro = Traz_Almoxarifado_Tela(objAlmoxarifado)
           If lErro <> SUCESSO Then Error 28554

        'Se o Almoxarifado não existe
        Else

            Call Limpa_Tela(Me)

            If objAlmoxarifado.iCodigo > 0 Then

                'Mantém o Código do Almoxarifado na tela
                Codigo.Text = CStr(objAlmoxarifado.iCodigo)

            End If

        End If

    End If

    'Zerar iAlterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 22236, 28554

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142686)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Traz_Almoxarifado_Tela(objAlmoxarifado As ClassAlmoxarifado) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer
Dim iCodigo As Integer
Dim sContaMascarada As String
Dim objEndereco As New ClassEndereco
Dim colEnderecos As New Collection

On Error GoTo Erro_Traz_Almoxarifado_Tela

    'Carrega lEndereco em objAlmoxarifado
    objEndereco.lCodigo = objAlmoxarifado.lEndereco

    'Lê o endereço à partir do Código
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 22241

    'Se não encomtrou o Endereço --> Erro
    If lErro <> SUCESSO Then Error 22242

    'Exibe os dados de objAlmoxarifado na tela
    If objAlmoxarifado.iCodigo = 0 Then
        Codigo.Text = ""
    Else
        Codigo.Text = CStr(objAlmoxarifado.iCodigo)
    End If
    Nome.Text = objAlmoxarifado.sDescricao
    NomeReduzido.Text = objAlmoxarifado.sNomeReduzido
    AlmoxarifadoLabel.Caption = Trim(NomeReduzido.Text)

    If objAlmoxarifado.sContaContabil <> "" Then
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objAlmoxarifado.sContaContabil, sContaMascarada)
        If lErro <> SUCESSO Then Error 28592
    Else
        sContaMascarada = ""
    End If

    ContaContabil.PromptInclude = False
    ContaContabil.Text = sContaMascarada
    ContaContabil.PromptInclude = True

    Call ContaContabil_Validate(bSGECancelDummy)
    
    colEnderecos.Add objEndereco
    
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 22241

    'Zerar iAlterado
    iAlterado = 0

    Traz_Almoxarifado_Tela = SUCESSO

    Exit Function

Erro_Traz_Almoxarifado_Tela:

    Traz_Almoxarifado_Tela = Err

    Select Case Err

        Case 22241, 28592

        Case 22242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142687)

    End Select

    Exit Function

End Function


Private Sub Label1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label13_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57822

    End If
    
    AlmoxarifadoLabel.Caption = Trim(NomeReduzido.Text)
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142689)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabil.Text = ""
    Else
        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 43571

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 43571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142690)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
            
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_ALMOXARIFADO_ID
                
            Case TAB_Endereco
                Parent.HelpContextID = IDH_ALMOXARIFADO_ENDERECO
                        
        End Select
    
    End If

End Sub

Private Sub AlmoxarifadoList_DblClick()

Dim lErro As Long
Dim sListBoxItem As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoList_DblClick

    'Guarda o valor do código do Almoxarifado selecionada na ListBox AlmoxarifadoList
    objAlmoxarifado.iCodigo = AlmoxarifadoList.ItemData(AlmoxarifadoList.ListIndex)

    'Lê o Almoxarifado no BD
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then Error 22248

    'Se Almoxarifado não está cadastrado, erro
    If lErro = 25056 Then Error 22249

    'Exibe os dados do Almoxarifado
    lErro = Traz_Almoxarifado_Tela(objAlmoxarifado)
    If lErro <> SUCESSO Then Error 22250

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_AlmoxarifadoList_DblClick:

    Select Case Err

        Case 22248, 22250

        Case 22249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
            AlmoxarifadoList.RemoveItem (AlmoxarifadoList.ListIndex)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142692)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Almoxarifado
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22251

    'Limpa a Tela
    Call Limpa_Tela_Almoxarifado

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 22251, 22252

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142693)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then Error 22253

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then Error 22254

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 22255
    
    'Preenche os objetos com os dados da tela
    lErro = Move_Tela_Memoria(objAlmoxarifado, objEndereco)
    If lErro <> SUCESSO Then Error 22256

    lErro = Trata_Alteracao(objAlmoxarifado, objAlmoxarifado.iCodigo)
    If lErro <> SUCESSO Then Error 32305

    'Grava o Almoxarifado no BD
    lErro = CF("Almoxarifado_Grava", objAlmoxarifado, objEndereco)
    If lErro <> SUCESSO Then Error 22257

    'Atualiza ListBox de Almoxarifado
    Call AlmoxarifadoList_Remove(objAlmoxarifado)
    Call AlmoxarifadoList_Adiciona(objAlmoxarifado)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 22254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)

        Case 22255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 22256, 22257, 32305

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142694)

    End Select

    Exit Function

End Function

Private Sub AlmoxarifadoList_Remove(objAlmoxarifado As ClassAlmoxarifado)
'Percorre a ListBox Almoxarifadolist para remover o Almoxarifado caso ela exista

Dim iIndice As Integer

    For iIndice = 0 To AlmoxarifadoList.ListCount - 1
    
        If AlmoxarifadoList.ItemData(iIndice) = objAlmoxarifado.iCodigo Then
    
            AlmoxarifadoList.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub
Private Sub AlmoxarifadoList_Adiciona(objAlmoxarifado As ClassAlmoxarifado)
'Inclui Almoxarifado na List

    AlmoxarifadoList.AddItem objAlmoxarifado.sNomeReduzido
    AlmoxarifadoList.ItemData(AlmoxarifadoList.NewIndex) = objAlmoxarifado.iCodigo

End Sub

Private Function Move_Tela_Memoria(objAlmoxarifado As ClassAlmoxarifado, objEndereco As ClassEndereco) As Long
'Lê os dados que estão na tela Almoxarifado e coloca em objAlmoxarifado

Dim lErro As Long
Dim iPais As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Move_Tela_Memoria

    'IDENTIFICACAO :
    If Len(Trim(Codigo.Text)) > 0 Then objAlmoxarifado.iCodigo = CInt(Codigo.Text)

    objAlmoxarifado.sDescricao = Trim(Nome.Text)
    objAlmoxarifado.sNomeReduzido = Trim(NomeReduzido.Text)
    objAlmoxarifado.iFilialEmpresa = giFilialEmpresa

    'Verifica se a Conta Contábil foi informada
    If Len(Trim(ContaContabil.ClipText)) > 0 Then
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 28553
        
        objAlmoxarifado.sContaContabil = sContaFormatada
    End If

    'ENDEREÇO
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
    If lErro <> SUCESSO Then Error 28553
    
    Set objEndereco = colEnderecos.Item(1)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 28553

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142695)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 22271

    objAlmoxarifado.iCodigo = CInt(Codigo.Text)
    
    'Lê os dados do Almoxarifado a ser excluido
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then Error 22272

    'Verifica se Almoxarifado está cadastrado
    If lErro = 25056 Then Error 22273

    'Envia aviso perguntando se realmente deseja excluir Almoxarifado
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_ALMOXARIFADO", objAlmoxarifado.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Almoxarifado
        lErro = CF("Almoxarifado_Exclui", objAlmoxarifado)
        If lErro <> SUCESSO Then Error 22274

        'Exclui da ListBox
        Call AlmoxarifadoList_Remove(objAlmoxarifado)

        'Limpa a Tela
        Call Limpa_Tela_Almoxarifado

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
            
        Case 22272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)

        Case 22275, 22274

        Case 22273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142696)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22312

    'Limpa a Tela
    Call Limpa_Tela_Almoxarifado

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 22312

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142697)

    End Select

End Sub

Sub Limpa_Tela_Almoxarifado()

Dim iIndice As Integer
Dim lErro As Long
Dim colEnderecos As New Collection

On Error GoTo Erro_Limpa_Tela_Almoxarifado

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa o Label Almoxarifado
    AlmoxarifadoLabel.Caption = ""

    'Desseleciona ListBox de Almoxarifado
    AlmoxarifadoList.ListIndex = -1

    Codigo.Text = ""
    
    Call gobjTabEnd.Limpa_Tela
    colEnderecos.Add objEndFilEMp
    Call gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    
    iAlterado = 0
    
    Exit Sub

Erro_Limpa_Tela_Almoxarifado:
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142698)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoContaContabil = Nothing
    Set objEndFilEMp = Nothing
    
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Private Sub AlmoxarifadoLabel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Almoxarifado"

    'Le os dados da Tela Almoxarifado
    lErro = Move_Tela_Memoria(objAlmoxarifado, objEndereco)
    If lErro <> SUCESSO Then Error 22314

    'No lEndereco armazena  0
    objAlmoxarifado.lEndereco = 0

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objAlmoxarifado.iCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objAlmoxarifado.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Descricao", objAlmoxarifado.sDescricao, STRING_ALMOXARIFADO_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objAlmoxarifado.sNomeReduzido, STRING_ALMOXARIFADO_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Endereco", objAlmoxarifado.lEndereco, 0, "Endereco"
    colCampoValor.Add "ContaContabil", objAlmoxarifado.sContaContabil, STRING_CONTA, "ContaContabil"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objAlmoxarifado.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 22314

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142699)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Tela_Preenche

    objAlmoxarifado.iCodigo = colCampoValor.Item("Codigo").vValor

    If objAlmoxarifado.iCodigo > 0 Then

        'Carrega objAlmoxarifado com os dados passados em colCampoValor
        objAlmoxarifado.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objAlmoxarifado.sDescricao = colCampoValor.Item("Descricao").vValor
        objAlmoxarifado.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objAlmoxarifado.lEndereco = colCampoValor.Item("Endereco").vValor
        objAlmoxarifado.sContaContabil = colCampoValor.Item("ContaContabil").vValor

        'Traz dados do Almoxarifado para a Tela
        lErro = Traz_Almoxarifado_Tela(objAlmoxarifado)
        If lErro <> SUCESSO Then Error 22315

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 22315

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142700)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALMOXARIFADO_ID
    Set Form_Load_Ocx = Me
    Caption = "Almoxarifado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Almoxarifado"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        End If
    End If

End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub Label56_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label56, Source, X, Y)
End Sub

Private Sub Label56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label56, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoLabel, Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub
