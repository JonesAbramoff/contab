VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TipoFornecedorOcx 
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   8490
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3390
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   8025
      Begin MSComctlLib.TreeView Contas 
         Height          =   2205
         Left            =   5415
         TabIndex        =   11
         Top             =   930
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   3889
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.ListBox Historicos 
         Height          =   2205
         Left            =   5415
         TabIndex        =   12
         Top             =   930
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.ComboBox CondicaoPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Text            =   " "
         Top             =   1290
         Width           =   2280
      End
      Begin VB.Frame SSFrame6 
         Height          =   510
         Left            =   150
         TabIndex        =   18
         Top             =   75
         Width           =   7785
         Begin VB.Label Label11 
            Caption         =   "Tipo de Fornecedor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   105
            TabIndex        =   21
            Top             =   195
            Width           =   1740
         End
         Begin VB.Label Tipo 
            Caption         =   "tipo de fornecedor"
            Height          =   210
            Left            =   1920
            TabIndex        =   22
            Top             =   195
            Width           =   5700
         End
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   315
         Left            =   2535
         TabIndex        =   7
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Contabilidade - Estoque / Despesa"
         Height          =   1290
         Left            =   255
         TabIndex        =   19
         Top             =   1815
         Width           =   4575
         Begin MSMask.MaskEdBox Historico 
            Height          =   315
            Left            =   2235
            TabIndex        =   10
            Top             =   840
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaDespesa 
            Height          =   315
            Left            =   2235
            TabIndex        =   9
            Top             =   405
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelHistorico 
            AutoSize        =   -1  'True
            Caption         =   "Histórico Padrão:"
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
            Left            =   645
            TabIndex        =   23
            Top             =   885
            Width           =   1485
         End
         Begin VB.Label LabelContaContabil 
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
            Height          =   195
            Left            =   795
            TabIndex        =   24
            Top             =   450
            Width           =   1335
         End
      End
      Begin VB.Label LabelHistoricos 
         Caption         =   "Históricos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5445
         TabIndex        =   25
         Top             =   690
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label8 
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1545
         TabIndex        =   26
         Top             =   855
         Width           =   885
      End
      Begin VB.Label CondicaoPagtoLabel 
         Caption         =   "Condição de Pagamento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   1350
         Width           =   2160
      End
      Begin VB.Label LabelContas 
         Caption         =   "Plano de Contas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5445
         TabIndex        =   28
         Top             =   690
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3390
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   870
      Width           =   8025
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2085
         Picture         =   "TipoFornecedorOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   390
         Width           =   300
      End
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1545
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2205
         Width           =   3675
      End
      Begin VB.ListBox Tipos 
         Height          =   2595
         ItemData        =   "TipoFornecedorOcx.ctx":00EA
         Left            =   5670
         List            =   "TipoFornecedorOcx.ctx":00EC
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   495
         Width           =   2310
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1545
         TabIndex        =   3
         Top             =   1245
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1545
         TabIndex        =   1
         Top             =   375
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "9999"
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
         Left            =   810
         TabIndex        =   29
         Top             =   405
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   540
         TabIndex        =   30
         Top             =   1290
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   375
         TabIndex        =   31
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Tipos de Fornecedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5655
         TabIndex        =   32
         Top             =   255
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6150
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoFornecedorOcx.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoFornecedorOcx.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoFornecedorOcx.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoFornecedorOcx.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   3810
      Left            =   120
      TabIndex        =   20
      Top             =   510
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6720
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Financeiros"
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
Attribute VB_Name = "TipoFornecedorOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Private WithEvents objEventoCondicaoPagto As AdmEvento
Attribute objEventoCondicaoPagto.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_DadosFinanceiros = 2

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("TipoFornecedor_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57557
    
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57557
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174852)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16322
    
    Call Limpa_Tela_TipoFornecedor
        
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 16322
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174853)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim iEncontrou As Integer
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16329

    objTipoFornecedor.iCodigo = CInt(Codigo.Text)

    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then Error 16331

    'Não achou o Tipo Fornecedor
    If lErro = 12765 Then Error 16330
    
    'Pedido de confirmacao de exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPODEFORNECEDOR", objTipoFornecedor.iCodigo)

    If vbMsgRes = vbYes Then

        'exclui Tipo de Fornecedor
        lErro = CF("TipoFornecedor_Exclui", objTipoFornecedor)
        If lErro <> SUCESSO Then Error 16332

        'procura indice do Tipo de Fornecedor na ListBox
        iEncontrou = 0
        For iIndice = 0 To Tipos.ListCount - 1
            
            If Tipos.ItemData(iIndice) = objTipoFornecedor.iCodigo Then
                iEncontrou = 1
                Exit For
            End If
            
        Next

        'remove Tipo de Fornecedor do ListBox
        If iEncontrou = 1 Then Tipos.RemoveItem (iIndice)

        Call Limpa_Tela_TipoFornecedor

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEFORNECEDOR_NAO_CADASTRADO", Err, objTipoFornecedor.iCodigo)
        
        Case 16331, 16332

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174854)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim sConta As String
Dim iContaPreenchida As Integer
Dim iPosicao As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Codigo.Text = "" Then Error 16324

    'verifica preenchimento da descricao
    If Len(Trim(Descricao.Text)) = 0 Then Error 16325
    
    lErro = Move_Tela_Memoria(objTipoFornecedor)
    If lErro <> SUCESSO Then Error 33951

    lErro = Trata_Alteracao(objTipoFornecedor, objTipoFornecedor.iCodigo)
    If lErro <> SUCESSO Then Error 16349

    lErro = CF("TipoFornecedor_Grava", objTipoFornecedor)
    If lErro <> SUCESSO Then Error 16328

    'Remove e adiciona na ListBox
    Call Tipos_Remove(objTipoFornecedor)
    Call Tipos_Adiciona(objTipoFornecedor)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 16328, 16349, 33951
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174855)

    End Select
    
    Exit Function
        
End Function

Function Move_Tela_Memoria(objTipoFornecedor As ClassTipoFornecedor) As Long
'Move os dados da tela para memória.

Dim lErro As Long
Dim sConta As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Preenche objTipoFornecedor
    If Len(Trim(Codigo.Text)) = 0 Then
        objTipoFornecedor.iCodigo = 0
    Else
        objTipoFornecedor.iCodigo = CInt(Codigo.Text)
    End If
    
    If Len(Trim(Desconto.Text)) = 0 Then
        objTipoFornecedor.dDesconto = 0
    Else
        objTipoFornecedor.dDesconto = CDbl(Desconto.Text) / 100
    End If
    
    If Len(Trim(Historico.Text)) = 0 Then
        objTipoFornecedor.iHistPadraoDespesa = 0
    Else
        objTipoFornecedor.iHistPadraoDespesa = CInt(Historico.Text)
    End If
       
    objTipoFornecedor.iCondicaoPagto = CondPagto_Extrai(CondicaoPagto)
    
    objTipoFornecedor.sObservacao = Observacao.Text
    objTipoFornecedor.sDescricao = Descricao.Text
    
    sConta = String(STRING_TIPO_FORNECEDOR_CONTA_DESPESA, 0)

    'Critica o formato da conta
    lErro = CF("Conta_Formata", ContaDespesa.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 16326

    'testa se a conta está preenchida
    If iContaPreenchida <> CONTA_PREENCHIDA Then
        objTipoFornecedor.sContaDespesa = ""
    Else
        objTipoFornecedor.sContaDespesa = sConta
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
    
        Case 16326

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174856)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80457

    Call Limpa_Tela_TipoFornecedor

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80457

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174857)

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

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then Error 16382
        
        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 16381
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, Codigo.Text)
    
        Case 16382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, Codigo.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174858)
    
    End Select

    Exit Sub
           
End Sub

Private Sub CondicaoPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagto_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CondicaoPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagto_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagto.Text = CondicaoPagto.List(CondicaoPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 33553

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 33554
        
        If lErro <> SUCESSO Then Error 33555 'Não encontrou CondicaoPagto no BD

        'Encontrou CondicaoPagto no BD e não é de Pagamento
        If objCondicaoPagto.iEmPagamento = 0 Then Error 33556

        'Coloca no Text da Combo
        CondicaoPagto.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 33557

    Exit Sub

Erro_CondicaoPagto_Validate:

    Cancel = True


    Select Case Err

        Case 33556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, iCodigo)
    
        Case 33553, 33554
    
        Case 33555  'Não encontrou CondicaoPagto no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
    
            Else
                'Segura o foco
    
            End If
    
        Case 33557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagto.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174859)

    End Select

    Exit Sub

End Sub

Private Sub ContaDespesa_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaDespesa_GotFocus()

    Contas.Visible = True
    LabelContas.Visible = True
    
    Historicos.Visible = False
    LabelHistoricos.Visible = False

End Sub

Private Sub ContaDespesa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_ContaDespesa_Validate

    If Len(Trim(ContaDespesa.ClipText)) > 0 Then
        
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", ContaDespesa.Text, ContaDespesa.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39799
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 39800
            
            ContaDespesa.PromptInclude = False
            ContaDespesa.Text = sContaMascarada
            ContaDespesa.PromptInclude = True
        
            'se não encontrou a conta simples
            ElseIf lErro = 44096 Or lErro = 44098 Then

                lErro = CF("Conta_Critica", ContaDespesa.Text, sContaFormatada, objPlanoConta, MODULO_CONTASAPAGAR)
                If lErro <> SUCESSO And lErro <> 5700 Then Error 16313
                
                If lErro = 5700 Then Error 47892
                
        End If
        
    End If
    
    Exit Sub
    
Erro_ContaDespesa_Validate:

    Cancel = True


    Select Case Err
    
        Case 16313, 39799
        
        Case 39800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
        
        Case 47892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaDespesa.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174860)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Contas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_Contas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta_Modulo1", objNode, Contas.Nodes, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO Then Error 47891
        
    End If
    
    Exit Sub
    
Erro_Contas_Expand:

    Select Case Err
    
        Case 47891
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174861)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub Contas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim sTipoConta As String
Dim sConta As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Contas_NodeClick

    sTipoConta = left(Node.Key, 1)
        
    If sTipoConta = "A" Then
    
        sConta = right(Node.Key, Len(Node.Key) - 1)
                
        lErro = CF("Conta_SelecionaUma", sConta, objPlanoConta, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 16375
    
        'Ausência de Conta no BD
        If lErro = 6030 Then Error 16315
        
        'Verfica se conta pode receber lançamentos
        If objPlanoConta.iAtivo <> CONTA_ATIVA Then Error 16376
                       
        'Exibe Conta Contábil
        sContaEnxuta = String(STRING_TIPO_FORNECEDOR_CONTA_DESPESA, 0)
    
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 16316
            
        'Exibe conta contábil
        ContaDespesa.PromptInclude = False
        ContaDespesa.Text = sContaEnxuta
        ContaDespesa.PromptInclude = True
        
    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub
    
Erro_Contas_NodeClick:

    Select Case Err
    
        Case 16315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
        
        Case 16316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
        
        Case 16375
        
        Case 16376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, sConta)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174862)
        
    End Select
    
    Exit Sub
        
End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Desconto_Validate

    'verifica se o campo foi preenchido
    If Len(Trim(Desconto.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 16306
        
    End If
        
    Exit Sub
    
Erro_Desconto_Validate:

    Cancel = True


    Select Case Err
    
        Case 16306
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174863)
        
    End Select
    
    Exit Sub
        
End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome
Dim colCodigo As New Collection
Dim sMascaraConta As String
 
On Error GoTo Erro_Form_Load

    Set objEventoCondicaoPagto = New AdmEvento
    
    'Inicializa variáveis globais
    iFrameAtual = 1
    iAlterado = 0
       
    Set colCodigoDescricao = New AdmColCodigoNome
    
    'Leitura dos codigos e descricoes dos tipos de fornecedores no BD
    lErro = CF("Cod_Nomes_Le", "TiposDeFornecedor", "Codigo", "Descricao", STRING_TIPO_FORNECEDOR_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16283

    'Preenche ListBox Tipos com descricao de TiposDeFornecedor
    For iIndice = 1 To colCodigoDescricao.Count
        Set objCodigoDescricao = colCodigoDescricao(iIndice)
        Tipos.AddItem objCodigoDescricao.sNome
        Tipos.ItemData(Tipos.NewIndex) = objCodigoDescricao.iCodigo
    Next
     
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagto, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 16284
     
'    Set colCodigoDescricao = New AdmColCodigoNome
'
'    'Lê cada código e descrição reduzida da tabela CondicoesPagto
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 16284
'
'    'preenche a ComboBox CondicaoPagto com os objetos da colecao colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'
'        CondicaoPagto.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
'        CondicaoPagto.ItemData(CondicaoPagto.NewIndex) = objCodigoDescricao.iCodigo
'
'    Next
'
    'Verifica se o modulo de contabilidade esta ativo antes das inicializacoes
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        Set colCodigoDescricao = New AdmColCodigoNome

        'Leitura dos codigos e descricoes da tabela HistPadrao
        lErro = CF("Cod_Nomes_Le", "HistPadrao", "HistPadrao", "DescHistPadrao", STRING_HISTPADRAO_DESCRICAO, colCodigoDescricao)
        If lErro <> SUCESSO Then Error 16285

        'preenche a ListBox com os codigos e descricoes da tabela HistPadrao
        For Each objCodigoDescricao In colCodigoDescricao
            Historicos.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
            Historicos.ItemData(Historicos.NewIndex) = objCodigoDescricao.iCodigo
        Next
    
        'Inicializa a Lista de Plano de Contas
        lErro = CF("Carga_Arvore_Conta_Modulo", Contas.Nodes, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO Then Error 16286
    
        'Inicializa mask de ContaDespesa
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 16287

        ContaDespesa.Mask = sMascaraConta
        
    Else
       
       'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 16287

        ContaDespesa.Mask = sMascaraConta
        
        LabelContaContabil.Enabled = False
        LabelHistorico.Enabled = False
        ContaDespesa.Enabled = False
        Historico.Enabled = False
        Historicos.Enabled = False
        LabelHistoricos.Enabled = False
        Contas.Enabled = False
        ContaDespesa.Enabled = False
        
    End If
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 16283, 16284, 16285, 16286, 16287
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174864)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoCondicaoPagto = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Function Trata_Parametros(Optional objTipoFornecedor As ClassTipoFornecedor) As Long

Dim lErro As Long
Dim sContaEnxuta As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se existir um Tipo de Fornecedor selecionado, exibir seus dados
    If Not (objTipoFornecedor Is Nothing) Then

        lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
        If lErro <> SUCESSO And lErro <> 12765 Then Error 16288

        If lErro = SUCESSO Then
        
            lErro = Exibe_Dados_TipoFornecedor(objTipoFornecedor)
            If lErro <> SUCESSO Then Error 12770
            
        Else
        
            'Limpa a tela
            Call Limpa_Tela_TipoFornecedor
            
            'Exibe apenas o codigo
             Codigo.Text = CStr(objTipoFornecedor.iCodigo)
             
        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 12770, 16288
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174865)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Exibe_Dados_TipoFornecedor(objTipoFornecedor As ClassTipoFornecedor) As Long
'Traz os dados para tela

Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_TipoFornecedor

    'IDENTIFICACAO :

    'Mostra dados na tela
    Codigo.Text = CStr(objTipoFornecedor.iCodigo)
    Descricao.Text = objTipoFornecedor.sDescricao
    
    If objTipoFornecedor.sObservacao <> "" Then
        Observacao.Text = objTipoFornecedor.sObservacao
    Else
        Observacao.Text = ""
    End If
    
    'DADOS FINANCEIROS :

    If objTipoFornecedor.dDesconto <> 0 Then
        Desconto.Text = CStr(objTipoFornecedor.dDesconto * 100)
    Else
        Desconto.Text = ""
    End If
    
    If objTipoFornecedor.sContaDespesa <> "" Then
        
        sContaEnxuta = String(STRING_TIPO_FORNECEDOR_CONTA_DESPESA, 0)

        lErro = Mascara_RetornaContaEnxuta(objTipoFornecedor.sContaDespesa, sContaEnxuta)
        If lErro <> SUCESSO Then Error 16314

        ContaDespesa.PromptInclude = False
        ContaDespesa.Text = sContaEnxuta
        ContaDespesa.PromptInclude = True
        
    Else
        ContaDespesa.PromptInclude = False
        ContaDespesa.Text = ""
        ContaDespesa.PromptInclude = True
    End If
    
    If objTipoFornecedor.iHistPadraoDespesa <> 0 Then
    
        Historico.PromptInclude = False
        Historico.Text = CStr(objTipoFornecedor.iHistPadraoDespesa)
        Historico.PromptInclude = True
    Else
        Historico.PromptInclude = False
        Historico.Text = ""
        Historico.PromptInclude = True
    End If
    
    If objTipoFornecedor.iCondicaoPagto <> 0 Then
        CondicaoPagto.Text = objTipoFornecedor.iCondicaoPagto
        Call CondicaoPagto_Validate(bSGECancelDummy)
    Else
        CondicaoPagto.Text = ""
    End If
    
    iAlterado = 0

    Exibe_Dados_TipoFornecedor = SUCESSO

Exit Function

Erro_Exibe_Dados_TipoFornecedor:

    Exibe_Dados_TipoFornecedor = Err

    Select Case Err

        Case 16314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTipoFornecedor.sContaDespesa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174866)

    End Select

    Exit Function

End Function

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Historico_GotFocus()

    Historicos.Visible = True
    LabelHistoricos.Visible = True
    
    Contas.Visible = False
    LabelContas.Visible = False

    Call MaskEdBox_TrataGotFocus(Historico, iAlterado)

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Historico_Validate

    If Len(Trim(Historico.Text)) > 0 Then
    
        objHistPadrao.iHistPadrao = CInt(Historico.Text)
    
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 16316
        
        If lErro <> SUCESSO Then Error 16317
        
    End If
    
    Exit Sub
    
Erro_Historico_Validate:

    Cancel = True


    Select Case Err
    
        Case 16316
                
        Case 16317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174867)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub Historicos_DblClick()

Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Historicos_DblClick

    objHistPadrao.iHistPadrao = Historicos.ItemData(Historicos.ListIndex)
    
    lErro = CF("HistPadrao_Le", objHistPadrao)
    If lErro <> SUCESSO And lErro <> 5446 Then Error 16320
        
    If lErro = 5446 Then Error 16321
        
    Historico.PromptInclude = False
    Historico.Text = CStr(objHistPadrao.iHistPadrao)
    Historico.PromptInclude = True
           
    Exit Sub

Erro_Historicos_DblClick:

    Select Case Err
    
        Case 16320
        
        Case 16321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174868)
            
    End Select
    
    Exit Sub
        
End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        If iFrameAtual = TAB_Identificacao Then Tipo.Caption = Descricao.Text
        
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_TIPOS_FORN_IDENT
                
            Case TAB_DadosFinanceiros
                Parent.HelpContextID = IDH_TIPOS_FORN_DADOS_FIN
                        
        End Select
        
    End If

End Sub

Private Sub Tipos_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_Tipos_DblClick

    objTipoFornecedor.iCodigo = Tipos.ItemData(Tipos.ListIndex)

    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then Error 16304
    
    'Tipo não cadastrado
    If lErro <> SUCESSO Then Error 16305

    lErro = Exibe_Dados_TipoFornecedor(objTipoFornecedor)
    If lErro <> SUCESSO Then Error 16373
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Tipos_DblClick:

    Select Case Err

        Case 16304, 16373

        Case 16305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEFORNECEDOR_NAO_CADASTRADO", Err, objTipoFornecedor.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174869)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_TipoFornecedor()
'Limpa todos os campos de input da tela TipoFornecedor

Dim lErro As Long
Dim iIndice As Integer

    Call Limpa_Tela(Me)
    
    Codigo.Text = ""
    
    'Desmarca ListBoxes e ComboBoxes
    CondicaoPagto.ListIndex = -1
    Historicos.ListIndex = -1
    Tipos.ListIndex = -1
    
    'Desmarca labels e checkboxes
    Tipo.Caption = ""
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
End Function

Private Sub CondicaoPagtoLabel_Click()
'chama browse de condicoes de pagto

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagto)
    
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagto)

End Sub

Private Sub objEventoCondicaoPagto_evSelecao(obj1 As Object)
'retorno do browse de condicoes de pagto
    
Dim objCondicaoPagto As ClassCondicaoPagto
Dim lErro As Long

On Error GoTo Erro_objEventoCondicaoPagto_evSelecao

    Set objCondicaoPagto = obj1
    
    CondicaoPagto.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagto_Validate(bSGECancelDummy)
        
    Me.Show
    
    Exit Sub
    
Erro_objEventoCondicaoPagto_evSelecao:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174870)
        
    End Select

    Exit Sub
        
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_Tela_Extrai
        
    'Informa tabela associada à Tela
    sTabela = "TiposDeFornecedor"
        
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
     
    lErro = Move_Tela_Memoria(objTipoFornecedor)
    If lErro <> SUCESSO Then Error 16378
   
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
      
    colCampoValor.Add "Codigo", objTipoFornecedor.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTipoFornecedor.sDescricao, STRING_TIPO_FORNECEDOR_DESCRICAO, "Descricao"
    colCampoValor.Add "CondicaoPagto", objTipoFornecedor.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Desconto", objTipoFornecedor.dDesconto, 0, "Desconto"
    colCampoValor.Add "Observacao", objTipoFornecedor.sObservacao, STRING_TIPO_FORNECEDOR_OBS, "Observacao"
    colCampoValor.Add "ContaDespesa", objTipoFornecedor.sContaDespesa, STRING_TIPO_FORNECEDOR_CONTA_DESPESA, "ContaDespesa"
    colCampoValor.Add "HistPadraoDespesa", objTipoFornecedor.iHistPadraoDespesa, 0, "HistPadraoDespesa"

    Exit Sub
        
Erro_Tela_Extrai:

    Select Case Err
    
        Case 16378
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174871)
        
    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_Tela_Preenche
    
    objTipoFornecedor.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTipoFornecedor.iCodigo <> 0 Then

        objTipoFornecedor.sDescricao = colCampoValor.Item("Descricao").vValor
        objTipoFornecedor.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objTipoFornecedor.dDesconto = colCampoValor.Item("Desconto").vValor
        objTipoFornecedor.sObservacao = colCampoValor.Item("Observacao").vValor
        objTipoFornecedor.sContaDespesa = colCampoValor.Item("ContaDespesa").vValor
        objTipoFornecedor.iHistPadraoDespesa = colCampoValor.Item("HistPadraoDespesa").vValor

        lErro = Exibe_Dados_TipoFornecedor(objTipoFornecedor)
        If lErro <> SUCESSO Then Error 16379
        
    End If
    
    'Desseleciona na ListBox Tipos
    Tipos.ListIndex = -1
    
    iAlterado = 0
                                                                                                       
    Exit Sub
    
Erro_Tela_Preenche:

    Select Case Err
    
        Case 16379
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174872)
        
    End Select
    
    Exit Sub
        
End Sub

Private Sub Tipos_Adiciona(objTipoFornecedor As ClassTipoFornecedor)
'Inclui na List

    'Insere Tipo na ListBox
    Tipos.AddItem objTipoFornecedor.sDescricao
    Tipos.ItemData(Tipos.NewIndex) = objTipoFornecedor.iCodigo

End Sub

Private Sub Tipos_Remove(objTipoFornecedor As ClassTipoFornecedor)
'Percorre a ListBox Tipos para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To Tipos.ListCount - 1
    
        If Tipos.ItemData(iIndice) = objTipoFornecedor.iCodigo Then
    
            Tipos.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_FORN_IDENT
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Fornecedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoFornecedor"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CondicaoPagto Then
            Call CondicaoPagtoLabel_Click
        End If
    
    End If
    
End Sub



Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Tipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Tipo, Source, X, Y)
End Sub

Private Sub Tipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Tipo, Button, Shift, X, Y)
End Sub

Private Sub LabelHistorico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistorico, Source, X, Y)
End Sub

Private Sub LabelHistorico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistorico, Button, Shift, X, Y)
End Sub

Private Sub LabelContaContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaContabil, Source, X, Y)
End Sub

Private Sub LabelContaContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaContabil, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondicaoPagtoLabel, Source, X, Y)
End Sub

Private Sub CondicaoPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondicaoPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

