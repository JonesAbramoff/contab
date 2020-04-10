VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoVendedorOcx 
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7575
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1680
      Picture         =   "TipoVendedorOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   465
      Width           =   300
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissão"
      Height          =   1470
      Left            =   120
      TabIndex        =   17
      Top             =   1530
      Width           =   4770
      Begin MSMask.MaskEdBox PercComissao 
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         Top             =   375
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercComissaoEmissao 
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         Top             =   930
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label PercComissaoBaixa 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3540
         TabIndex        =   19
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Na Baixa:"
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
         Left            =   2625
         TabIndex        =   20
         Top             =   990
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   840
         TabIndex        =   21
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Na Emissão:"
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
         Left            =   285
         TabIndex        =   22
         Top             =   990
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incide sobre"
      Height          =   1305
      Left            =   120
      TabIndex        =   18
      Top             =   3090
      Width           =   4770
      Begin VB.CheckBox ComissaoIPI 
         Caption         =   "IPI"
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
         Left            =   4005
         TabIndex        =   10
         Top             =   780
         Width           =   600
      End
      Begin VB.CheckBox ComissaoICM 
         Caption         =   "Outras Desp."
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
         Left            =   2370
         TabIndex        =   9
         Top             =   780
         Width           =   1470
      End
      Begin VB.CheckBox ComissaoSeguro 
         Caption         =   "Seguro"
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
         Left            =   1190
         TabIndex        =   8
         Top             =   780
         Width           =   990
      End
      Begin VB.CheckBox ComissaoFrete 
         Caption         =   "Frete"
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
         Left            =   210
         TabIndex        =   7
         Top             =   780
         Width           =   780
      End
      Begin VB.CheckBox ComissaoVenda 
         Caption         =   "Venda"
         Enabled         =   0   'False
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
         Left            =   1190
         TabIndex        =   6
         Top             =   375
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox ComissaoSobreTotal 
         Caption         =   "Tudo"
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
         Left            =   225
         TabIndex        =   5
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5265
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoVendedorOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoVendedorOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoVendedorOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoVendedorOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Tipos 
      Height          =   3180
      ItemData        =   "TipoVendedorOcx.ctx":0A7E
      Left            =   5235
      List            =   "TipoVendedorOcx.ctx":0A80
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1185
      Width           =   2190
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   450
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      Top             =   1035
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
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
      Height          =   210
      Left            =   135
      TabIndex        =   23
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   405
      TabIndex        =   24
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label13 
      Caption         =   "Tipos de Vendedor"
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
      Left            =   5220
      TabIndex        =   25
      Top             =   960
      Width           =   1650
   End
End
Attribute VB_Name = "TipoVendedorOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Private iEmEvento As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático
    lErro = CF("TipoVendedor_Automatico",iCodigo)
    If lErro <> SUCESSO Then Error 57559

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57559
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175067)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iEncontrou As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then Error 16240

    objTipoVendedor.iCodigo = CInt(Codigo.Text)

    lErro = CF("TipoVendedor_Le",objTipoVendedor)
    If lErro <> SUCESSO And lErro <> 16216 Then Error 16242

    'Não achou o Tipo de Vendedor --> erro
    If lErro = 16216 Then Error 16241

    'Pedido de confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPODEVENDEDOR", objTipoVendedor.iCodigo)

    If vbMsgRes = vbYes Then

        'exclui Tipo de Vendedor
        lErro = CF("TipoVendedor_Exclui",objTipoVendedor)
        If lErro <> SUCESSO Then Error 16245

        'Remove da ListBox
        Call Tipos_Remove(objTipoVendedor)

        'Limpa a Tela
        Call Limpa_Tela_TipoVendedor

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEVENDEDOR_NAO_CADASTRADO", Err, objTipoVendedor.iCodigo)

        Case 16242, 16245

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175068)

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
    If lErro <> SUCESSO Then Error 16223

    Call Limpa_Tela_TipoVendedor

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 16223
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175069)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then Error 16225

    'Verifica preenchimento da descrição
    If Len(Trim(Descricao.Text)) = 0 Then Error 16226

    If Len(Trim(PercComissaoEmissao.Text)) = 0 Then Error 33956
    If Len(PercComissaoBaixa.Caption) <> 0 Then If StrParaDbl(PercComissaoEmissao.Text) + StrParaDbl(Left(PercComissaoBaixa.Caption, Len(PercComissaoBaixa.Caption) - 1)) <> 100 Then Error 25225

    lErro = Move_Tela_Memoria(objTipoVendedor)
    If lErro <> SUCESSO Then Error 33957

    lErro = Trata_Alteracao(objTipoVendedor, objTipoVendedor.iCodigo)
    If lErro <> SUCESSO Then Error 32285

    lErro = CF("TipoVendedor_Grava",objTipoVendedor)
    If lErro <> SUCESSO Then Error 16227

    'Atualiza ListBox de Tipos
    Call Tipos_Remove(objTipoVendedor)
    Call Tipos_Adiciona(objTipoVendedor)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 32285

        Case 16225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 16227, 33957
            'Erro tratado na rotina chamada
           
        Case 25225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_EMISSAO_MAIS_BAIXA", Err)

        Case 33956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTAGEM_EMISSAO_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175070)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objTipoVendedor As ClassTipoVendedor) As Long
'Move os dados da tela para memória.

Dim lErro As Long
Dim iPosicao As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Codigo.Text)) <> 0 Then objTipoVendedor.iCodigo = CInt(Codigo.Text)

    objTipoVendedor.sDescricao = Descricao.Text

    If Len(Trim(PercComissao.Text)) <> 0 Then
        objTipoVendedor.dPercComissao = CDbl(PercComissao.Text) / 100
    Else
        objTipoVendedor.dPercComissao = 0
    End If

    If Len(Trim(PercComissaoEmissao.Text)) <> 0 Then
        objTipoVendedor.dPercComissaoEmissao = CDbl(PercComissaoEmissao.Text) / 100
    Else
        objTipoVendedor.dPercComissaoEmissao = 0
    End If

    iPosicao = InStr(PercComissaoBaixa.Caption, "%")

    If PercComissaoBaixa.Caption = "" Then
        objTipoVendedor.dPercComissaoBaixa = 0
    Else
        objTipoVendedor.dPercComissaoBaixa = CDbl(Left(PercComissaoBaixa.Caption, iPosicao - 1)) / 100
    End If

    With objTipoVendedor

        .iComissaoSobreTotal = ComissaoSobreTotal.Value
        .iComissaoFrete = ComissaoFrete.Value
        .iComissaoICM = ComissaoICM.Value
        .iComissaoIPI = ComissaoIPI.Value
        .iComissaoSeguro = ComissaoSeguro.Value

    End With

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175071)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 16270

    Call Limpa_Tela_TipoVendedor

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 16270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175072)

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

    'verifica se codigo está preenchido
    If Len(Trim(Codigo.Text)) > 0 Then

        'verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then Error 16390

        'verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 16389

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, Codigo.Text)

        Case 16390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175073)

    End Select

    Exit Sub

End Sub

Private Sub ComissaoFrete_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

       If Incidencias_Selecionadas = vbChecked Then
           ComissaoSobreTotal = vbChecked
       Else
           ComissaoSobreTotal = vbUnchecked
       End If

       iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComissaoICM_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal.Value = vbChecked
        Else
            ComissaoSobreTotal.Value = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Incidencias_Selecionadas() As CheckBoxConstants
'retorna vbChecked se a base da comissao inclui todas opcoes (frete, seguro,...)

    Incidencias_Selecionadas = ComissaoFrete.Value * ComissaoICM.Value * ComissaoIPI.Value * ComissaoSeguro.Value

End Function

Private Sub ComissaoIPI_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal.Value = vbChecked
        Else
            ComissaoSobreTotal.Value = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComissaoSeguro_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        If Incidencias_Selecionadas = vbChecked Then
            ComissaoSobreTotal.Value = vbChecked
        Else
            ComissaoSobreTotal.Value = vbUnchecked
        End If

        iEmEvento = 0

    End If

    If ComissaoFrete.Value = vbChecked And ComissaoICM.Value = vbChecked And ComissaoIPI.Value = vbChecked And ComissaoSeguro.Value = vbChecked Then ComissaoSobreTotal.Value = vbChecked

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComissaoSobreTotal_Click()

    If iEmEvento = 0 Then

        iEmEvento = 1

        'se tudo foi selecionado então marca todas as comissões
        If ComissaoSobreTotal.Value = vbChecked Then
            ComissaoFrete.Value = vbChecked
            ComissaoICM.Value = vbChecked
            ComissaoIPI.Value = vbChecked
            ComissaoSeguro.Value = vbChecked
        Else
            'desmarca todas as comissões
            ComissaoFrete.Value = vbUnchecked
            ComissaoICM.Value = vbUnchecked
            ComissaoIPI.Value = vbUnchecked
            ComissaoSeguro.Value = vbUnchecked
        End If

       iEmEvento = 0

    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Limpa_Tela_TipoVendedor()
'Limpa todos os campos de input da tela TipoVendedor

Dim lErro As Long

    Call Limpa_Tela(Me)

    'Desmarca ListBox de Tipos de Vendedor
    Tipos.ListIndex = -1
    
    Codigo.Text = ""
    
    PercComissao.Text = ""
    PercComissaoEmissao.Text = ""
    PercComissaoBaixa.Caption = ""
    
    ComissaoFrete.Value = vbUnchecked
    ComissaoICM.Value = vbUnchecked
    ComissaoIPI.Value = vbUnchecked
    ComissaoSeguro.Value = vbUnchecked
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
     'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    'Leitura dos códigos e descrições dos tipos de vendedor no BD
    lErro = CF("Cod_Nomes_Le","TiposDeVendedor", "Codigo", "Descricao", STRING_TIPO_DE_VENDEDOR_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16209

    'Preenche listbox com descricao dos tipos de vendedor
    For iIndice = 1 To colCodigoDescricao.Count
        Set objCodigoDescricao = colCodigoDescricao(iIndice)
        Tipos.AddItem objCodigoDescricao.sNome
        Tipos.ItemData(Tipos.NewIndex) = objCodigoDescricao.iCodigo
    Next

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16209

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175074)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoVendedor As ClassTipoVendedor) As Long
'Trata os parâmetros que podem ser passados quando ocorre a chamada da tela TipoVendedor

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parâmetro
    If Not (objTipoVendedor Is Nothing) Then

        lErro = CF("TipoVendedor_Le",objTipoVendedor)
        If lErro <> SUCESSO And lErro <> 16216 Then Error 16211

        If lErro = SUCESSO Then

            lErro = Traz_TipoVendedor_Tela(objTipoVendedor)
            If lErro <> SUCESSO Then Error 16380

        Else  'Não encontrou no BD

            'Limpa a Tela
            Call Limpa_Tela_TipoVendedor

            'Mostra apenas o código do Tipo de Vendedor
            Codigo.Text = CStr(objTipoVendedor.iCodigo)

        End If

    End If

    Trata_Parametros = SUCESSO

    iAlterado = 0

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16211, 16380

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175075)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Traz_TipoVendedor_Tela(objTipoVendedor As ClassTipoVendedor) As Long
'Traz os dados para tela

Dim lErro As Long

On Error GoTo Erro_Traz_TipoVendedor_Tela

    'Mostra dados na tela
    Codigo.Text = objTipoVendedor.iCodigo
    Descricao.Text = objTipoVendedor.sDescricao
    
    If objTipoVendedor.dPercComissao = 0 Then
        PercComissao.Text = ""
    Else
        PercComissao.Text = CStr(objTipoVendedor.dPercComissao * 100)
    End If
    
    PercComissaoEmissao.Text = objTipoVendedor.dPercComissaoEmissao * 100
    PercComissaoBaixa.Caption = Format(objTipoVendedor.dPercComissaoBaixa * 100, "#0.#0\%")
    
    ComissaoSobreTotal.Value = objTipoVendedor.iComissaoSobreTotal
    ComissaoICM.Value = objTipoVendedor.iComissaoICM
    ComissaoIPI.Value = objTipoVendedor.iComissaoIPI
    ComissaoFrete.Value = objTipoVendedor.iComissaoFrete
    ComissaoSeguro.Value = objTipoVendedor.iComissaoSeguro

    iAlterado = 0

    Traz_TipoVendedor_Tela = SUCESSO

Exit Function

Erro_Traz_TipoVendedor_Tela:

    Traz_TipoVendedor_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175076)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub PercComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercComissao_Validate(Cancel As Boolean)

Dim lErro As String
Dim sPercComissao As String

On Error GoTo Erro_PercComissao_Validate

    sPercComissao = PercComissao.Text

    'Verifica se foi preenchido a PercComissao
    If Len(Trim(PercComissao.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(PercComissao.Text)
    If lErro <> SUCESSO Then Error 33955

    PercComissao.Text = Format(sPercComissao, "Fixed")

    Exit Sub

Erro_PercComissao_Validate:

    Cancel = True


    Select Case Err

        Case 33955

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175077)

    End Select

    Exit Sub

End Sub

Private Sub PercComissaoEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercComissaoEmissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercComissaoEmissao As Double
Dim dPercComissaoBaixa As Double

On Error GoTo Erro_PercComissaoEmissao_Validate

    If Len(Trim(PercComissaoEmissao.Text)) <> 0 Then

        lErro = Porcentagem_Critica(PercComissaoEmissao.Text)
        If lErro <> SUCESSO Then Error 16219

        PercComissaoEmissao.Text = Format(PercComissaoEmissao.Text, "Fixed")

        dPercComissaoEmissao = CDbl(PercComissaoEmissao.Text)

        dPercComissaoBaixa = 100 - dPercComissaoEmissao

        PercComissaoBaixa.Caption = Format(dPercComissaoBaixa, "#0.#0\%")

    End If

    Exit Sub

Erro_PercComissaoEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 16219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175078)

    End Select

    Exit Sub

End Sub

Private Sub Tipos_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_Tipos_DblClick

    objTipoVendedor.iCodigo = Tipos.ItemData(Tipos.ListIndex)

    lErro = CF("TipoVendedor_Le",objTipoVendedor)
    If lErro <> SUCESSO And lErro <> 16216 Then Error 16217

    'Não encontrou o Tipo de Vendedor --> erro
    If lErro <> SUCESSO Then Error 16218

    lErro = Traz_TipoVendedor_Tela(objTipoVendedor)
    If lErro <> SUCESSO Then Error 33952

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Tipos_DblClick:

    Select Case Err

        Case 16217, 33952

        Case 16218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEVENDEDOR_NAO_CADASTRADO", Err, objTipoVendedor.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175079)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
     objTipoVendedor.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTipoVendedor.iCodigo <> 0 Then

        objTipoVendedor.sDescricao = colCampoValor.Item("Descricao").vValor
        objTipoVendedor.dPercComissaoEmissao = (colCampoValor.Item("PercComissaoEmissao").vValor)
        objTipoVendedor.dPercComissaoBaixa = (colCampoValor.Item("PercComissaoBaixa").vValor)
        objTipoVendedor.dPercComissao = (colCampoValor.Item("PercComissao").vValor)
        objTipoVendedor.iComissaoSobreTotal = colCampoValor.Item("ComissaoSobreTotal").vValor
        objTipoVendedor.iComissaoFrete = colCampoValor.Item("ComissaoFrete").vValor
        objTipoVendedor.iComissaoSeguro = colCampoValor.Item("ComissaoSeguro").vValor
        objTipoVendedor.iComissaoIPI = colCampoValor.Item("ComissaoIPI").vValor
        objTipoVendedor.iComissaoICM = colCampoValor.Item("ComissaoICM").vValor

        'Chama Traz_TipoVendedor_Tela
        lErro = Traz_TipoVendedor_Tela(objTipoVendedor)
        If lErro <> SUCESSO Then Error 33953

    End If

    'Desseleciona na ListBox Tipos
    Tipos.ListIndex = -1

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 33953

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175080)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTipoVendedor As New ClassTipoVendedor

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDeVendedor"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
     lErro = Move_Tela_Memoria(objTipoVendedor)
    If lErro <> SUCESSO Then Error 33954

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTipoVendedor.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTipoVendedor.sDescricao, STRING_TIPO_DE_VENDEDOR_DESCRICAO, "Descricao"
    colCampoValor.Add "PercComissao", objTipoVendedor.dPercComissao, 0, "PercComissao"
    colCampoValor.Add "PercComissaoBaixa", objTipoVendedor.dPercComissaoBaixa, 0, "PercComissaoBaixa"
    colCampoValor.Add "PercComissaoEmissao", objTipoVendedor.dPercComissaoEmissao, 0, "PercComissaoEmissao"
    colCampoValor.Add "ComissaoSobreTotal", objTipoVendedor.iComissaoSobreTotal, 0, "ComissaoSobreTotal"
    colCampoValor.Add "ComissaoFrete", objTipoVendedor.iComissaoFrete, 0, "ComissaoFrete"
    colCampoValor.Add "ComissaoICM", objTipoVendedor.iComissaoICM, 0, "ComissaoICM"
    colCampoValor.Add "ComissaoIPI", objTipoVendedor.iComissaoIPI, 0, "ComissaoIPI"
    colCampoValor.Add "ComissaoSeguro", objTipoVendedor.iComissaoSeguro, 0, "ComissaoSeguro"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 33954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175081)

    End Select

    Exit Sub

End Sub

Private Sub Tipos_Adiciona(objTipoVendedor As ClassTipoVendedor)
'Inclui tipo na List

    'Insere Tipo na ListBox
    Tipos.AddItem objTipoVendedor.sDescricao
    Tipos.ItemData(Tipos.NewIndex) = objTipoVendedor.iCodigo

End Sub

Private Sub Tipos_Remove(objTipoVendedor As ClassTipoVendedor)
'Percorre a ListBox Tipos para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To Tipos.ListCount - 1
    
        If Tipos.ItemData(iIndice) = objTipoVendedor.iCodigo Then
    
            Tipos.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_VENDEDOR
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Vendedor"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoVendedor"
    
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
End Sub


Private Sub PercComissaoBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercComissaoBaixa, Source, X, Y)
End Sub

Private Sub PercComissaoBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercComissaoBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

