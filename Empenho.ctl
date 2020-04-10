VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl Empenho 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   KeyPreview      =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   5805
   Begin VB.Frame Frame4 
      Caption         =   "Item da Ordem de Produção"
      Height          =   1590
      Left            =   210
      TabIndex        =   23
      Top             =   735
      Width           =   5436
      Begin VB.ComboBox ComboItem 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   330
         Width           =   615
      End
      Begin VB.Label ItemLabel 
         Caption         =   "Item:"
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
         Height          =   270
         Left            =   360
         TabIndex        =   32
         Top             =   352
         Width           =   435
      End
      Begin VB.Label DescricaoItemOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1530
         TabIndex        =   31
         Top             =   330
         Width           =   3555
      End
      Begin VB.Label UMItemOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4410
         TabIndex        =   30
         Top             =   1155
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   3870
         TabIndex        =   29
         Top             =   1215
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Qtde a Produzir:"
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
         Left            =   75
         TabIndex        =   28
         Top             =   1215
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Versão:"
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
         Left            =   810
         TabIndex        =   27
         Top             =   810
         Width           =   660
      End
      Begin VB.Label SaldoItemOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1530
         TabIndex        =   26
         Top             =   1155
         Width           =   1155
      End
      Begin VB.Label Versao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1530
         TabIndex        =   25
         Top             =   750
         Width           =   3555
      End
   End
   Begin VB.TextBox CodigoOP 
      Height          =   300
      Left            =   810
      MaxLength       =   6
      TabIndex        =   22
      Top             =   240
      Width           =   1260
   End
   Begin VB.CommandButton BotaoEmpenhos 
      Caption         =   "Empenhos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4050
      TabIndex        =   21
      Top             =   5220
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Componente"
      Height          =   2715
      Left            =   210
      TabIndex        =   5
      Top             =   2430
      Width           =   5436
      Begin VB.ComboBox ComboProduto 
         Height          =   315
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   150
         Width           =   3855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quantidades"
         Height          =   744
         Left            =   168
         TabIndex        =   7
         Top             =   1800
         Width           =   5076
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1215
            TabIndex        =   8
            Top             =   278
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "Empenhada:"
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
            Height          =   240
            Left            =   132
            TabIndex        =   11
            Top             =   315
            Width           =   1104
         End
         Begin VB.Label Label6 
            Caption         =   "Requisitada:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   2676
            TabIndex        =   10
            Top             =   333
            Width           =   1092
         End
         Begin VB.Label Requisitada 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3780
            TabIndex        =   9
            Top             =   285
            Width           =   1152
         End
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
         Height          =   315
         Left            =   3660
         TabIndex        =   6
         Top             =   990
         Width           =   1530
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2415
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1410
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1335
         TabIndex        =   14
         Top             =   1410
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   315
         Left            =   1335
         TabIndex        =   15
         Top             =   990
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label UM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1335
         TabIndex        =   20
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Left            =   795
         TabIndex        =   19
         Top             =   630
         Width           =   480
      End
      Begin VB.Label AlmoxarifadoLabel 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label ProdutoLabel 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   540
         TabIndex        =   17
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Left            =   510
         TabIndex        =   16
         Top             =   1470
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3495
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Empenho.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Empenho.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Empenho.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Empenho.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label OPLabel 
      AutoSize        =   -1  'True
      Caption         =   "O.P.:"
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
      Left            =   300
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   300
      Width           =   450
   End
End
Attribute VB_Name = "Empenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iCodigoAlterado As Integer
Dim iCodigoOPAlterado As Integer

Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1

Private Sub Limpa_Alterados()

    iAlterado = 0
    iCodigoAlterado = 0
    iCodigoOPAlterado = 0

End Sub

Private Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Almoxarifado_Validate

    lErro = Almoxarifado_Critica()
    If lErro <> SUCESSO Then gError 55475

    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True


    Select Case gErr

        Case 55475
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159425)

    End Select

    Exit Sub

End Sub

Private Function Almoxarifado_Critica() As Long

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProduto As String

On Error GoTo Erro_Almoxarifado_Critica

    If Len(Trim(Almoxarifado.Text)) <> 0 Then

        lErro = TP_Almoxarifado_Filial_Le(Almoxarifado, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 22984

        If lErro = 25136 Then gError 22985

        If lErro = 25143 Then gError 22986

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        'Verifica se o Produto foi preenchido
        If ComboProduto.ListIndex <> -1 Then
        
            sProduto = SCodigoProduto_Extrai(ComboProduto.Text)
    
            'Passa para o formato do BD
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
            If lErro <> SUCESSO Then gError 33144
    
            'testa se o codigo está preenchido
            If iPreenchido = PRODUTO_PREENCHIDO Then
    
                objEstoqueProduto.sProduto = sProdutoFormatado
    
                'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
                lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                If lErro <> SUCESSO And lErro <> 21306 Then gError 55485
    
                'Se não encontrou EstoqueProduto no Banco de Dados
                If lErro = 21306 Then gError 55486

            End If
    
        End If

    End If

    Almoxarifado_Critica = SUCESSO

    Exit Function

Erro_Almoxarifado_Critica:

    Almoxarifado_Critica = gErr

    Select Case gErr

        Case 22984, 55485

        Case 22985, 22986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, Almoxarifado.Text)

        Case 55486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", gErr, objEstoqueProduto.iAlmoxarifado, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159426)

    End Select

    Exit Function

End Function

Private Sub BotaoEmpenhos_Click()

Dim colSelecao As New Collection
Dim objEmpenho As New ClassEmpenho, sProduto As String
Dim lErro As Long

On Error GoTo Erro_BotaoEmpenhos_Click

    lErro = Move_Tela_Memoria(objEmpenho)
    If lErro <> SUCESSO Then gError 41882

    Call Chama_Tela("EmpenhoLista", colSelecao, objEmpenho, objEventoCodigo)

    Exit Sub
    
Erro_BotaoEmpenhos_Click:

    Select Case gErr
          
        Case 41882
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159427)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objEmpenho As New ClassEmpenho
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = VerificaChavePreenchida
    If lErro <> SUCESSO Then gError 41883
    
    lErro = Move_Tela_Memoria(objEmpenho)
    If lErro <> SUCESSO Then gError 41885
    
    'Verifica se o Empenho existe
    lErro = CF("Empenho_Le_SemCodigo", objEmpenho)
    If lErro <> SUCESSO And lErro <> 41889 Then gError 33148

    'Não encontrou o Empenho ==> Erro
    If lErro <> SUCESSO Then gError 33149

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_EMPENHO")

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o Empenho
    lErro = CF("Empenho_Exclui", objEmpenho)
    If lErro <> SUCESSO Then gError 33150

    Call Limpa_Tela_Empenho

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Alterados

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 33148, 33150, 41883, 41885

        Case 33149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPENHO_NAO_CADASTRADO", gErr, objEmpenho.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159428)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Cliente
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 33129

    Call Limpa_Tela_Empenho

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Alterados

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 33129

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159429)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 33127

    Call Limpa_Tela_Empenho

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Alterados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 33127

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159430)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoOP_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoOPAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigoOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao
Dim objEmpenho As New ClassEmpenho

On Error GoTo Erro_CodigoOP_Validate

    'Se houve alteração nos dados da tela
    If iCodigoOPAlterado = REGISTRO_ALTERADO Then

        If Len(Trim(CodigoOP.Text)) > 0 Then

            lErro = PreencheItensOP(CodigoOP.Text, giFilialEmpresa)
            If lErro <> SUCESSO Then gError 41365
            
        End If
        
        iCodigoOPAlterado = 0

    End If

    Exit Sub

Erro_CodigoOP_Validate:

    Cancel = True


    Select Case gErr

        Case 41365
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159431)
    
    End Select

    Exit Sub

End Sub

Private Sub ComboItem_Click()

Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim dQuantResultado As Double
Dim dQuantOrdenada As Double
Dim dQuantProduzida As Double

On Error GoTo Erro_ComboItem_Click

    'Verifica se foram preenchidos o Codigo da OP e o Item
    If ComboItem.ListIndex = -1 Or Len(Trim(CodigoOP.Text)) = 0 Then

        DescricaoItemOP.Caption = ""
        SaldoItemOP.Caption = ""
        UMItemOP.Caption = ""
        ComboProduto.Clear
        UM.Caption = ""
        Versao.Caption = ""
        Exit Sub

    End If

    objItemOP.iItem = ComboItem.ItemData(ComboItem.ListIndex)
    objItemOP.sCodigo = CodigoOP.Text
    objItemOP.iFilialEmpresa = giFilialEmpresa

    'Lê o Item da Ordem de Produção
    lErro = CF("ItemOP_Le_Item", objItemOP)
    If lErro <> SUCESSO And lErro <> 33215 Then gError 33209

    If lErro = SUCESSO Then

        sProduto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProduto)
        If lErro <> SUCESSO Then gError 55972

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 33210

        If lErro = 25041 Then gError 33211

        'Preenche a Descrição do Item da OP
        DescricaoItemOP.Caption = objItemOP.sProduto & SEPARADOR & objProduto.sDescricao

        dQuantOrdenada = objItemOP.dQuantidade
        dQuantProduzida = objItemOP.dQuantidadeProd

        'Em "saldo" colocar a qtde (ordenada) do item menos a qtde produzida
        dQuantResultado = dQuantOrdenada - dQuantProduzida
        
        If dQuantResultado < 0 Then dQuantResultado = 0

        'Preenche o Saldo do Item da OP
        SaldoItemOP.Caption = Formata_Estoque(dQuantResultado)

        'Preenche a Sigla da Unidade de Medida
        UMItemOP.Caption = objItemOP.sSiglaUM
        
        'Preenche a Versao do Kit
        Versao.Caption = objItemOP.sVersao

    End If
    
    'Verifica se Item faz parte de um Kit
    lErro = Verifica_Kit(objItemOP)
    If lErro <> SUCESSO Then gError 43809
        
    Exit Sub

Erro_ComboItem_Click:

    Select Case gErr

        Case 33209, 33210, 43809
    
        Case 33211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Error, sProduto)
    
        Case 55972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemOP.sProduto)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159432)
    
    End Select

    Exit Sub

End Sub

Private Function Verifica_Kit(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim iIndice As Long
Dim objKit As New ClassKit
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos

On Error GoTo Erro_Verifica_Kit

    'Verifica se ComboItem foi preenchida
    If ComboItem.ListIndex = -1 Then
        ComboItem.Clear
        Exit Function
    End If
    
    ComboProduto.Clear
    UM.Caption = ""
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then objKit.dtData = CDate(DataEmissao.Text)
    
    objKit.sProdutoRaiz = objItemOP.sProduto
    
    'Verifica se o Produto é Kit
    lErro = CF("Kit_Le_Empenho", objKit)
    If lErro <> SUCESSO And lErro <> 43814 Then gError 43810
    
    'Se encontrou o Produto no Kit
    If lErro = SUCESSO Then
        
        'Alteracao Daniel em 29/07/2002
        objProdutoKitProdutos.sProdutoRaiz = objItemOP.sProduto
        objProdutoKitProdutos.sVersao = objItemOP.sVersao
        objProdutoKitProdutos.dQuantidade = objItemOP.dQuantidade
        objProdutoKitProdutos.iClasseUM = objItemOP.iClasseUM
        objProdutoKitProdutos.sUnidadeMed = objItemOP.sSiglaUM
        
        'Obtem uma Colecao com os itens do kit para empenho (a embalagem do ProdutoRaiz será o ultimo elemento da colecao)
        lErro = CF("OrdemProducao_Le_Col_Empenho", objProdutoKitProdutos, objKit.colComponentes)
        If lErro <> SUCESSO And lErro <> 106393 Then gError 106435
    
        'Se nao Encontrou => Erro
        If lErro = 106393 Then gError 106436
        
        For iIndice = 1 To objKit.colComponentes.Count
        
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_MascararProduto(objKit.colComponentes.Item(iIndice).sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 48743
            
            objProduto.sCodigo = objKit.colComponentes.Item(iIndice).sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 106438
            
            'se o produto não estiver cadastrado
            If lErro = 28030 Then gError 106439

            ComboProduto.AddItem sProdutoMascarado & " " & SEPARADOR & " " & objProduto.sDescricao

        Next

        ComboProduto.ListIndex = -1
        
    End If
    
    Verifica_Kit = SUCESSO
    
    Exit Function
    
Erro_Verifica_Kit:

    Verifica_Kit = gErr
    
    Select Case gErr
    
        Case 43810, 43816, 106435, 106438
        
        Case 43817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, objKit.sProdutoRaiz)
    
        Case 106436
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, objKit.sProdutoRaiz)
        
        Case 106439
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
                    
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159433)
    
    End Select
    
    Exit Function
 
End Function

Private Sub ComboProduto_Click()
'O produto não pode ser igual ao que se vai produzir (o do item da op)
'preencher o campo de UM com a UM de estoque do produto

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim iCodigo As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_ComboProduto_Click

    'Verifica preenchimento de Produto
    If ComboProduto.ListIndex <> -1 Then

        sProduto = SCodigoProduto_Extrai(ComboProduto.Text)

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 33122

        If lErro = 25041 Then gError 33123

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 22965

        'Formata o código do Produto como no BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 19768

        objProduto.sCodigo = sProdutoFormatado

        'Preenche a Unidade de Medida
        UM.Caption = objProduto.sSiglaUMEstoque

        'Comentado por Daniel em 02/08/2002
        'Agora um Produto pode fazer parte da composição dele mesmo.
        'O produto não pode ser igual ao que se vai produzir (o do item da op)
        'lErro = Critica_Item(objProduto.sCodigo)
        'If lErro <> SUCESSO Then gError 33221

    Else
        UM.Caption = ""

    End If

    Exit Sub

Erro_ComboProduto_Click:

    Select Case gErr

        Case 22965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 33122, 33221, 19768

        Case 33123 'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159434)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se foi preenchida a Data
    If Len(Trim(DataEmissao.Text)) = 0 Then Exit Sub

    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 33108

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 33108
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159435)

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
Dim lCodigo As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Form_Load

    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    Set objEventoOP = New AdmEvento
    Set objEventoEstoque = New AdmEvento

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 33117

    Call Limpa_Tela_Empenho

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Preenche a Data de Emissão
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    Quantidade.Format = FORMATO_ESTOQUE

    Call Limpa_Alterados

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 33117

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159436)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoAlmoxPadrao = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoOP = Nothing
    Set objEventoEstoque = Nothing

   'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As New ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche o Almoxarifado Padrao
    Almoxarifado.Text = CStr(objAlmoxarifado.sNomeReduzido)

    Me.Show

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEmpenho As New ClassEmpenho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objEmpenho = obj1

    lErro = Traz_Empenho_Tela(objEmpenho)
    If lErro <> SUCESSO Then gError 33231

    Call Limpa_Alterados

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 33231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159437)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOP As New ClassOrdemDeProducao

    Set objOP = obj1

    'Preenche o Código da OP
    CodigoOP.Text = objOP.sCodigo

    Call CodigoOP_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Private Sub OPLabel_Click()

Dim colSelecao As New Collection
Dim objOP As New ClassOrdemDeProducao

    'Chama a tela de browse
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOP)

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    If Len(Trim(Quantidade.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 41321

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True


    Select Case gErr

        Case 41321

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159438)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    DataEmissao.SetFocus

    If Len(DataEmissao.ClipText) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 33109

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 33109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159439)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    DataEmissao.SetFocus

    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 33110

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 33110

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159440)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_Empenho() As Long
'Limpa os campos tela Empenho

Dim lErro As Long

    'Função generica que limpa campos da tela
    Call Limpa_Tela(Me)

    CodigoOP.Text = ""
    ComboItem.Clear
    DescricaoItemOP.Caption = ""
    SaldoItemOP.Caption = ""
    UMItemOP.Caption = ""
    ComboProduto.Clear
    UM.Caption = ""
    Requisitada.Caption = ""
    Versao.Caption = ""
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

End Function

Function Gravar_Registro() As Long
'Verifica se dados de Empenho necessários foram preenchidos

Dim lErro As Long
Dim objEmpenho As New ClassEmpenho

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = VerificaChavePreenchida()
    If lErro <> SUCESSO Then gError 41884
    
    'Verifica se foi preenchida a Data de Emissão
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 33136

    'Verifica se foi preenchida a Quantidade Empenhada
    If Len(Trim(Quantidade.Text)) = 0 Then gError 33137

    'Lê os dados da Tela relacionados ao Empenho
    lErro = Move_Tela_Memoria(objEmpenho)
    If lErro <> SUCESSO Then gError 33138

    'Grava o Cliente no BD
    lErro = CF("Empenho_Grava", objEmpenho)
    If lErro <> SUCESSO Then gError 33139

    Call Limpa_Alterados

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 33136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 33137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_EMPENHADA_NAO_PRRENCHIDA", gErr)

        Case 33138, 33139, 41884

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159441)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objEmpenho As ClassEmpenho) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Move_Tela_Memoria

    objEmpenho.iFilialEmpresa = giFilialEmpresa
    objEmpenho.lCodigo = 0
    
    'Verifica se o Produto foi preenchido
    If ComboProduto.ListIndex <> -1 Then
    
        sProduto = SCodigoProduto_Extrai(ComboProduto.Text)

        'Passa para o formato do BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
        If lErro <> SUCESSO Then gError 33144

        'testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then objEmpenho.sProduto = sProdutoFormatado

    End If

   'Recolhe os demais dados
   If Len(Trim(Almoxarifado.Text)) > 0 Then

        objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
        
        'Lê o Almoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 33145

        'Se não encontrou o Nome Reduzido do Almoxarifado
        If lErro <> SUCESSO Then gError 33146

        objEmpenho.iAlmoxarifado = objAlmoxarifado.iCodigo

   End If

   If Len(Trim(Quantidade.Text)) > 0 Then
        objEmpenho.dQuantidade = CDbl(Quantidade.Text)
   End If

   If Len(Trim(Requisitada.Caption)) > 0 Then
    objEmpenho.dQuantidadeRequisitada = CDbl(Requisitada.Caption)
   End If

   If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objEmpenho.dtData = CDate(DataEmissao.Text)
   End If

   If Len(Trim(CodigoOP.Text)) > 0 Then
        objEmpenho.sCodigoOP = CodigoOP.Text
   End If

   If ComboItem.ListIndex <> -1 Then
        objEmpenho.iItemOP = ComboItem.ItemData(ComboItem.ListIndex)
    End If

   Move_Tela_Memoria = SUCESSO

   Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 33144, 33145

        Case 33146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159442)

    End Select

    Exit Function

End Function

Private Function PreencheItensOP(sCodigoOP As String, iFilialEmpresa As Integer) As Long
'preenche a combo de itens da OP

Dim lErro As Long
Dim objOrdemOP As New ClassOrdemDeProducao
Dim sProdutoEnxuto As String, iIndice As Integer

On Error GoTo Erro_PreencheItensOP

    'Preenche objOrdemOP
    objOrdemOP.iFilialEmpresa = iFilialEmpresa
    objOrdemOP.sCodigo = sCodigoOP
    
    'Lê os Itens da Ordem de Produção
    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemOP)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 43807
    
    'Se não encontrou a Ordem de Produção --> Erro
    If lErro <> SUCESSO Then gError 43808

    ComboItem.Clear
    
    'Preenche a ComboItem
    For iIndice = 1 To objOrdemOP.colItens.Count
    
        ComboItem.AddItem objOrdemOP.colItens.Item(iIndice).iItem
        ComboItem.ItemData(ComboItem.NewIndex) = objOrdemOP.colItens.Item(iIndice).iItem

    Next
    
    PreencheItensOP = SUCESSO
     
    Exit Function
    
Erro_PreencheItensOP:

    PreencheItensOP = gErr
     
    Select Case gErr
          
        Case 43807
        
        Case 43808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemOP.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159443)
     
    End Select
     
    Exit Function

End Function

Private Function Traz_Empenho_Tela(objEmpenho As ClassEmpenho) As Long
'Traz os dados do Empenho para tela objEmpenho.lCodigo tem que estar preenchido

Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objItemOP As New ClassItemOP, iIndice As Integer
Dim iAchou As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_Empenho_Tela

    objEmpenho.iFilialEmpresa = giFilialEmpresa

    'Lê o Empenho
    lErro = CF("Empenho_Le", objEmpenho)
    If lErro <> SUCESSO And lErro <> 33114 Then gError 33227

    'Se não encontrou o Empenho ==> Erro
    If lErro <> SUCESSO Then gError 33228

    'Limpa a Tela
    Call Limpa_Tela_Empenho

    objItemOP.lNumIntDoc = objEmpenho.lNumIntDocItemOP
    objItemOP.iFilialEmpresa = objEmpenho.iFilialEmpresa

    'Lê o item do Empenho
    lErro = CF("ItemOP_Le_NumIntDoc", objItemOP)
    If lErro <> SUCESSO And lErro <> 33226 Then gError 33229

    'Se não encontrou o item do empenho ==> Erro
    If lErro <> SUCESSO Then gError 33230

    CodigoOP.Text = CStr(objItemOP.sCodigo)
    
    'preenche a combo de itens
    lErro = PreencheItensOP(objItemOP.sCodigo, objItemOP.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 41890
    
    'seleciona o item da OP na combo
    For iIndice = 0 To ComboItem.ListCount - 1
        
        If ComboItem.ItemData(iIndice) = objItemOP.iItem Then
        
            'vai preencher a combo de produtos com os componentes
            ComboItem.ListIndex = iIndice
            Exit For
            
        End If
    
    Next

    iAchou = 0

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_MascararProduto(objEmpenho.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 55483
        
    sProdutoMascarado = Trim(sProdutoMascarado)
    
    'seleciona o produto na combo
    For iIndice = 0 To ComboProduto.ListCount - 1
        
        If SCodigoProduto_Extrai(ComboProduto.List(iIndice)) = sProdutoMascarado Then
        
            ComboProduto.ListIndex = iIndice
            iAchou = 1
            Exit For
            
        End If
    
    Next

    If iAchou = 0 Then gError 55484
        
    Almoxarifado.Text = CStr(objEmpenho.iAlmoxarifado)
    
    lErro = Almoxarifado_Critica()
    If lErro <> SUCESSO Then gError 55476

    DataEmissao.Text = Format(objEmpenho.dtData, "dd/mm/yy")

    Quantidade.Text = Formata_Estoque(objEmpenho.dQuantidade)

    Requisitada.Caption = Formata_Estoque(objEmpenho.dQuantidadeRequisitada)
    
    iAlterado = 0

    Traz_Empenho_Tela = SUCESSO

    Exit Function

Erro_Traz_Empenho_Tela:

    Traz_Empenho_Tela = gErr

    Select Case gErr

        Case 33227, 33229, 41890, 55476, 55478

        Case 33228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPENHO_NAO_CADASTRADO", gErr, objEmpenho.lCodigo)

        Case 33230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_EMPENHO_NAO_CADASTRADO", gErr, objItemOP.lNumIntDoc, objItemOP.iFilialEmpresa)

        Case 55483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objEmpenho.sProduto)
        
        Case 55484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPONENTE", gErr, objEmpenho.sProduto, objItemOP.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159444)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objEmpenho As New ClassEmpenho

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Empenho_OP"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objEmpenho)
    If lErro <> SUCESSO Then gError 33158

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objEmpenho.lCodigo, 0, "Codigo"
    colCampoValor.Add "CodigoOP", objEmpenho.sCodigoOP, STRING_ORDEM_DE_PRODUCAO, "CodigoOP"
    colCampoValor.Add "Item", objEmpenho.iItemOP, 0, "Item"
    colCampoValor.Add "Produto", objEmpenho.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Almoxarifado", objEmpenho.iAlmoxarifado, 0, "Almoxarifado"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objEmpenho.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 33158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159445)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objEmpenho As New ClassEmpenho

On Error GoTo Erro_Tela_Preenche

    objEmpenho.lCodigo = colCampoValor.Item("Codigo").vValor

    If objEmpenho.lCodigo <> 0 Then

        'Carrega objEmpenho com os dados passados em colCampoValor
        objEmpenho.lCodigo = colCampoValor.Item("Codigo").vValor

        'Chama Traz_Empenho_Tela
        lErro = Traz_Empenho_Tela(objEmpenho)
        If lErro <> SUCESSO Then gError 33159

        Call Limpa_Alterados

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 33159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159446)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objEmpenho As ClassEmpenho) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Trata_Parametros

    If Not (objEmpenho Is Nothing) Then

        lErro = Traz_Empenho_Tela(objEmpenho)
        If lErro <> SUCESSO Then gError 33170

        Call Limpa_Alterados

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 33170

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159447)

    End Select

    iAlterado = 0

    Exit Function

End Function

'Private Function Critica_Item(sProduto As String) As Long
''Critica se o produto é igual ao que se vai produzir (o do item da op)
'
'Dim lErro As Long
'Dim sItem As String
'
'On Error GoTo Erro_Critica_Item
'
'    If Len(Trim(DescricaoItemOP.Caption)) = 0 Then Exit Function
'
'    sItem = SCodigo_Extrai(DescricaoItemOP.Caption)
'
'    If sProduto = sItem Then gError 33222
'
'    Critica_Item = SUCESSO
'
'    Exit Function
'
'Erro_Critica_Item:
'
'    Critica_Item = gErr
'
'    Select Case gErr
'
'        Case 33222
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_IGUAL_ITEM_OP", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159448)
'
'    End Select
'
'    Exit Function
'
'End Function
    
Private Function VerificaChavePreenchida() As Long

Dim lErro As Long

On Error GoTo Erro_VerificaChavePreenchida

    'Verifica se foi preenchida a Ordem de Produção
    If Len(Trim(CodigoOP.Text)) = 0 Then gError 33132

    'Verifica se foi preenchido do Item
    If ComboItem.ListIndex = -1 Then gError 33133

    'Verifica se foi preenchido o Produto
    If ComboProduto.ListIndex = -1 Then gError 33134

    'Verifica se foi preenchido o Almoxarifado
    If Len(Trim(Almoxarifado.Text)) = 0 Then gError 33135

    VerificaChavePreenchida = SUCESSO
     
    Exit Function
    
Erro_VerificaChavePreenchida:

    VerificaChavePreenchida = gErr
     
    Select Case gErr
          
        Case 33132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)

        Case 33133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_PREENCHIDO", gErr)

        Case 33134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 33135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159449)
     
    End Select
     
    Exit Function

End Function

Private Sub BotaoEstoque_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If Len(Trim(ComboProduto.Text)) = 0 Then gError 41896
    
    sCodProduto = SCodigoProduto_Extrai(ComboProduto.Text)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41894

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
        
    Else
    
        gError 41895
        
    End If
    
    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 41894
        
        Case 41895, 41896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159450)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159451)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMPENHO
    Set Form_Load_Ocx = Me
    Caption = "Empenho"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Empenho"
    
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
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CodigoOP Then
            Call OPLabel_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        End If
    End If

End Sub


Private Sub Requisitada_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Requisitada, Source, X, Y)
End Sub

Private Sub Requisitada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Requisitada, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoLabel, Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub UM_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UM, Source, X, Y)
End Sub

Private Sub UM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UM, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub SaldoItemOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoItemOP, Source, X, Y)
End Sub

Private Sub SaldoItemOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoItemOP, Button, Shift, X, Y)
End Sub

Private Sub UMItemOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UMItemOP, Source, X, Y)
End Sub

Private Sub UMItemOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UMItemOP, Button, Shift, X, Y)
End Sub

Private Sub DescricaoItemOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoItemOP, Source, X, Y)
End Sub

Private Sub DescricaoItemOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoItemOP, Button, Shift, X, Y)
End Sub

Private Sub ItemLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ItemLabel, Source, X, Y)
End Sub

Private Sub ItemLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ItemLabel, Button, Shift, X, Y)
End Sub

Private Sub OPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OPLabel, Source, X, Y)
End Sub

Private Sub OPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OPLabel, Button, Shift, X, Y)
End Sub

Private Function SCodigoProduto_Extrai(ByVal sProdDesc As String) As String
    Dim sAux As String, iPos As Integer
    iPos = InStr(1, sProdDesc, " ")
    SCodigoProduto_Extrai = Trim(left(sProdDesc, iPos))
End Function

