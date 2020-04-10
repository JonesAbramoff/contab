VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PadraoTribEntradaOcx 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6645
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PadraoTribEntradaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PadraoTribEntradaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PadraoTribEntradaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PadraoTribEntradaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton PesquisarPadroesTrib 
      Caption         =   "Padrões Existentes ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   2088
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   2676
      Left            =   135
      TabIndex        =   12
      Top             =   765
      Width           =   6276
      Begin VB.ComboBox TipoDocumento 
         Height          =   315
         ItemData        =   "PadraoTribEntradaOcx.ctx":0994
         Left            =   1452
         List            =   "PadraoTribEntradaOcx.ctx":0996
         TabIndex        =   1
         Top             =   276
         Width           =   4740
      End
      Begin VB.Frame Frame3 
         Caption         =   "Produtos"
         Height          =   1356
         Left            =   144
         TabIndex        =   13
         Top             =   1140
         Width           =   4464
         Begin VB.ComboBox ItemCategoriaProduto 
            Height          =   315
            Left            =   1605
            TabIndex        =   5
            Top             =   960
            Width           =   2448
         End
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            ItemData        =   "PadraoTribEntradaOcx.ctx":0998
            Left            =   1230
            List            =   "PadraoTribEntradaOcx.ctx":099A
            TabIndex        =   4
            Top             =   564
            Width           =   2856
         End
         Begin VB.CheckBox TodosProdutos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   192
            TabIndex        =   3
            Top             =   240
            Value           =   1  'Checked
            Width           =   1320
         End
         Begin VB.Label Label5 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   228
            TabIndex        =   15
            Top             =   600
            Width           =   936
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
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
            Height          =   192
            Left            =   996
            TabIndex        =   14
            Top             =   1020
            Width           =   504
         End
      End
      Begin MSMask.MaskEdBox NaturezaOp 
         Height          =   300
         Left            =   1455
         TabIndex        =   2
         Top             =   750
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoDoc 
         Caption         =   "Tipo de Documento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   144
         TabIndex        =   18
         Top             =   252
         Width           =   1152
      End
      Begin VB.Label DescrNatOp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2055
         TabIndex        =   17
         Top             =   750
         Width           =   4110
      End
      Begin VB.Label NaturezaLabel 
         AutoSize        =   -1  'True
         Caption         =   "Natureza Oper.:"
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
         Height          =   192
         Left            =   132
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   792
         Width           =   1308
      End
   End
   Begin MSMask.MaskEdBox TipoTrib 
      Height          =   300
      Left            =   1965
      TabIndex        =   6
      Top             =   3630
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label DescrTipoTrib 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2580
      TabIndex        =   20
      Top             =   3630
      Width           =   3810
   End
   Begin VB.Label TipoLabel 
      Caption         =   "Tipo de Tributação:"
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
      Height          =   225
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   19
      Top             =   3675
      Width           =   1695
   End
End
Attribute VB_Name = "PadraoTribEntradaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'responsavel: Jones
'revisada em: 12/09/98
'pendencias:
        
'sugestoes:
    'implementar browses p/categorias e colocar label com descricao completa do tipo doc
    
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iNaturezaAlterado As Integer
Dim iTipoAlterado As Integer

Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoPadrao As AdmEvento
Attribute objEventoPadrao.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPadraoTribEnt As New ClassPadraoTribEnt
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se Todos Produtos não foi preenchido
    If TodosProdutos.Value = 0 Then

        'Verifica se CategoriaProduto foi preenchida
        If Len(Trim(CategoriaProduto.Text)) = 0 Then Error 33514

        'Verifica se o Item da CategoriaProduto foi preenchido
        If Len(Trim(ItemCategoriaProduto.Text)) = 0 Then Error 33515

    Else

        'Verifica se os campos Natureza de Operação ou Tipo de Documento foram preenchidos
        If Len(Trim(NaturezaOp.Text)) = 0 And TipoDocumento.ListIndex = -1 Then Error 33516

    End If

    'Verifica se o Padrão de Tributação existe
    objPadraoTribEnt.sNaturezaOperacao = NaturezaOp.Text
    objPadraoTribEnt.sSiglaMovto = SCodigo_Extrai(TipoDocumento)
    objPadraoTribEnt.sCategoriaProduto = CategoriaProduto.Text
    objPadraoTribEnt.sItemCategoria = ItemCategoriaProduto.Text
    
    lErro = CF("PadraoTribEntrada_Le", objPadraoTribEnt)
    If lErro <> SUCESSO And lErro <> 33483 Then Error 33517

    'Não encontrou o Padrão de Tributação Entrada ==> Erro
    If lErro = 33483 Then Error 33518

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PADRAOTRIBENTRADA", objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

    If vbMsgRes = vbNo Then Exit Sub

    'Exclui o Padrão de Tributação Entrada
    lErro = CF("PadraoTribEntrada_Exclui", objPadraoTribEnt)
    If lErro <> SUCESSO Then Error 33519

    Call Limpa_Tela_PadraoTribEntrada

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err
    
        Case 33514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", Err)

        Case 33515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", Err)

        Case 33516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_PADRAO_TRIBUTACAO_ENTRADA_NAO_PREENCHIDOS", Err)

        Case 33517, 33519

        Case 33518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOTRIBENTRADA_NAO_CADASTRADO", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164070)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 33469

    'Limpa a Tela
    Call Limpa_Tela_PadraoTribEntrada

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 33469

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164071)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 33468

    'Limpa a Tela
    Call Limpa_Tela_PadraoTribEntrada

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 33468

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164072)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CategoriaProduto_Click()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se a CategoriaProduto foi preenchida
    If CategoriaProduto.ListIndex <> -1 Then

        'Preenche o objeto com a Categoria
        objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 40542
        
        ItemCategoriaProduto.Enabled = True
        ItemCategoriaProduto.Clear

        'Preenche ItemCategoriaProduto
        For Each objCategoriaProdutoItem In colCategoria

            ItemCategoriaProduto.AddItem (objCategoriaProdutoItem.sItem)

        Next
        
        TodosProdutos.Value = 0

    Else
        ItemCategoriaProduto.Clear
        ItemCategoriaProduto.Enabled = False
    End If

    Exit Sub

Erro_CategoriaProduto_Click:

    Select Case Err

        Case 40542

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164073)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(CategoriaProduto) <> 0 And CategoriaProduto.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 40540
        
        If lErro <> SUCESSO Then Error 40541
        
    End If
    
    Exit Sub

Erro_CategoriaProduto_Validate:

    Cancel = True


    Select Case Err

        Case 40540
        
        Case 40541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaProduto.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164074)

    End Select

    Exit Sub

    CategoriaProduto_Click
    

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoTipo = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoPadrao = New AdmEvento

    'Lê as Categorias de Produtos
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 33464

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    'Carrega os Tipos de Documento
    lErro = Carrega_TiposDocInfo()
    If lErro <> SUCESSO Then Error 33465

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 33464, 33465

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164075)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega os Tipos de Documentos

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Carrega_TiposDocInfo

    'Chama todos os tipos de documentos contidos na tabela TiposDocInfo
    lErro = CF("TiposDocInfo_Le", colTipoDocInfo)
    If lErro <> SUCESSO Then Error 33466

    'Joga os objetos da coleção na ComboBox TipoDocumento
    For Each objTipoDocInfo In colTipoDocInfo
         TipoDocumento.AddItem objTipoDocInfo.sSigla & SEPARADOR & objTipoDocInfo.sNomeReduzido
         TipoDocumento.ItemData(TipoDocumento.NewIndex) = objTipoDocInfo.iCodigo
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = Err

    Select Case Err

        Case 33466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164076)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoTipo = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoPadrao = Nothing

End Sub

Private Sub ItemCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_ItemCategoriaProduto_Validate

    If Len(ItemCategoriaProduto.Text) <> 0 And ItemCategoriaProduto.ListIndex = -1 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 40543
        
        If lErro <> SUCESSO Then Error 40544
    
    End If

    Exit Sub

Erro_ItemCategoriaProduto_Validate:

    Cancel = True


    Select Case Err

        Case 40543
        
        Case 40544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, ItemCategoriaProduto.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164077)

    End Select

    Exit Sub

End Sub


Private Sub NaturezaLabel_Click()

Dim colSelecao As New Collection
Dim objNaturezaOperacao As New ClassNaturezaOp

    'Chama a tela de browse
    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, objNaturezaOperacao, objEventoNatureza)

End Sub

Private Sub NaturezaOp_Change()

    iNaturezaAlterado = 1
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NaturezaOp_GotFocus()
Dim iNaturezaAux As Integer

    iNaturezaAux = iNaturezaAlterado
    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
    iNaturezaAlterado = iNaturezaAux

End Sub

Private Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOperacao As New ClassNaturezaOp

On Error GoTo Erro_NaturezaOp_Validate

    If iNaturezaAlterado = 1 Then

        'Verifica se NaturezaOP foi preenchida
        If Len(Trim(NaturezaOp.Text)) > 0 Then

            objNaturezaOperacao.sCodigo = NaturezaOp.Text

            'Verifica se a Natureza de Operação existe
            lErro = CF("NaturezaOperacao_Le", objNaturezaOperacao)
            If lErro <> SUCESSO And lErro <> 17958 Then Error 33477

            'Se não encontrou a Natureza de Operação em questão
            If lErro = 17958 Then Error 33478

            'Mostra a descrição na tela
            DescrNatOp.Caption = objNaturezaOperacao.sDescricao

        Else
            DescrNatOp.Caption = ""

        End If

        iNaturezaAlterado = 0

    End If

    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True


    Select Case Err

        Case 33477

        Case 33478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNaturezaOperacao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164078)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNaturezaOperacao As New ClassNaturezaOp

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNaturezaOperacao = obj1

    'Mostra o Código e a Descrição da Natureza da Operação na tela
    NaturezaOp.Text = objNaturezaOperacao.sCodigo
    DescrNatOp.Caption = objNaturezaOperacao.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164079)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPadrao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPadraoTribEnt As ClassPadraoTribEnt

On Error GoTo Erro_objEventoPadrao_evSelecao

    Set objPadraoTribEnt = obj1

    'Mostra o Padrão de Tributação na tela
    lErro = Traz_PadraoTribEntrada_Tela(objPadraoTribEnt)
    If lErro <> SUCESSO Then Error 33467

    Me.Show

    Exit Sub

Erro_objEventoPadrao_evSelecao:

    Select Case Err

        Case 33467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164080)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoTributacao As ClassTipoDeTributacaoMovto

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoTributacao = obj1

    'Mostra o Código e a Descrição do Tipo de Tributação na tela
    TipoTrib.Text = CStr(objTipoTributacao.iTipo)
    DescrTipoTrib.Caption = objTipoTributacao.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164081)

    End Select

    Exit Sub

End Sub

Private Sub PesquisarPadroesTrib_Click()
'Ativar browse p/padrões existentes
'um "padrão" selecionado deve ser trazido p/a tela

Dim colSelecao As New Collection
Dim objPadraoTribEnt As New ClassPadraoTribEnt

    'Chama a tela de browse
    Call Chama_Tela("PadroesTribEntLista", colSelecao, objPadraoTribEnt, objEventoPadrao)

End Sub

Private Sub TipoDocumento_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoLabel_Click()

Dim colSelecao As New Collection
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

    colSelecao.Add "1"
    colSelecao.Add "0"
    
    'Chama a tela de browse
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTributacao, objEventoTipo)

End Sub

Private Function Limpa_Tela_PadraoTribEntrada() As Long
'Limpa os campos tela PadraoTribEntrada

Dim lErro As Long

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    DescrNatOp.Caption = ""
    TipoDocumento.ListIndex = -1
    DescrTipoTrib.Caption = ""

    TodosProdutos.Value = vbChecked
    
    iAlterado = 0

End Function

Private Sub TipoTrib_Change()

    iTipoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTrib_GotFocus()
Dim iTipoAux As Integer
    
    iTipoAux = iTipoAlterado
    Call MaskEdBox_TrataGotFocus(TipoTrib, iAlterado)
    iTipoAlterado = iTipoAux

End Sub

Private Sub TipoTrib_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_TipoTrib_Validate

    If iTipoAlterado = 1 Then

        'Verifica se o Tipo de Tributação foi preenchido
        If Len(Trim(TipoTrib.Text)) > 0 Then

            objTipoTributacao.iTipo = CInt(TipoTrib.Text)

            'Verifica se o Tipo de Tributação existe
            lErro = CF("TipoTributacao_Le", objTipoTributacao)
            If lErro <> SUCESSO And lErro <> 27259 Then Error 33478

            'Se não encontrou ==> erro
            If lErro = 27259 Then Error 33479

            'Mostra a descrição na tela
            DescrTipoTrib.Caption = objTipoTributacao.sDescricao
            
        Else
            'Limpa a descrição
            DescrTipoTrib.Caption = ""
            
        End If

        iTipoAlterado = 0

    End If

    Exit Sub

Erro_TipoTrib_Validate:

    Cancel = True


    Select Case Err

        Case 33478

        Case 33479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objTipoTributacao.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164082)

    End Select

    Exit Sub

End Sub

Private Sub TodosProdutos_Click()

    iAlterado = REGISTRO_ALTERADO
    If TodosProdutos.Value = 1 Then
        CategoriaProduto.ListIndex = -1
    End If

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Padrão De Tributação Entrada necessários foram preenchidos
'e faz a gravação

Dim lErro As Long
Dim objPadraoTribEnt As New ClassPadraoTribEnt

On Error GoTo Erro_Gravar_Registro

    'Verifica se Todos Produtos não foi preenchido
    If TodosProdutos.Value = 0 Then

        'Verifica se CategoriaProduto foi preenchida
        If Len(Trim(CategoriaProduto.Text)) = 0 Then Error 33484

        'Verifica se o Item da CategoriaProduto foi preenchido
        If Len(Trim(ItemCategoriaProduto.Text)) = 0 Then Error 33485

    Else

        'Verifica se os campos Natureza de Operação ou Tipo de Documento foram preenchidos
        If Len(Trim(NaturezaOp.Text)) = 0 And TipoDocumento.ListIndex = -1 Then Error 33486

    End If

    'Verifica se foi Tipo de Tributação foi preenchido
    If Len(Trim(TipoTrib.Text)) = 0 Then Error 33487

    'Lê os dados da Tela relacionados ao Padrão de Tributação Entrada
    lErro = Move_Tela_Memoria(objPadraoTribEnt)
    If lErro <> SUCESSO Then Error 33488

    lErro = Trata_Alteracao(objPadraoTribEnt, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)
    If lErro <> SUCESSO Then Error 32324

    'Grava o Padrão de Tributaçno BD
    lErro = CF("PadraoTribEntrada_Grava", objPadraoTribEnt)
    If lErro <> SUCESSO Then Error 33489

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 32324

        Case 33484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", Err)

        Case 33485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", Err)

        Case 33486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_PADRAO_TRIBUTACAO_ENTRADA_NAO_PREENCHIDOS", Err)

        Case 33487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", Err)

        Case 33488, 33489

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164083)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Move os dados da tela para memória

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica os campos da tela que estão preenchidos
    objPadraoTribEnt.sNaturezaOperacao = NaturezaOp.Text

    If TipoDocumento.ListIndex <> -1 Then
        objPadraoTribEnt.sSiglaMovto = SCodigo_Extrai(TipoDocumento)
    Else
        objPadraoTribEnt.sSiglaMovto = ""
    End If

    objPadraoTribEnt.sCategoriaProduto = CategoriaProduto.Text
    objPadraoTribEnt.sItemCategoria = ItemCategoriaProduto.Text

    If Len(Trim(TipoTrib.Text)) > 0 Then objPadraoTribEnt.iTipoTributacaoPadrao = CInt(TipoTrib.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164084)

    End Select

    Exit Function

End Function

Private Function Traz_PadraoTribEntrada_Tela(objPadraoTribEnt As ClassPadraoTribEnt) As Long
'Traz os dados do Padrão de Tributação Entrada para tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Traz_PadraoTribEntrada_Tela

    'Lê o Padrão de Tributação Entrada
    lErro = CF("PadraoTribEntrada_Le", objPadraoTribEnt)
    If lErro <> SUCESSO And lErro <> 33483 Then Error 33528

    'Não encontrou o Padrão de Tributação Entrada ==> erro
    If lErro = 33483 Then Error 33529

    NaturezaOp.Text = objPadraoTribEnt.sNaturezaOperacao
    Call NaturezaOp_Validate(bSGECancelDummy)

    If objPadraoTribEnt.sSiglaMovto <> "" Then

        objTipoDocInfo.sSigla = objPadraoTribEnt.sSiglaMovto

        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27263 Then Error 33530

        'Se não encontrou ==> erro
        If lErro = 27263 Then Error 33531

        'Percorre todos os elementos da ComboBox
        For iIndice = 0 To TipoDocumento.ListCount - 1

            'Compara se código já existe na ComboBox
            If TipoDocumento.ItemData(iIndice) = objTipoDocInfo.iCodigo Then

                'Seleciona o item na ComboBox
                TipoDocumento.ListIndex = iIndice
                Exit For

            End If

        Next
        
    Else
        TipoDocumento.ListIndex = -1
    End If

    'Verifica se existe CategoriaProduto
    If objPadraoTribEnt.sCategoriaProduto <> "" Then
    
        'Transfere a CategoriaProduto para a tela
        CategoriaProduto.Text = objPadraoTribEnt.sCategoriaProduto
        Call CategoriaProduto_Validate(bSGECancelDummy)
        
        'Transfere a ItemCategoriaProduto para a tela
        ItemCategoriaProduto.Text = objPadraoTribEnt.sItemCategoria
        Call ItemCategoriaProduto_Validate(bSGECancelDummy)
        
    Else
    
        'se nao existir CategoriaProduto
        TodosProdutos.Value = 1
    End If

    TipoTrib.Text = CStr(objPadraoTribEnt.iTipoTributacaoPadrao)
    Call TipoTrib_Validate(bSGECancelDummy)

    iAlterado = 0
    
    Traz_PadraoTribEntrada_Tela = SUCESSO

    Exit Function

Erro_Traz_PadraoTribEntrada_Tela:

    Traz_PadraoTribEntrada_Tela = Err

    Select Case Err

        Case 33538, 33530

        Case 33529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOTRIBENTRADA_NAO_CADASTRADO", Err, objPadraoTribEnt.sNaturezaOperacao, objPadraoTribEnt.sSiglaMovto, objPadraoTribEnt.sCategoriaProduto, objPadraoTribEnt.sItemCategoria)

        Case 33531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", Err, objPadraoTribEnt.sSiglaMovto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164085)

    End Select

    Exit Function

End Function

Private Sub TipoDocumento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Error_TipoDocumento_Validate

    'Verifica se foi preenchida a ComboBox TipoDocumento
    If Len(Trim(TipoDocumento.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoDocumento
    If TipoDocumento.Text = TipoDocumento.List(TipoDocumento.ListIndex) Then Exit Sub

    objTipoDocInfo.sSigla = TipoDocumento.Text

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then Error 41511

    'Se não encontrou ==> erro
    If lErro = 27263 Then Error 41512

    'Percorre todos os elementos da ComboBox
    For iIndice = 0 To TipoDocumento.ListCount - 1

        'Compara se código já existe na ComboBox
        If TipoDocumento.ItemData(iIndice) = objTipoDocInfo.iCodigo Then

            'Seleciona o item na ComboBox
            TipoDocumento.ListIndex = iIndice
            Exit For

        End If

    Next

    Exit Sub

Error_TipoDocumento_Validate:

    Cancel = True


    Select Case Err

        Case 41511

        Case 41512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO1", Err, TipoDocumento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164086)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PADRAO_TRIB_ENTRADA
    Set Form_Load_Ocx = Me
    Caption = "Padrões de Tributação para Operações com Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PadraoTribEntrada"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NaturezaOp Then
            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is TipoTrib Then
            Call TipoLabel_Click
        End If
    
    End If

End Sub


Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoDoc, Source, X, Y)
End Sub

Private Sub LabelTipoDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoDoc, Button, Shift, X, Y)
End Sub

Private Sub DescrNatOp_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrNatOp, Source, X, Y)
End Sub

Private Sub DescrNatOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrNatOp, Button, Shift, X, Y)
End Sub

Private Sub NaturezaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NaturezaLabel, Source, X, Y)
End Sub

Private Sub NaturezaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NaturezaLabel, Button, Shift, X, Y)
End Sub

Private Sub DescrTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrTipoTrib, Source, X, Y)
End Sub

Private Sub DescrTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrTipoTrib, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

