VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PadraoTribSaidaOcx 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6630
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
         Picture         =   "PadraoTribSaidaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PadraoTribSaidaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PadraoTribSaidaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PadraoTribSaidaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   2676
      Left            =   135
      TabIndex        =   12
      Top             =   765
      Width           =   6276
      Begin VB.Frame Frame3 
         Caption         =   "Clientes"
         Height          =   1356
         Left            =   180
         TabIndex        =   13
         Top             =   1170
         Width           =   4464
         Begin VB.CheckBox TodosClientes 
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
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   1212
            TabIndex        =   4
            Top             =   564
            Width           =   2844
         End
         Begin VB.ComboBox ItemCategoriaCliente 
            Height          =   315
            Left            =   1605
            TabIndex        =   5
            Top             =   960
            Width           =   2448
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
            TabIndex        =   15
            Top             =   1020
            Width           =   504
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
            TabIndex        =   14
            Top             =   600
            Width           =   936
         End
      End
      Begin VB.ComboBox TipoDocumento 
         Height          =   315
         Left            =   1452
         TabIndex        =   1
         Top             =   300
         Width           =   4728
      End
      Begin MSMask.MaskEdBox NaturezaOp 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   810
         Width           =   630
         _ExtentX        =   1111
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
         Height          =   195
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label DescrNatOp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   825
         Width           =   4005
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
         Height          =   420
         Left            =   120
         TabIndex        =   16
         Top             =   276
         Width           =   1044
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
      Top             =   210
      Width           =   2088
   End
   Begin MSMask.MaskEdBox TipoTrib 
      Height          =   300
      Left            =   1950
      TabIndex        =   6
      Top             =   3585
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
      TabIndex        =   20
      Top             =   3645
      Width           =   1800
   End
   Begin VB.Label DescrTipoTrib 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2580
      TabIndex        =   19
      Top             =   3600
      Width           =   3810
   End
End
Attribute VB_Name = "PadraoTribSaidaOcx"
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
Dim objPadraoTribSaida As New ClassPadraoTribSaida
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se Todos Clientes não foi preenchido
    If TodosClientes.Value = 0 Then

        'Verifica se CategoriaCliente foi preenchida
        If Len(Trim(CategoriaCliente.Text)) = 0 Then Error 33379

        'Verifica se o Item da CategoriaCliente foi preenchido
        If Len(Trim(ItemCategoriaCliente.Text)) = 0 Then Error 33380

    Else

        'Verifica se os campos Natureza de Operação ou Tipo de Documento foram preenchidos
        If Len(Trim(NaturezaOp.Text)) = 0 And TipoDocumento.ListIndex = -1 Then Error 33381

    End If

    'Verifica se o Padrão de Tributação existe
    objPadraoTribSaida.sNaturezaOperacao = NaturezaOp.Text
    objPadraoTribSaida.sSiglaMovto = SCodigo_Extrai(TipoDocumento)
    objPadraoTribSaida.sCategoriaFilialCliente = CategoriaCliente.Text
    objPadraoTribSaida.sItemCategoria = ItemCategoriaCliente.Text
    lErro = CF("PadraoTribSaida_Le", objPadraoTribSaida)
    If lErro <> SUCESSO And lErro <> 33385 Then Error 33386

    'Não encontrou o Padrão de Tributação Saída ==> Erro
    If lErro = 33385 Then Error 33387

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PADRAO_TRIBUTACAO", objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

    If vbMsgRes = vbNo Then Exit Sub

    'Exclui o Padrão de Tributação Saída
    lErro = CF("PadraoTribSaida_Exclui", objPadraoTribSaida)
    If lErro <> SUCESSO Then Error 33388

    Call Limpa_Tela_PadraoTribSaida

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 33379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA", Err)

        Case 33380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA", Err)

        Case 33381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_PADRAO_TRIBUTACAO_NAO_PREENCHIDOS", Err)

        Case 33386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_TRIBUTACAO_NAO_CADASTRADO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33387, 33388

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164087)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 33337

    'Limpa a Tela
    Call Limpa_Tela_PadraoTribSaida

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 33337

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164088)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaCliente_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se a CategoriaCliente foi preenchida
    If CategoriaCliente.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = CategoriaCliente.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then Error 40547

        ItemCategoriaCliente.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        ItemCategoriaCliente.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            ItemCategoriaCliente.AddItem objCategoriaClienteItem.sItem

        Next
        TodosClientes.Value = 0
    
    Else
        
        ItemCategoriaCliente.Clear
        ItemCategoriaCliente.Enabled = False
    
    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case Err

        Case 40547

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164089)

    End Select

    Exit Sub

End Sub
Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 40545
        
        If lErro <> SUCESSO Then Error 40546
    
    End If

    Exit Sub

Erro_CategoriaCliente_Validate:

    Cancel = True


    Select Case Err

        Case 40545
         
        Case 40546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaCliente.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164090)

    End Select

    Exit Sub

End Sub
Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Form_Load

    Set objEventoTipo = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoPadrao = New AdmEvento

    'Lê as Categorias de Cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then Error 33329

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        CategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next

    'Carrega os Tipos de Documento
    lErro = Carrega_TiposDocInfo()
    If lErro <> SUCESSO Then Error 33377

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 33329, 33377

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164091)

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
    If lErro <> SUCESSO Then Error 33378

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

        Case 33378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164092)

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

Private Sub ItemCategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ItemCategoriaCliente_Validate

    If Len(ItemCategoriaCliente.Text) <> 0 And ItemCategoriaCliente.ListIndex = -1 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 40547
        
        If lErro <> SUCESSO Then Error 40548
    
    End If

    Exit Sub

Erro_ItemCategoriaCliente_Validate:

    Cancel = True


    Select Case Err

        Case 40547
        
        Case 40548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, ItemCategoriaCliente.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164093)

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
            If lErro <> SUCESSO And lErro <> 17958 Then Error 33339

            'Se não encontrou a Natureza de Operação em questão
            If lErro = 17958 Then Error 33340

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

        Case 33339

        Case 33340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNaturezaOperacao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164094)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164095)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPadrao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPadraoTribSaida As ClassPadraoTribSaida

On Error GoTo Erro_objEventoPadrao_evSelecao

    Set objPadraoTribSaida = obj1

    'Mostra o Padrão de Tributação na tela
    lErro = Traz_PadraoTribSaida_Tela(objPadraoTribSaida)
    If lErro <> SUCESSO Then Error 33341

    Me.Show

    Exit Sub

Erro_objEventoPadrao_evSelecao:

    Select Case Err

        Case 33341

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164096)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164097)

    End Select

    Exit Sub

End Sub

Private Sub PesquisarPadroesTrib_Click()
'ativar browse p/padroes existentes
'um "padrao" selecionado deve ser trazido p/a tela

Dim colSelecao As New Collection
Dim objPadraoTribSaida As New ClassPadraoTribSaida

    'Chama a tela de browse
    Call Chama_Tela("PadroesTribSaidaLista", colSelecao, objPadraoTribSaida, objEventoPadrao)

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 33338

    'Limpa a Tela
    Call Limpa_Tela_PadraoTribSaida

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 33338

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164098)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Padrão De Tributação necessários foram preenchidos
'e faz a gravação

Dim lErro As Long
Dim objPadraoTribSaida As New ClassPadraoTribSaida

On Error GoTo Erro_Gravar_Registro

    'Verifica se Todos Clientes não foi preenchido
    If TodosClientes.Value = 0 Then

        'Verifica se CategoriaCliente foi preenchida
        If Len(Trim(CategoriaCliente.Text)) = 0 Then Error 33342

        'Verifica se o Item da CategoriaCliente foi preenchido
        If Len(Trim(ItemCategoriaCliente.Text)) = 0 Then Error 33343

    Else

        'Verifica se os campos Natureza de Operação ou Tipo de Documento foram preenchidos
        If Len(Trim(NaturezaOp.Text)) = 0 And Len(Trim(TipoDocumento.Text)) = 0 Then Error 33344

    End If

    'Verifica se Tipo de Tributação foi preenchido
    If Len(Trim(TipoTrib.Text)) = 0 Then Error 33345

    'Lê os dados da Tela relacionados ao Padrão de Tributação Saída
    lErro = Move_Tela_Memoria(objPadraoTribSaida)
    If lErro <> SUCESSO Then Error 33346

    lErro = Trata_Alteracao(objPadraoTribSaida, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)
    If lErro <> SUCESSO Then Error 32332

    'Grava o Padrão de Tributação Saída no BD
    lErro = CF("PadraoTribSaida_Grava", objPadraoTribSaida)
    If lErro <> SUCESSO Then Error 33347

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 32332

        Case 33342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA", Err)

        Case 33343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA", Err)

        Case 33344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_PADRAO_TRIBUTACAO_NAO_PREENCHIDOS", Err)

        Case 33345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", Err)

        Case 33346, 33347

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164099)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Move os dados da tela para memória

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objPadraoTribSaida.sNaturezaOperacao = NaturezaOp.Text

    If Len(Trim(TipoDocumento.Text)) > 0 Then
        objPadraoTribSaida.sSiglaMovto = SCodigo_Extrai(TipoDocumento)
    Else
        objPadraoTribSaida.sSiglaMovto = ""
    End If

    objPadraoTribSaida.sCategoriaFilialCliente = CategoriaCliente.Text
    objPadraoTribSaida.sItemCategoria = ItemCategoriaCliente.Text

    If Len(Trim(TipoTrib.Text)) > 0 Then
        objPadraoTribSaida.iTipoTributacaoPadrao = CInt(TipoTrib.Text)
    End If

   Move_Tela_Memoria = SUCESSO

   Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164100)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_PadraoTribSaida() As Long
'Limpa os campos tela PadraoDeTributacao

Dim lErro As Long

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    DescrNatOp.Caption = ""
    TipoDocumento.ListIndex = -1
    DescrTipoTrib.Caption = ""
    
    TodosClientes.Value = vbChecked
     
    iAlterado = 0

End Function

Private Sub TipoDocumento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDocumento_Click()

    iAlterado = REGISTRO_ALTERADO
    Call TipoDocumento_Validate(bSGECancelDummy)

End Sub

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
    If lErro <> SUCESSO And lErro <> 27263 Then Error 33333

    'Se não encontrou ==> erro
    If lErro = 27263 Then Error 33334

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

        Case 33333

        Case 33334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO1", Err, TipoDocumento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164101)

    End Select

    Exit Sub

End Sub

Private Sub TipoLabel_Click()

Dim colSelecao As New Collection
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

    colSelecao.Add "1"
    colSelecao.Add "0"
    
    'Chama a tela de browse
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTributacao, objEventoTipo)

End Sub

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
            If lErro <> SUCESSO And lErro <> 27259 Then Error 33335

            'Se não encontrou ==> erro
            If lErro = 27259 Then Error 33336

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

        Case 33335

        Case 33336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objTipoTributacao.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164102)

    End Select

    Exit Sub

End Sub

Private Function Traz_PadraoTribSaida_Tela(objPadraoTribSaida As ClassPadraoTribSaida) As Long
'Traz os dados do Padrão de Tributação Saída para tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Traz_PadraoTribSaida_Tela

    'Lê o Padrão de Tributação Saída
    lErro = CF("PadraoTribSaida_Le", objPadraoTribSaida)
    If lErro <> SUCESSO And lErro <> 33385 Then Error 33402

    'Não encontrou o Padrão de Tributação Saída ==> erro
    If lErro = 33385 Then Error 33403

    NaturezaOp.Text = objPadraoTribSaida.sNaturezaOperacao
    Call NaturezaOp_Validate(bSGECancelDummy)

    If objPadraoTribSaida.sSiglaMovto <> "" Then

        objTipoDocInfo.sSigla = objPadraoTribSaida.sSiglaMovto

        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27263 Then Error 33408

        'Se não encontrou ==> erro
        If lErro = 27263 Then Error 33409

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
        TipoDocumento.Text = ""
    End If

    If objPadraoTribSaida.sCategoriaFilialCliente <> "" Then
    
        CategoriaCliente.Text = objPadraoTribSaida.sCategoriaFilialCliente
        Call CategoriaCliente_Validate(bSGECancelDummy)
       
        ItemCategoriaCliente.Text = objPadraoTribSaida.sItemCategoria
        Call ItemCategoriaCliente_Validate(bSGECancelDummy)
        
    Else
        TodosClientes.Value = 1
    End If

    TipoTrib.Text = CStr(objPadraoTribSaida.iTipoTributacaoPadrao)
    Call TipoTrib_Validate(bSGECancelDummy)

    iAlterado = 0
    
    Traz_PadraoTribSaida_Tela = SUCESSO

    Exit Function

Erro_Traz_PadraoTribSaida_Tela:

    Traz_PadraoTribSaida_Tela = Err

    Select Case Err

        Case 33402, 33408

        Case 33403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_TRIBUTACAO_NAO_CADASTRADO", Err, objPadraoTribSaida.sNaturezaOperacao, objPadraoTribSaida.sSiglaMovto, objPadraoTribSaida.sCategoriaFilialCliente, objPadraoTribSaida.sItemCategoria)

        Case 33409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", Err, objPadraoTribSaida.sSiglaMovto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164103)

    End Select

    Exit Function

End Function

Private Sub TodosClientes_Click()

    iAlterado = REGISTRO_ALTERADO
    If TodosClientes.Value = 1 Then
        CategoriaCliente.ListIndex = -1
    End If
    
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PADRAO_TRIB_SAIDA
    Set Form_Load_Ocx = Me
    Caption = "Padrões de Tributação para Operações com Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PadraoTribSaida"
    
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





Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub NaturezaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NaturezaLabel, Source, X, Y)
End Sub

Private Sub NaturezaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NaturezaLabel, Button, Shift, X, Y)
End Sub

Private Sub DescrNatOp_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrNatOp, Source, X, Y)
End Sub

Private Sub DescrNatOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrNatOp, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoDoc, Source, X, Y)
End Sub

Private Sub LabelTipoDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoDoc, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub DescrTipoTrib_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrTipoTrib, Source, X, Y)
End Sub

Private Sub DescrTipoTrib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrTipoTrib, Button, Shift, X, Y)
End Sub

