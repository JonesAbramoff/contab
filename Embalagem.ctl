VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Embalagem 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9135
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4155
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   810
      Width           =   8685
      Begin VB.ListBox Embalagens 
         Height          =   2310
         IntegralHeight  =   0   'False
         Left            =   5805
         TabIndex        =   25
         Top             =   1710
         Width           =   2760
      End
      Begin VB.Frame FrameDetalhamento 
         Caption         =   "Detalhamento"
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   1485
         Width           =   5535
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   390
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Sigla 
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   930
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Capacidade 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Top             =   1995
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Peso 
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            Top             =   1440
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00####"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Kg"
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
            Left            =   2760
            TabIndex        =   24
            Top             =   1500
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "l"
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
            Left            =   2760
            TabIndex        =   23
            Top             =   2055
            Width           =   60
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   720
            TabIndex        =   22
            Top             =   1500
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade:"
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
            Left            =   135
            TabIndex        =   21
            Top             =   2055
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Sigla:"
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
            Left            =   720
            TabIndex        =   20
            Top             =   990
            Width           =   495
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
            Left            =   285
            TabIndex        =   19
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Frame FrameIdentificacao 
         Caption         =   "Identificação"
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   8475
         Begin VB.CommandButton BotaoProxNum 
            Height          =   300
            Left            =   1650
            Picture         =   "Embalagem.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   960
            TabIndex        =   9
            Top             =   330
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   960
            TabIndex        =   10
            Top             =   750
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
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
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   660
         End
         Begin VB.Label ProdutoLabel1 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   12
            Top             =   825
            Width           =   735
         End
         Begin VB.Label DescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2445
            TabIndex        =   11
            Top             =   750
            Width           =   5820
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Embalagens"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   1440
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4095
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   810
      Visible         =   0   'False
      Width           =   8685
      Begin VB.Frame Frame2 
         Caption         =   "Expedição / Rótulos"
         Height          =   3915
         Left            =   90
         TabIndex        =   28
         Top             =   120
         Width           =   8505
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "Produtos"
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
            Left            =   6540
            TabIndex        =   33
            Top             =   3450
            Width           =   1680
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   255
            Left            =   5580
            TabIndex        =   32
            Top             =   1980
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoItem 
            Height          =   255
            Left            =   4470
            TabIndex        =   31
            Top             =   960
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   4500
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1440
            Width           =   3195
         End
         Begin MSFlexGridLib.MSFlexGrid GridExpedicao 
            Height          =   2925
            Left            =   330
            TabIndex        =   29
            Top             =   390
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   5159
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6825
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Embalagem.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Embalagem.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Embalagem.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Embalagem.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4545
      Left            =   150
      TabIndex        =   5
      Top             =   480
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   8017
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Expedição"
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
Attribute VB_Name = "Embalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoProdutoItem As AdmEvento
Attribute objEventoProdutoItem.VB_VarHelpID = -1

'por leo
Dim objGridExpedicao As AdmGrid
Dim iGrid_ProdutoItem_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_Quantidade_Col As Integer

'??? ERRO_PRODUTO_NAO_EMBALAGEM = O Produto %s não é é de Natureza igual a Embalagem. -> sCodEmbalagem

'Property Variables:
Dim m_Caption As String
Event Unload()

'por leo
Dim iFrameAtual As Integer

Function Trata_Parametros(Optional objEmbalagem As ClassEmbalagem) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma Embalagem selecionada
    If Not (objEmbalagem Is Nothing) Then

        'Verifica se a Embalagem existe, lendo no BD a partir do Código
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 80297

        'Se a Embalagem existe
        If lErro = SUCESSO Then

            lErro = Traz_Embalagem_Tela(objEmbalagem)
            If lErro <> SUCESSO Then gError 80298

        'Se a Embalagem não existe
        Else

            Call Limpa_Tela(Me)

            If objEmbalagem.iCodigo > 0 Then

                'Mantém o código da Embalagem na tela
                Codigo.Text = CStr(objEmbalagem.iCodigo)

            End If

        End If

    End If

    'Zerar iAlterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 80297, 80298

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159365)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento

    'por leo
    Set objEventoProdutoItem = New AdmEvento
    Set objGridExpedicao = New AdmGrid

    'por leo
    iFrameAtual = 1

    'Carrega a list de Embalagens
    lErro = Embalagens_Carrega()
    If lErro <> SUCESSO Then gError 82725

    '??? por leo
    lErro = Inicializa_GridExpedicao(objGridExpedicao)
    If lErro <> SUCESSO Then gError 103428

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 103104

    'por leo
    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoItem)
    If lErro <> SUCESSO Then gError 103422

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 82725, 103104, 103422, 103428

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159366)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Traz_Embalagem_Tela(objEmbalagem As ClassEmbalagem) As Long
'Traz os dados da embalagem para a tela

Dim lErro As Long
Dim sProdutoMascarado As String
Dim objEmbalagensExpedicao As ClassEmbalagensExpedicao
Dim iIndex As Integer
Dim sProdutoItemMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_Embalagem_Tela

    'Limpa a tela
    Call Limpa_Tela_Embalagem

    'Lê a Embalagem
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 82763 Then gError 82775
    'se não encontrou a embalagem ==> erro
    If lErro = 82763 Then gError 82776

    'Traz os dados para a tela
    Codigo.Text = objEmbalagem.iCodigo
    Sigla.Text = objEmbalagem.sSigla
    Descricao.Text = objEmbalagem.sDescricao
    Capacidade.Text = objEmbalagem.dCapacidade
    Peso.Text = objEmbalagem.dPeso

    If Len(Trim(objEmbalagem.sProduto)) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objEmbalagem.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 103108

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        Call Produto_Validate(bSGECancelDummy)

    End If

    'Por Leo
    For Each objEmbalagensExpedicao In objEmbalagem.colEmbExpedicao

        objProduto.sCodigo = objEmbalagensExpedicao.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 103445

        If lErro = SUCESSO Then

            iIndex = iIndex + 1

            lErro = Mascara_MascararProduto(objEmbalagensExpedicao.sProduto, sProdutoItemMascarado)
            If lErro <> SUCESSO Then gError 103446

            GridExpedicao.TextMatrix(iIndex, iGrid_ProdutoItem_Col) = sProdutoItemMascarado
            GridExpedicao.TextMatrix(iIndex, iGrid_DescricaoItem_Col) = objProduto.sDescricao
            GridExpedicao.TextMatrix(iIndex, iGrid_Quantidade_Col) = objEmbalagensExpedicao.dQuantidade

         End If

    Next

    objGridExpedicao.iLinhasExistentes = iIndex
    'Por Leo até aqui

    iAlterado = 0

    Traz_Embalagem_Tela = SUCESSO

    Exit Function

Erro_Traz_Embalagem_Tela:

    Traz_Embalagem_Tela = gErr

    Select Case gErr

        Case 82775
            'Erro tratado na rotina chamada

        Case 82776
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_CADASTRADA", gErr, objEmbalagem.iCodigo)

        Case 103108

        Case 103445, 103446

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159367)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objEmbalagem As ClassEmbalagem) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndex As Integer
Dim objEmbalagensExpedicao As ClassEmbalagensExpedicao
Dim sProdutoItemFormatado As String
Dim iProdutoItemPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objEmbalagem.iCodigo = StrParaInt(Codigo.Text)
    objEmbalagem.dPeso = StrParaDbl(Peso.Text)
    objEmbalagem.dCapacidade = StrParaDbl(Capacidade.Text)

    If Len(Trim(Descricao.Text)) > 0 Then
        objEmbalagem.sDescricao = Descricao.Text
    End If

    If Len(Trim(Sigla.Text)) > 0 Then
        objEmbalagem.sSigla = Sigla.Text
    End If

    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 103107

    objEmbalagem.sProduto = sProdutoFormatado

   'por Leo

    For iIndex = 1 To objGridExpedicao.iLinhasExistentes

        Set objEmbalagensExpedicao = New ClassEmbalagensExpedicao

        'Formata o produto
        lErro = CF("Produto_Formata", GridExpedicao.TextMatrix(iIndex, iGrid_ProdutoItem_Col), sProdutoItemFormatado, iProdutoItemPreenchido)
        If lErro <> SUCESSO Then gError 103429

        objEmbalagensExpedicao.sProduto = sProdutoItemFormatado
        objEmbalagensExpedicao.dQuantidade = StrParaDbl(GridExpedicao.TextMatrix(iIndex, iGrid_Quantidade_Col))
        objEmbalagensExpedicao.iEmbalagem = objEmbalagem.iCodigo
        objEmbalagensExpedicao.iSequencial = iIndex

        objEmbalagem.colEmbExpedicao.Add objEmbalagensExpedicao

    Next

   'por leo até aqui

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 103107, 103429

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159368)

    End Select

    Exit Function

End Function

Private Function Embalagens_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem
Dim colEmbalagem As New Collection

On Error GoTo Erro_Embalagens_Carrega

    'Le todas as Embalagens
    lErro = CF("Embalagem_Le_Todas", colEmbalagem)
    If lErro <> SUCESSO And lErro <> 82731 Then gError 82726

    For Each objEmbalagem In colEmbalagem
        Embalagens.AddItem objEmbalagem.iCodigo & SEPARADOR & objEmbalagem.sSigla
        Embalagens.ItemData(Embalagens.NewIndex) = objEmbalagem.iCodigo
    Next

    Embalagens_Carrega = SUCESSO

    Exit Function

Erro_Embalagens_Carrega:

    Embalagens_Carrega = gErr

    Select Case gErr

        Case 82726
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159369)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    'por leo
    Set objGridExpedicao = Nothing
    Set objEventoProdutoItem = Nothing

    Set objEventoProduto = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Embalagens"

    'Le os dados da Tela
    lErro = Move_Tela_Memoria(objEmbalagem)
    If lErro <> SUCESSO Then gError 82723

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objEmbalagem.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objEmbalagem.sDescricao, STRING_BUFFER_MAX_TEXTO, "Descricao"
    colCampoValor.Add "Sigla", objEmbalagem.sSigla, STRING_BUFFER_MAX_TEXTO, "Sigla"
    colCampoValor.Add "Capacidade", objEmbalagem.dCapacidade, 0, "Capacidade"
    colCampoValor.Add "Peso", objEmbalagem.dPeso, 0, "Peso"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 82723
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159370)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Tela_Preenche

    objEmbalagem.iCodigo = colCampoValor.Item("Codigo").vValor
    objEmbalagem.dCapacidade = colCampoValor.Item("Capacidade").vValor
    objEmbalagem.dPeso = colCampoValor.Item("Peso").vValor
    objEmbalagem.sDescricao = colCampoValor.Item("Descricao").vValor
    objEmbalagem.sSigla = colCampoValor.Item("Sigla").vValor

    'Traz dados da Embalagem para a Tela
    lErro = Traz_Embalagem_Tela(objEmbalagem)
    If lErro <> SUCESSO Then gError 82724

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 82724
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159371)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Embalagem"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Embalagem"

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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 82756

    objEmbalagem.iCodigo = StrParaInt(Codigo.Text)

    'Lê a embalagem
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 82763 Then gError 82757

    'Se não encontrou a Embalagem ==> erro
    If lErro = 82763 Then gError 82758

    'Envia aviso perguntando se realmente deseja excluir Embalagem
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_EMBALAGEM", objEmbalagem.iCodigo)

    If vbMsgRes = vbYes Then

        'Recolhe os dados da tela
        lErro = Move_Tela_Memoria(objEmbalagem)
        If lErro <> SUCESSO Then gError 82764

        'Exclui Embalagem
        '''04/09/01 - Marcelo alteração da chamada da função
        lErro = CF("Embalagem_Exclui", objEmbalagem)
        If lErro <> SUCESSO Then gError 82759

        'Exclui da ListBox
        Call Exclui_Lista(objEmbalagem)

        'Limpa a Tela
        Call Limpa_Tela_Embalagem

    End If

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 82756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 82757, 82759
            'Erro tratado na rotina chamada

        Case 82758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_CADASTRADA", gErr, objEmbalagem.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159372)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Grava a embalagem
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 82737

    'limpa a tela
    Call Limpa_Tela_Embalagem

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGravar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 82737
            'erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159373)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

   'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 82736

    'Limpa a Tela
    Call Limpa_Tela_Embalagem

    'Fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 82736
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159374)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Gravar_Registro

    'Verifica se o código está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 82738

    'Verifica se o Produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 103113

    'Verifica se a Descricao está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 82739

    'verifica se a sigla está preenchida
    If Len(Trim(Sigla.Text)) = 0 Then gError 82740

    'verifica se a capacidade está preenchida
    If Len(Trim(Capacidade.Text)) = 0 Then gError 82741

    'verifica se o peso está preenchido
    If Len(Trim(Peso.Text)) = 0 Then gError 82742

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objEmbalagem)
    If lErro <> SUCESSO Then gError 82743

    lErro = Trata_Alteracao(objEmbalagem, objEmbalagem.iCodigo)
    If lErro <> SUCESSO Then Error 32310

    'Grava a Embalagem no BD
    '''04/09/01 - Marcelo alteração da chamada da função
    lErro = CF("Embalagem_Grava", objEmbalagem)
    If lErro <> SUCESSO Then gError 82744

    'Exclui a embalagem da lista
    Call Exclui_Lista(objEmbalagem)

    'Adiciona na listbox se necessário
    Call Adiciona_Lista(objEmbalagem)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32310

        Case 82738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 82739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 82740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_NAO_PREENCHIDA", gErr)

        Case 82741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAPACIDADE_NAO_PREENCHIDA", gErr)

        Case 82742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESO_NAO_PREENCHIDO", gErr)

        Case 82743, 82744
            'Erros tratados nas rotinas chamadas

        Case 103113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159375)

    End Select

    Exit Function

End Function

Private Sub Exclui_Lista(objEmbalagem As ClassEmbalagem)
'Remove a embalagem da lista

Dim iIndice As Integer

    For iIndice = 0 To Embalagens.ListCount - 1
        If Embalagens.ItemData(iIndice) = objEmbalagem.iCodigo Then
            Embalagens.RemoveItem iIndice
            Exit For
        End If
    Next

End Sub

Private Sub Adiciona_Lista(objEmbalagem As ClassEmbalagem)
'Adiciona a embalagam na lista

Dim iIndice As Integer

    For iIndice = 0 To Embalagens.ListCount - 1
        If Embalagens.ItemData(iIndice) > objEmbalagem.iCodigo Then Exit For
    Next

    Embalagens.AddItem objEmbalagem.iCodigo & SEPARADOR & objEmbalagem.sSigla, iIndice
    Embalagens.ItemData(iIndice) = objEmbalagem.iCodigo

End Sub

Private Function Limpa_Tela_Embalagem() As Long
'Limpa os campos tela Embalagem

    'Função que limpa campos da tela
    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridExpedicao)

    DescProd.Caption = ""

    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridExpedicao.Row = 0 Then gError 43718

    'Verifica se o Produto está preenchido
    If Len(Trim(GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_ProdutoItem_Col))) > 0 Then

        lErro = CF("Produto_Formata", GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_ProdutoItem_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 55325

        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoItem)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 43718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 55325

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159376)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o próximo código disponível para embalagem
    lErro = Embalagem_Automatico(iCodigo)
    If lErro <> SUCESSO Then gError 82734

    'Coloca o código na tela
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 82734
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159377)

    End Select

    Exit Sub

End Sub

Private Sub Capacidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Capacidade_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dCapacidade As Double

On Error GoTo Erro_Capacidade_Validate

    'Verifica se  esta preenchido
    If Len(Trim(Capacidade.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_Positivo_Critica(Capacidade.Text)
        If lErro <> SUCESSO Then gError 82777

        dCapacidade = StrParaDbl(Capacidade.Text)

        'Coloca o valor no formato standard da tela
        Capacidade.Text = Format(dCapacidade, "Standard")

    End If

    Exit Sub

Erro_Capacidade_Validate:

    Cancel = True

    Select Case gErr

        Case 82777
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159378)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se o codigo está preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    'Critica o código informado
    lErro = Valor_Positivo_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 82733

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 82733
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159379)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Embalagens_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Embalagens_DblClick()

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Embalagens_DblClick

    objEmbalagem.iCodigo = Codigo_Extrai(Embalagens.List(Embalagens.ListIndex))

    'Traz os dados da embalagem para a tela
    lErro = Traz_Embalagem_Tela(objEmbalagem)
    If lErro <> SUCESSO Then gError 82732

    'Fecha o sistema de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Embalagens_DblClick:

    Select Case gErr

        Case 82732
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159380)

    End Select

    Exit Sub

End Sub


Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103109

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103110

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, DescProd)
    If lErro <> SUCESSO Then gError 103111

    'Se o tipo de produto não estiver cadastrado como embalagem
    If objProduto.iNatureza <> NATUREZA_PROD_EMBALAGENS Then gError 103112

    If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 103123

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 103109, 103111

        Case 103110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 103112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_EMBALAGEM", gErr, objProduto.sCodigo)

        Case 103123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159381)

    End Select

    Exit Sub

End Sub

Private Sub Peso_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Peso_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPeso As Double

On Error GoTo Erro_Peso_Validate

    'Verifica se  esta preenchido
    If Len(Trim(Peso.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(Peso.Text)
        If lErro <> SUCESSO Then gError 82778

        dPeso = StrParaDbl(Peso.Text)

        'Coloca o valor no formato standard da tela
        Peso.Text = Format(dPeso, "#,##0.00####")

    End If

    Exit Sub

Erro_Peso_Validate:

    Cancel = True

    Select Case gErr

        Case 82778
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159382)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate

    lErro = ProdutoEmbalagem_Perde_Foco(Produto, DescProd)
    If lErro <> SUCESSO And lErro <> 103117 Then gError 103105

    'se a descrição da embalagem ainda não estiver preenchida, sugere a descrição do produto p/ a embalgem
    If Len(Trim(Descricao.Text)) = 0 Then

        Descricao.Text = DescProd.Caption

    End If

    If lErro <> SUCESSO Then gError 103106

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 103105

        Case 103106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159383)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel1_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_ProdutoLabel1_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103108

        objProduto.sCodigo = sProdutoFormatado

    End If

    colSelecao.Add CInt(NATUREZA_PROD_EMBALAGENS)
    colSelecao.Add CInt(PRODUTO_GERENCIAL)

    sSelecao = "Natureza=? AND Gerencial<>?"

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_ProdutoLabel1_Click:

    Select Case gErr

        Case 103108

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159384)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridExpedicao)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExpedicao)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExpedicao.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridExpedicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_Click()
'??? por leo

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    'Se o usuário está tentando gerar um próximo nº automático
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click

    'Se o usuário está tentando chamar um browser
    ElseIf KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel1_Click
        End If

    End If


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

'???Subir para ClassESTGrava
Function Embalagem_Automatico(iCodigo As Integer) As Long
'retorna o número do proximo código de Embalagem disponível

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Embalagem_Automatico

    lErro = CF("Config_Obter_Inteiro_Automatico", "EstConfig", "NUM_PROX_COD_EMBALAGEM", "Embalagens", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 82735

    Embalagem_Automatico = SUCESSO

    Exit Function

Erro_Embalagem_Automatico:

    Embalagem_Automatico = gErr

    Select Case gErr

        Case 82735
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159385)

    End Select

    Exit Function

End Function

'Inserida por Leo
Private Function ProdutoEmbalagem_Perde_Foco(ByVal Produto As Object, ByVal Desc As Object) As Long
'recebe MaskEdBox do Produto e o label da descrição

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoEmbalagem_Perde_Foco

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", Produto.Text, sProdFormatado, iProdPreenchido)
    If lErro Then gError 103115

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 103116

        'Se o produto não existe ou o produto não é de natureza embalagem ou não é analítico.
        If lErro = 28030 Or objProduto.iNatureza <> NATUREZA_PROD_EMBALAGENS Or objProduto.iGerencial = PRODUTO_GERENCIAL Then

            Produto.PromptInclude = False
            Produto.Text = ""
            Produto.PromptInclude = True
            If Produto.Visible Then Produto.SetFocus
            Desc.Caption = ""

            'se o produto não existe -> Erro.
            If lErro = 28030 Then

                gError 103117

            'se o produto não é de natureza embalagem -> Erro
            ElseIf objProduto.iNatureza <> NATUREZA_PROD_EMBALAGENS Then

                gError 103114

            'Se o produto for gerencial -> Erro.
            Else

                gError 103124

            End If

        End If

        Desc.Caption = objProduto.sDescricao

    Else

        Desc.Caption = ""

    End If

    ProdutoEmbalagem_Perde_Foco = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Perde_Foco:

    ProdutoEmbalagem_Perde_Foco = gErr

    Select Case gErr

        Case 103114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_EMBALAGEM", gErr, objProduto.sCodigo)

        Case 103115, 103116, 103117

        Case 103124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159386)

    End Select

    Exit Function

End Function

'por leo
Private Function Inicializa_GridExpedicao(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoItem.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (Quantidade.Name)

    'Colunas do Grid
    iGrid_ProdutoItem_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_Quantidade_Col = 3

    objGrid.objGrid = GridExpedicao

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 15

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridExpedicao.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridExpedicao = SUCESSO

End Function

Private Sub GridExpedicao_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridExpedicao, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridExpedicao, iAlterado)
        End If

End Sub

Private Sub GridExpedicao_GotFocus()
    Call Grid_Recebe_Foco(objGridExpedicao)
End Sub

Private Sub GridExpedicao_EnterCell()

    Call Grid_Entrada_Celula(objGridExpedicao, iAlterado)

End Sub

Private Sub GridExpedicao_LeaveCell()
    Call Saida_Celula(objGridExpedicao)
End Sub

Private Sub GridExpedicao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridExpedicao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExpedicao, iAlterado)
    End If

End Sub

Private Sub GridExpedicao_RowColChange()

    Call Grid_RowColChange(objGridExpedicao)

End Sub

Private Sub GridExpedicao_Scroll()

    Call Grid_Scroll(objGridExpedicao)

End Sub

Private Sub GridExpedicao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridExpedicao)

End Sub

Private Sub GridExpedicao_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridExpedicao)

End Sub

Private Sub ProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridExpedicao)

End Sub

Private Sub ProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExpedicao)

End Sub

Private Sub ProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExpedicao.objControle = ProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridExpedicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridExpedicao)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExpedicao)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExpedicao.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGridExpedicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        Select Case GridExpedicao.Col

            Case iGrid_ProdutoItem_Col

                lErro = Saida_Celula_ProdutoItem(objGridInt)
                If lErro <> SUCESSO Then gError 103410

            Case iGrid_Quantidade_Col

                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 103411

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 103412

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 103410, 103411

        Case 103412
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159387)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoItem(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente`

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_ProdutoItem

    Set objGridInt.objControle = ProdutoItem

    lErro = CF("Produto_Formata", ProdutoItem.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 103413

    'se o produto foi preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = VerificaUso_Produto(sProdutoFormatado)
        If lErro <> SUCESSO Then gError 103414

        objProduto.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 103415

        If lErro <> SUCESSO Then gError 103416

        GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridExpedicao.Row - GridExpedicao.FixedRows) = objGridExpedicao.iLinhasExistentes Then
            objGridExpedicao.iLinhasExistentes = objGridExpedicao.iLinhasExistentes + 1
        End If

    Else

        GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_DescricaoItem_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 103417

    Saida_Celula_ProdutoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoItem:

    Saida_Celula_ProdutoItem = gErr

    Select Case gErr

        Case 103413, 103414, 103415, 103417
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 103416
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159388)

    End Select

    Exit Function

End Function

Private Function VerificaUso_Produto(sCodigo As String) As Long
'Verifica se existem produtos repetidos no grid

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer


On Error GoTo Erro_VerificaUso_Produto

    If objGridExpedicao.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGridExpedicao.iLinhasExistentes

            If GridExpedicao.Row <> iIndice Then

                lErro = CF("Produto_Formata", GridExpedicao.TextMatrix(iIndice, iGrid_ProdutoItem_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 103418

                If sProdutoFormatado = sCodigo Then gError 103419

            End If

        Next

    End If

    VerificaUso_Produto = SUCESSO

    Exit Function

Erro_VerificaUso_Produto:

    VerificaUso_Produto = gErr

    Select Case gErr

        Case 103418

        Case 103419
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO_GRID", gErr, sCodigo, iIndice)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159389)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 103420

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 103420
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159390)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_ProdutoItem_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 103421

    'Pesquisa o controle da coluna em questão
'    Select Case objControl.Name

        'Produto
'        Case Produto.Name

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                ProdutoItem.Enabled = False
                Quantidade.Enabled = True
            Else
                ProdutoItem.Enabled = True
                Quantidade.Enabled = False
            End If

 '   End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 103421

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159391)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoItem_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoItem_evSelecao

    Set objProduto = obj1

    If GridExpedicao.Row <> 0 Then

        lErro = CF("Produto_Formata", GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_ProdutoItem_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103423

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            'verifica sua existência na OP
            lErro = VerificaUso_Produto(objProduto.sCodigo)
            If lErro <> SUCESSO Then gError 103424

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 103426

            If lErro = 28030 Then gError 103427

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 103425

            ProdutoItem.PromptInclude = False
            ProdutoItem.Text = sProdutoMascarado
            ProdutoItem.PromptInclude = True

            If (GridExpedicao.Row - GridExpedicao.FixedRows) = objGridExpedicao.iLinhasExistentes Then
                objGridExpedicao.iLinhasExistentes = objGridExpedicao.iLinhasExistentes + 1
            End If

            If Not (Me.ActiveControl Is Produto) Then

                'preenche produto
                GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_ProdutoItem_Col) = sProdutoMascarado
                GridExpedicao.TextMatrix(GridExpedicao.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoItem_evSelecao:

    Select Case gErr

        Case 103423, 103426, 103424

        Case 103425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 103427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159392)

    End Select

    Exit Sub

End Sub

