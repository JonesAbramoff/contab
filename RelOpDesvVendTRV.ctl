VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpDesvVendTRV 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   7785
   Begin VB.Frame Frame6 
      Caption         =   "Outros Filtros"
      Height          =   945
      Left            =   180
      TabIndex        =   38
      Top             =   4935
      Width           =   7440
      Begin VB.ComboBox Responsavel 
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Top             =   210
         Width           =   2280
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   1515
         TabIndex        =   16
         Top             =   210
         Width           =   2205
      End
      Begin VB.ComboBox UsuRespCallCenter 
         Height          =   315
         Left            =   1515
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   570
         Width           =   2205
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   5040
         TabIndex        =   19
         Top             =   570
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
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
         Left            =   3840
         TabIndex        =   43
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Promotor:"
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
         Left            =   4170
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Região:"
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
         TabIndex        =   40
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Re. Call Center:"
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
         Left            =   105
         TabIndex        =   39
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exibição"
      Height          =   2310
      Left            =   180
      TabIndex        =   34
      Top             =   2595
      Width           =   7440
      Begin VB.Frame Frame7 
         Caption         =   "Ordenação"
         Height          =   435
         Left            =   90
         TabIndex        =   41
         Top             =   1815
         Width           =   7230
         Begin VB.OptionButton OptOrdAlfa 
            Caption         =   "Alfabética"
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
            Left            =   3750
            TabIndex        =   15
            Top             =   150
            Width           =   1395
         End
         Begin VB.OptionButton OptOrdRelevancia 
            Caption         =   "Relevância"
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
            Left            =   1395
            TabIndex        =   14
            Top             =   150
            Value           =   -1  'True
            Width           =   1395
         End
      End
      Begin VB.CheckBox TrazerCliNComp 
         Caption         =   "Trazer os clientes que não fizeram nenhuma compra no mês\ano selecionados mesmo que não atendam aos demais critérios de seleção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   135
         TabIndex        =   13
         Top             =   1365
         Width           =   7170
      End
      Begin VB.Frame Frame5 
         Caption         =   "Desconsiderar clientes que não compraram pelo menos"
         Height          =   630
         Left            =   105
         TabIndex        =   36
         Top             =   735
         Width           =   7215
         Begin MSMask.MaskEdBox MinQtdVou 
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   210
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Format          =   "##"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MinValorVou 
            Height          =   315
            Left            =   2040
            TabIndex        =   12
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "vouchers e                    reais em pelo menos um dos meses analisados"
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
            Height          =   300
            Left            =   1035
            TabIndex        =   37
            Top             =   285
            Width           =   6135
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Valor Base"
         Height          =   510
         Left            =   105
         TabIndex        =   35
         Top             =   225
         Width           =   7230
         Begin VB.OptionButton ValorBase 
            Caption         =   "Faturável (Bruto - CMA)"
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
            Index           =   1
            Left            =   3525
            TabIndex        =   9
            Top             =   210
            Width           =   2415
         End
         Begin VB.OptionButton ValorBase 
            Caption         =   "Líquido (Bruto - Todas Comissões)"
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
            Index           =   3
            Left            =   195
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   3510
         End
         Begin VB.OptionButton ValorBase 
            Caption         =   "Bruto"
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
            Index           =   2
            Left            =   6135
            TabIndex        =   10
            Top             =   225
            Value           =   -1  'True
            Width           =   900
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mês / Ano"
      Height          =   630
      Left            =   180
      TabIndex        =   26
      Top             =   645
      Width           =   5325
      Begin VB.ComboBox Ano 
         Height          =   315
         ItemData        =   "RelOpDesvVendTRV.ctx":0000
         Left            =   3795
         List            =   "RelOpDesvVendTRV.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   1095
      End
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "RelOpDesvVendTRV.ctx":0004
         Left            =   1500
         List            =   "RelOpDesvVendTRV.ctx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label LabelAno 
         Caption         =   "Ano:"
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
         Height          =   240
         Left            =   3345
         TabIndex        =   30
         Top             =   240
         Width           =   420
      End
      Begin VB.Label labelMes 
         Caption         =   "Mês:"
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
         Height          =   285
         Left            =   1020
         TabIndex        =   29
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame FrameDesvios 
      Caption         =   "Desvios"
      Height          =   1215
      Left            =   180
      TabIndex        =   27
      Top             =   1350
      Width           =   7440
      Begin VB.Frame Frame1 
         Caption         =   "Considerar desvios"
         Height          =   555
         Left            =   105
         TabIndex        =   31
         Top             =   555
         Width           =   7230
         Begin VB.OptionButton Desvio 
            Caption         =   "Ambos"
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
            Index           =   3
            Left            =   6150
            TabIndex        =   7
            Top             =   285
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton Desvio 
            Caption         =   "Só para mais"
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
            Index           =   1
            Left            =   225
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   225
            Width           =   1695
         End
         Begin VB.OptionButton Desvio 
            Caption         =   "Só para menos"
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
            Index           =   2
            Left            =   3540
            TabIndex        =   6
            Top             =   255
            Width           =   1920
         End
      End
      Begin MSMask.MaskEdBox DesvMes 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   195
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesvAno 
         Height          =   315
         Left            =   5415
         TabIndex        =   4
         Top             =   195
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCargaMax 
         Caption         =   "No mesmo mês do ano anterior:"
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
         Left            =   2655
         TabIndex        =   33
         Top             =   270
         Width           =   2865
      End
      Begin VB.Label LabelCargaMin 
         Caption         =   "No mês anterior:"
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
         Left            =   60
         TabIndex        =   32
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6135
      Picture         =   "RelOpDesvVendTRV.ctx":0098
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   675
      Width           =   1485
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDesvVendTRV.ctx":019A
      Left            =   795
      List            =   "RelOpDesvVendTRV.ctx":019C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2820
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5475
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDesvVendTRV.ctx":019E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDesvVendTRV.ctx":031C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDesvVendTRV.ctx":084E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDesvVendTRV.ctx":09D8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Left            =   120
      TabIndex        =   28
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDesvVendTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colItensCategoria As New Collection
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
   
    Call Carrega_Mes_Ano
    
    UsuRespCallCenter.Clear
    
    objCategoriaCliente.sCategoria = TRV_CATEGORIA_RESPONSAVEL

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objUsuarios In colUsuarios
        UsuRespCallCenter.AddItem objUsuarios.sCodUsuario
    Next
    
    Regiao.Clear
    
    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Regiao.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    Responsavel.Clear
    
    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Responsavel.AddItem ""

    'Insere na combo CategoriaClienteItem
    For Each objCategoriaClienteItem In colItensCategoria
        'Insere na combo CategoriaCliente
        Responsavel.AddItem objCategoriaClienteItem.sItem
    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169106)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 87218

    'pega o Desvio e exibe
    lErro = objRelOpcoes.ObterParametro("NDESVIO", sParam)
    If lErro <> SUCESSO Then gError 87248

    Desvio(StrParaInt(sParam)).Value = True

    'pega o ValorBase e exibe
    lErro = objRelOpcoes.ObterParametro("NVALORBASE", sParam)
    If lErro <> SUCESSO Then gError 87248

    ValorBase(StrParaInt(sParam)).Value = True

    'pega o ValorBase e exibe
    lErro = objRelOpcoes.ObterParametro("NTRAZERCLINCOMP", sParam)
    If lErro <> SUCESSO Then gError 87248

    If StrParaInt(sParam) = MARCADO Then
        TrazerCliNComp.Value = vbChecked
    Else
        TrazerCliNComp.Value = vbUnchecked
    End If

    'pega o mês
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 87221
        
    'Aribui mês
    sMes = sParam
            
    'pega o ano
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    
    'Atribui o ano
    sAno = sParam
    
    'Com valores atribuídos de mês e ano, carrega as combos
    Call Carrega_Mes_Ano(sMes, sAno)
    
    'pega o MinQtdVou e exibe
    lErro = objRelOpcoes.ObterParametro("NMINQTDVOU", sParam)
    If lErro <> SUCESSO Then gError 87248

    MinQtdVou.PromptInclude = False
    MinQtdVou.Text = Format(StrParaInt(sParam), "###")
    MinQtdVou.PromptInclude = True
    
    'pega o MinValorVou e exibe
    lErro = objRelOpcoes.ObterParametro("NMINVALORVOU", sParam)
    If lErro <> SUCESSO Then gError 87248

    MinValorVou.Text = Format(StrParaDbl(sParam), "STANDARD")
   
    'pega o DesvMes e exibe
    lErro = objRelOpcoes.ObterParametro("NDESVMES", sParam)
    If lErro <> SUCESSO Then gError 87248

    DesvMes.Text = CStr(StrParaDbl(sParam) * 100)
   
    'pega o DesvAno e exibe
    lErro = objRelOpcoes.ObterParametro("NDESVANO", sParam)
    If lErro <> SUCESSO Then gError 87248

    DesvAno.Text = CStr(StrParaDbl(sParam) * 100)
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError 87248

    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = sParam
        Vendedor_Validate (bSGECancelDummy)
    Else
        Vendedor.Text = ""
    End If
    
    lErro = objRelOpcoes.ObterParametro("NREGIAO", sParam)
    If lErro <> SUCESSO Then gError 87248

    If StrParaInt(sParam) <> 0 Then
        Regiao.Text = sParam
        Regiao_Validate (bSGECancelDummy)
    Else
        Regiao.Text = ""
    End If
    
    lErro = objRelOpcoes.ObterParametro("NORDEM", sParam)
    If lErro <> SUCESSO Then gError 87248

    If StrParaInt(sParam) = 0 Then
        OptOrdRelevancia.Value = True
    Else
        OptOrdAlfa.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("TRESPCALLCENTER", sParam)
    If lErro <> SUCESSO Then gError 87248

    UsuRespCallCenter.Text = sParam
    UsuRespCallCenter_Validate (bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("NRESPONSAVEL", sParam)
    If lErro <> SUCESSO Then gError 87248
       
    Responsavel.ListIndex = StrParaInt(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 87218 To 87221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169107)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoVendedor = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 87222

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 87223

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 87222

        Case 87223
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169108)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 87224

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Desvio(3).Value = True
    ValorBase(2).Value = True
    TrazerCliNComp.Value = vbUnchecked

    Call Carrega_Mes_Ano

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 87224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169109)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sMes As String, sAno As String
Dim objSel As New ClassRelTRVDesviosVendSel
Dim iOrdem As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 87226
        
    sMes = CStr(Mes.ItemData(Mes.ListIndex))
            
    lErro = objRelOpcoes.IncluirParametro("NMES", sMes)
    If lErro <> AD_BOOL_TRUE Then gError 87231

    lErro = objRelOpcoes.IncluirParametro("NANO", Ano.Text)
    If lErro <> AD_BOOL_TRUE Then gError 87232

    sAno = Ano.Text

    For iIndice = 1 To 3
        If Desvio(iIndice).Value Then objSel.iDesvios = iIndice
        If ValorBase(iIndice).Value Then objSel.iValorBase = iIndice
    Next
    
    objSel.iMes = StrParaInt(sMes)
    objSel.iAno = StrParaInt(sAno)
    objSel.dMinVendVlr = StrParaDbl(MinValorVou.Text)
    objSel.dPercDesvAno = StrParaDbl(Val(DesvAno.Text) / 100)
    objSel.dPercDesvMes = StrParaDbl(Val(DesvMes.Text) / 100)
    objSel.iMinVendQtd = StrParaInt(MinQtdVou.Text)
    
    objSel.iVendedor = Codigo_Extrai(Vendedor.Text)
    objSel.sResponsavel = Responsavel.Text
    objSel.sRespCallCenter = UsuRespCallCenter.Text
    objSel.iRegiao = Codigo_Extrai(Regiao.Text)
    
    If TrazerCliNComp.Value = vbChecked Then
        objSel.iTrazerCliNComp = MARCADO
    Else
        objSel.iTrazerCliNComp = DESMARCADO
    End If
    
    If OptOrdRelevancia.Value Then
        iOrdem = 0
    Else
        iOrdem = 1
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(objSel.iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("TRESPONSAVEL", CStr(objSel.sResponsavel))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NRESPONSAVEL", CStr(Responsavel.ListIndex))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("TRESPCALLCENTER", CStr(objSel.sRespCallCenter))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NREGIAO", CStr(objSel.iRegiao))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NORDEM", CStr(iOrdem))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NDESVIO", CStr(objSel.iDesvios))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NVALORBASE", CStr(objSel.iValorBase))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NTRAZERCLINCOMP", CStr(objSel.iTrazerCliNComp))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NMINQTDVOU", CStr(objSel.iMinVendQtd))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NMINVALORVOU", CStr(objSel.dMinVendVlr))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NDESVMES", CStr(objSel.dPercDesvMes))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    lErro = objRelOpcoes.IncluirParametro("NDESVANO", CStr(objSel.dPercDesvAno))
    If lErro <> AD_BOOL_TRUE Then gError 87244
    
    If bExecutando Then
    
        GL_objMDIForm.MousePointer = vbHourglass

        lErro = CF("RelTRVDesviosVend_Prepara", objSel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        GL_objMDIForm.MousePointer = vbDefault
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(objSel.lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = gErr

    Select Case gErr

        Case 87225 To 87232

        Case 87244, 87245

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169110)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 87233

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 87234

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 87235

        ComboOpcoes.Text = ""

        Call Carrega_Mes_Ano
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 87233
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 87234, 87235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169111)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 87236

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 87236

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169112)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 87237

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 87238

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 87239

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 87240

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 87237
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 87238, 87239, 87240

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169113)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FAT_VENDEDOR
    Set Form_Load_Ocx = Me
    Caption = "Desvios de Venda por Cliente"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpDesvVendTRV"

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

Public Sub Unload(objme As Object)
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

End Sub

Private Sub LabelMes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelMes, Source, X, Y)
End Sub

Private Sub LabelMes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelMes, Button, Shift, X, Y)
End Sub

Private Sub LabelAno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAno, Source, X, Y)
End Sub

Private Sub LabelAno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAno, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Carrega_Mes_Ano(Optional sMes As String, Optional sAno As String)
'Função responsável pelo carregamento das combos ano e mês que não são editaveis

Dim iMes As Integer
Dim iAno As Integer
Dim iIndice As Integer
Dim iMax As Integer

    If Ano.ListCount = 0 Then
        For iIndice = 11 To 1 Step -1
            Ano.AddItem (Year(Date) + 1 - iIndice)
        Next
    End If

    'Se a função for chamada de Define_Padrao
    If Len(Trim(sMes)) = 0 Then
        iMes = Month(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iMes = CInt(sMes)
    End If
    
    'Se a função for chamada de Define_Padrao
    If Len(Trim(sAno)) = 0 Then
        iAno = Year(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iAno = CInt(sAno)
    End If
    
    Mes.ListIndex = iMes - 1
    
    iMax = Ano.ListCount
       
    For iIndice = 1 To iMax
        If Ano.List(iIndice) = CStr(iAno) Then
            Ano.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Sub DesvAno_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DesvAno_Validate

    'Veifica se DesvAno está preenchida
    If Len(Trim(DesvAno.Text)) <> 0 Then

       'Critica a DesvAno
       lErro = Porcentagem_Critica(DesvAno.Text)
       If lErro <> SUCESSO Then gError 134368

    End If

    Exit Sub

Erro_DesvAno_Validate:

    Cancel = True

    Select Case gErr

        Case 134368

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144357)

    End Select

    Exit Sub

End Sub

Private Sub DesvAno_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DesvAno, iAlterado)
    
End Sub

Private Sub DesvMes_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DesvMes_Validate

    'Veifica se DesvMes está preenchida
    If Len(Trim(DesvMes.Text)) <> 0 Then

       'Critica a DesvMes
       lErro = Porcentagem_Critica(DesvMes.Text)
       If lErro <> SUCESSO Then gError 134368

    End If

    Exit Sub

Erro_DesvMes_Validate:

    Cancel = True

    Select Case gErr

        Case 134368

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144357)

    End Select

    Exit Sub

End Sub

Private Sub DesvMes_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DesvMes, iAlterado)
    
End Sub

Private Sub MinQtdVou_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MinQtdVou, iAlterado)
    
End Sub

Private Sub MinValorVou_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MinValorVou_Validate

    'Veifica se MinValorVou está preenchida
    If Len(Trim(MinValorVou.Text)) <> 0 Then

       'Critica a MinValorVou
       lErro = Valor_Positivo_Critica(MinValorVou.Text)
       If lErro <> SUCESSO Then gError 137574

    End If

    Exit Sub

Erro_MinValorVou_Validate:

    Cancel = True

    Select Case gErr

        Case 137574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174977)

    End Select

    Exit Sub

End Sub

Private Sub MinValorVou_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MinValorVou, iAlterado)
    
End Sub

Public Sub UsuRespCallCenter_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_UsuRespCallCenter_Validate
    
    If Len(Trim(UsuRespCallCenter.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = UsuRespCallCenter.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 190458
        
        If lErro <> SUCESSO Then gError 190459
        
    End If
    
    Exit Sub
    
Erro_UsuRespCallCenter_Validate:

    Cancel = True

    Select Case gErr
            
        Case 190458
        
        Case 190459 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190460)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'Verifica se foi preenchido o campo Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
    If Regiao.Text = Regiao.List(Regiao.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19235

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tenta ler Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 19236
        
        'Não encontrou Regiao Venda BD
        If lErro <> SUCESSO Then Error 19237
        
        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19238

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case Err

    Case 19235, 19236

    Case 19237  'Não encontrou RegiaoVenda no BD
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

        If vbMsgRes = vbYes Then
            'Chama a tela RegiaoVenda
            Call Chama_Tela("RegiaoVenda", objRegiaoVenda)

        End If

    Case 19238
        lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", Err, Regiao.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155581)

    End Select

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
        
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
   
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub
