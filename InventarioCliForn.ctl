VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl InventarioCliForn 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.TextBox Observacao 
      Height          =   1455
      Left            =   4980
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4470
      Width           =   4440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saldos em Outros Clientes"
      Height          =   1755
      Left            =   30
      TabIndex        =   29
      Top             =   4170
      Width           =   4860
      Begin VB.TextBox FilialGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   40
         Top             =   660
         Width           =   435
      End
      Begin VB.TextBox Tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   675
         MaxLength       =   50
         TabIndex        =   39
         Top             =   675
         Width           =   435
      End
      Begin MSMask.MaskEdBox QtdCli 
         Height          =   210
         Left            =   3465
         TabIndex        =   35
         Top             =   930
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.TextBox ClienteGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   315
         MaxLength       =   50
         TabIndex        =   34
         Top             =   945
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid GridProdCli 
         Height          =   1500
         Left            =   60
         TabIndex        =   9
         Top             =   210
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   2646
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   570
      Left            =   15
      TabIndex        =   26
      Top             =   30
      Width           =   7230
      Begin VB.ComboBox Escaninho 
         Height          =   315
         ItemData        =   "InventarioCliForn.ctx":0000
         Left            =   3555
         List            =   "InventarioCliForn.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   3615
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2205
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1125
         TabIndex        =   0
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelEscaninho 
         AutoSize        =   -1  'True
         Caption         =   "Escaninho:"
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
         Left            =   2565
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   570
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   2820
      Left            =   15
      TabIndex        =   24
      Top             =   1335
      Width           =   9465
      Begin MSMask.MaskEdBox QtdAcerto 
         Height          =   210
         Left            =   6405
         TabIndex        =   37
         Top             =   1485
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.TextBox Obs 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   2505
         MaxLength       =   50
         TabIndex        =   33
         Top             =   1485
         Width           =   2600
      End
      Begin MSMask.MaskEdBox QtdDistribData 
         Height          =   210
         Left            =   7830
         TabIndex        =   32
         Top             =   1215
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QtdData 
         Height          =   210
         Left            =   7380
         TabIndex        =   31
         Top             =   855
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QtdEncontCliData 
         Height          =   210
         Left            =   6285
         TabIndex        =   30
         Top             =   840
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.TextBox UM 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   3945
         TabIndex        =   25
         Top             =   690
         Width           =   630
      End
      Begin VB.TextBox Descricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1125
         Width           =   2430
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   870
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QtdCliData 
         Height          =   210
         Left            =   5145
         TabIndex        =   18
         Top             =   840
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridProd 
         Height          =   1965
         Left            =   45
         TabIndex        =   8
         Top             =   210
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   3466
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   7275
      ScaleHeight     =   435
      ScaleWidth      =   2145
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   105
      Width           =   2205
      Begin VB.CommandButton BotaoAtualizar 
         Height          =   330
         Left            =   15
         Picture         =   "InventarioCliForn.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Atualizar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   438
         Picture         =   "InventarioCliForn.ctx":0456
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1284
         Picture         =   "InventarioCliForn.ctx":05B0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1710
         Picture         =   "InventarioCliForn.ctx":0AE2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   861
         Picture         =   "InventarioCliForn.ctx":0C60
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   390
      End
   End
   Begin VB.Frame FrameClienteFornec 
      Caption         =   "Dados do Terceiro"
      Height          =   750
      Left            =   15
      TabIndex        =   19
      Top             =   585
      Width           =   9465
      Begin VB.OptionButton OptionTipoTerc 
         Caption         =   "Cliente"
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
         Index           =   1
         Left            =   3315
         TabIndex        =   3
         Top             =   150
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton OptionTipoTerc 
         Caption         =   "Fornecedor"
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
         Left            =   4860
         TabIndex        =   4
         Top             =   150
         Width           =   1380
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   6675
         TabIndex        =   7
         Top             =   375
         Width           =   1965
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   1125
         TabIndex        =   5
         Top             =   375
         Visible         =   0   'False
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   1125
         TabIndex        =   6
         Top             =   375
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label FornecedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   435
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Left            =   6180
         TabIndex        =   21
         Top             =   435
         Width           =   465
      End
      Begin VB.Label ClienteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   435
         Width           =   660
      End
   End
   Begin VB.Label NumIntDoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   7320
      TabIndex        =   38
      Top             =   4185
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4995
      TabIndex        =   36
      Top             =   4245
      Width           =   1080
   End
End
Attribute VB_Name = "InventarioCliForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'obj c/ eventos dos browsers
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoInventario As AdmEvento
Attribute objEventoInventario.VB_VarHelpID = -1

'variaveis do grid
Dim objGridProd As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_QtdData_Col As Integer
Dim iGrid_QtdEncontCliData_Col As Integer
Dim iGrid_QtdCliData_Col As Integer
Dim iGrid_QtdDistribData_Col As Integer
Dim iGrid_QtdAcerto_Col As Integer
Dim iGrid_Obs_Col As Integer

Dim objGridProdCli As AdmGrid
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_QtdCli_Col As Integer

'variaveis de alteração
Dim iAlterado As Integer
Dim iTipoCliFornAnt As Integer
Dim lCliFornAnt As Long
Dim iFilialAnt As Integer
Dim dtDataAnt As Date
Dim iEscaninhoAnt As Integer

Dim gobjInv As New ClassInvCliForn

'variavel de controle de cód. do escaninho
Dim giEscaninho As Integer
Dim gbTrazendoDados As Boolean

Public Function Trata_Parametros(Optional objInv As ClassInvCliForn) As Long
'traz os dados p/ a tela de acordo c/ os parametros passados
    
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
        
    'se foi passado algum parametro
    If Not objInv Is Nothing Then
        
        'traz os dados p/ a tela de acordo c/ os parametros
        lErro = Traz_InventarioTerc_Tela(objInv)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209540)
            
    End Select
            
    Exit Function
        
End Function

Public Sub Form_Load()
'Carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    gbTrazendoDados = False
      
    Set objGridProd = New AdmGrid
    Set objGridProdCli = New AdmGrid

    'Inicializa os eventos dos browser
    Set objEventoCliente = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoInventario = New AdmEvento
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'Le todos os escaninhos que podem ser de terceiro \ nosso
    lErro = Carrega_Escaninhos
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inicializa o grid
    lErro = Inicializa_GridProd(objGridProd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_GridProdCli(objGridProdCli)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'zera as variáveis de alteração
    iAlterado = 0
    giEscaninho = -1
    
    Call DateParaMasked(Data, gdtDataAtual)
    
    iTipoCliFornAnt = TIPO_TERC_CLIENTE
    lCliFornAnt = 0
    dtDataAnt = DATA_NULA
    iEscaninhoAnt = 0
    iFilialAnt = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209541)

    End Select
    
    Exit Sub

End Sub

Private Function Carrega_Escaninhos() As Long
'carrega os escaninhos referentes a 3ºs
    
Dim lErro As Long
Dim colEscaninhos As New Collection
Dim objEscaninho As ClassEscaninho

On Error GoTo Erro_Carrega_Escaninhos

    'lê na tabela de escaninhos os escaninhos relacionados a 3ºs
    lErro = CF("Escaninhos_Le_Terceiros", colEscaninhos)
    If lErro <> SUCESSO And lErro <> 119709 Then gError ERRO_SEM_MENSAGEM
    
    'sem dados
    If lErro = 119709 Then gError 209542
    
    'p/ cada escaninho na coleção de escaninhos
    For Each objEscaninho In colEscaninhos
    
        'adiciona na combo box e guarda o cód. do escaninho no itemdata
        Escaninho.AddItem objEscaninho.sNome
        Escaninho.ItemData(Escaninho.NewIndex) = objEscaninho.iCodigo
    
    Next
    
    Carrega_Escaninhos = SUCESSO

    Exit Function

Erro_Carrega_Escaninhos:

    Carrega_Escaninhos = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 209542
            Call Rotina_Erro(vbOKOnly, "ERRO_ESCANINHOS_INEXISTENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209543)
            
    End Select
        
    Exit Function
        
End Function

Private Function Move_Tela_Memoria(ByVal objInv As ClassInvCliForn) As Long
'move os dados da tela p/ a memória, no objInventarioTerc

Dim lErro As Long
Dim objCliente As ClassCliente
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    'se for cliente
    If OptionTipoTerc(1).Value = True Then
        
        'instancia o obj cliente
        Set objCliente = New ClassCliente
        
        'carrega o obj c/ o tipo de terceiro (cliente)
        objInv.iTipoCliForn = TIPO_TERC_CLIENTE
        
        If Len(Trim(Cliente.Text)) > 0 Then
        
            'preenche o objcliente c/ o nomered do cliente na tela
            objCliente.sNomeReduzido = Trim(Cliente.Text)
            
            'busca o cód. do cliente a apartir do nomereduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
            
            'cliente não cadastrado
            If lErro = 12348 Then gError 209544
            
            'preenche o obj c/ o cód do cliente
            objInv.lCliForn = objCliente.lCodigo
            
        End If
        
    'senão, é fornecedor
    Else
    
        'instancia o obj fornecedor
        Set objFornecedor = New ClassFornecedor
        
        'carrega o obj c/ o tipo de terceiro (fornecedor)
        objInv.iTipoCliForn = TIPO_TERC_FORNECEDOR
        
        If Len(Trim(Fornecedor.Text)) > 0 Then
        
            'carrega o objforncedor com o nomered do fornecedor da tela
            objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
            
            'busca o cód. do fornecedor a apartir do nomereduzido
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM
        
            'fornecedor não cadastrado
            If lErro = 6681 Then gError 209545
        
            'preenche o obj c/ o cód. do fornecedor
            objInv.lCliForn = objFornecedor.lCodigo
            
        End If
    
    End If

    'passa o código da filial do cliente / fornecedor p/ o obj
    objInv.iFilial = Codigo_Extrai(Filial.Text)
    objInv.dtData = StrParaDate(Data.Text)
    objInv.sOBS = Observacao.Text
    objInv.iFilialEmpresa = giFilialEmpresa
    objInv.iEscaninho = giEscaninho
    
    lErro = Move_GridProd_Memoria(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 209544
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor.Text)
                    
        Case 209545
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209546)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'inicia a etapa de exclusão de registros

Dim lErro As Long
Dim objInv As New ClassInvCliForn
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'ponteiro p/ ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'verifica se os campos obrigatórios estão preenchidos
    lErro = Verifica_Preenchimento_Obrigatorio
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_INVENTARIOTERCPROD")
    
    'se não, sai da rotina
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'move os dados de fora do grid p/ o obj
    lErro = Move_Tela_Memoria(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'exclui os dados de terceiros cadastrados
    lErro = CF("InvCliForn_Exclui", objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Limpa a Tela
    Call Limpa_Tela_InventarioTerc
    
    'ampulheta p/ padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209548)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'inicia a gravacao

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a Tela
    Call Limpa_Tela_InventarioTerc

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209549)

    End Select

    Exit Sub
End Sub

Public Function Gravar_Registro() As Long
'verifica o preenchimento da tela e chama a rotina de gravação

Dim lErro As Long
Dim objInv As New ClassInvCliForn

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro um ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica o preenchimento dos campos obrigatórios de fora do grid
    lErro = Verifica_Preenchimento_Obrigatorio
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'carrega o obj c/ os dados da tela
    lErro = Move_Tela_Memoria(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'verifica se ja existe algum registro, se existir, pergunta se deseja atualizar o registro existente
    lErro = Trata_Alteracao(objInv, objInv.iFilialEmpresa, objInv.dtData, objInv.iEscaninho, objInv.iTipoCliForn, objInv.lCliForn, objInv.iFilial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'guarda o registro de 3ºs em InventarioTerc
    lErro = CF("InvCliForn_Grava", objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'volta o ponteiro no padrao
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209550)

    End Select

    Exit Function

End Function

Private Function Verifica_Preenchimento_Obrigatorio() As Long
'veririfica se os campos obrigatórios estão preenchidos

On Error GoTo Erro_Verifica_Preenchimento_Obrigatorio

    If gobjMAT.iTrataEstTercCliForn = DESMARCADO Then gError 209551

    'se a data não estiver preenchida ==> erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 119725

    'se o tipo de terceiro for cliente
    If OptionTipoTerc(1).Value = True Then
        
        'verifica se o cliente está preenchido
        If Len(Trim(Cliente.Text)) = 0 Then gError 119658
        
    'senão, é fornecedor
    Else
            
        'verifica se o fornecedor está preenchido
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 119659
    
    End If

    'verifica se a filial está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 119660
   
    'verifica se o escaninho foi selecionado
    If Escaninho.ListIndex = -1 Then gError 119661
       
    Verifica_Preenchimento_Obrigatorio = SUCESSO

    Exit Function

Erro_Verifica_Preenchimento_Obrigatorio:

    Verifica_Preenchimento_Obrigatorio = gErr

    Select Case gErr
    
        Case 209551
            Call Rotina_Erro(vbOKOnly, "ERRO_CTRL_EST_TERC_NAO_CONFIG", gErr)
    
        Case 209552
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
    
        Case 209553
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 209554
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 209555
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 209556
            Call Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209557)

    End Select

    Exit Function

End Function

Private Function Move_GridProd_Memoria(objInv As ClassInvCliForn) As Long
'move os dados do grid p/ o obj passado como parametro

Dim lErro As Long
Dim iPreenchido As Integer
Dim sProduto As String
Dim objInvItem As ClassInvCliFornItens
Dim iIndice As Integer

On Error GoTo Erro_Move_GridProd_Memoria

    'preenche uma colecao com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridProd.iLinhasExistentes

        'instancia o obj p/ receber os dados do grid
        Set objInvItem = New ClassInvCliFornItens

        'formata o produto
        lErro = CF("Produto_Formata", GridProd.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Se o produto não estiver preenchido => erro
        If iPreenchido = PRODUTO_VAZIO Then gError 209558
        
        'verifica se a quantidade está preenchida
        If StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_QtdAcerto_Col)) <> 0 Then
                        
            'preenche o obj c/ o produto formatado
            objInvItem.sProduto = sProduto
            
            'carrega o obj c/ a quant de cada produto
            objInvItem.iSeq = iIndice
            objInvItem.dQtdAcerto = StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_QtdAcerto_Col))
            objInvItem.dQtdCliData = StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_QtdCliData_Col))
            objInvItem.dQtdData = StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_QtdData_Col))
            objInvItem.dQtdEncontCliData = StrParaDbl(GridProd.TextMatrix(iIndice, iGrid_QtdEncontCliData_Col))
            objInvItem.sOBS = GridProd.TextMatrix(iIndice, iGrid_Obs_Col)
                
            'adiciona o objInventarioTercProd na col. do objInventarioTerc que foi passada como parametro
            objInv.colItens.Add objInvItem
            
        End If

    Next
            
    Move_GridProd_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridProd_Memoria:
    
    Move_GridProd_Memoria = gErr
    
    Select Case gErr
        
        Case 209558
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209559)

    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objInv As New ClassInvCliForn

On Error GoTo Erro_Tela_Extrai

    'Informa a view associada à Tela
    sTabela = "InvCliForn"
    
    'carrega o obj c/ o tipo de terc, cód e filial (se estiverem preenchidos)
    lErro = Move_Tela_Memoria(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Data", objInv.dtData, 0, "Data"
    colCampoValor.Add "TipoCliForn", objInv.iTipoCliForn, 0, "TipoCliForn"
    colCampoValor.Add "CliForn", objInv.lCliForn, 0, "CliForn"
    colCampoValor.Add "Filial", objInv.iFilial, 0, "Filial"
    colCampoValor.Add "Escaninho", objInv.iEscaninho, 0, "Escaninho"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209560)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objInv As New ClassInvCliForn

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objInventarioTerc
    objInv.dtData = colCampoValor.Item("Data").vValor
    objInv.lCliForn = colCampoValor.Item("CliForn").vValor
    objInv.iTipoCliForn = colCampoValor.Item("TipoCliForn").vValor
    objInv.iFilial = colCampoValor.Item("Filial").vValor
    objInv.iEscaninho = colCampoValor.Item("Escaninho").vValor
    objInv.iFilialEmpresa = giFilialEmpresa

    'preenche a tela c/ os dados carregados em objInventarioTerc
    lErro = Traz_InventarioTerc_Tela(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209561)

    End Select

    Exit Sub

End Sub

Private Function Traz_InventarioTerc_Tela(objInv As ClassInvCliForn) As Long
'traz os dados do da view InventarioTerc p/ a tela, que foram carregados no obj

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_InventarioTerc_Tela

    lErro = CF("InvCliForn_Le", objInv)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
    If lErro = SUCESSO Then
    
        gbTrazendoDados = True
            
        lErro = CF("InvCliFornItens_Le", objInv)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'preenche a data da tela c/ a data gravada no obj
        Data.PromptInclude = False
        Data.Text = Format(objInv.dtData, "dd/mm/yy")
        Data.PromptInclude = True
        
        NumIntDoc.Caption = CStr(objInv.lNumIntDoc)
        
        'se for cliente
        If objInv.iTipoCliForn = TIPO_TERC_CLIENTE Then
            
            'coloca a opt. cliente como marcada
            OptionTipoTerc(1).Value = True
            Call OptionTipoTerc_Trata
            
            'coloca o cód. do cliente na tela e chama o validate
            Cliente.Text = objInv.lCliForn
            Call Cliente_Validate(bSGECancelDummy)
            
            'coloca a filial do cliente na tela
            Filial.Text = objInv.iFilial
            Call Filial_Validate(bSGECancelDummy)
                   
        'senão, é o fornecedor
        ElseIf objInv.iTipoCliForn = TIPO_TERC_FORNECEDOR Then
        
            'coloca a opt. do fornecedor marcada
            OptionTipoTerc(2).Value = True
            Call OptionTipoTerc_Trata
        
            'coloca o cód. do fornecedor na tela e chama o validate
            Fornecedor.Text = objInv.lCliForn
            Call Fornecedor_Validate(bSGECancelDummy)
            
            'coloca a filial do fornecedor na tela
            Filial.Text = objInv.iFilial
            Call Filial_Validate(bSGECancelDummy)
            
        End If
            
        'traz o escaninho gravado no bd p/ a tela
        For iIndice = 0 To (Escaninho.ListCount - 1)
            If Escaninho.ItemData(iIndice) = objInv.iEscaninho Then
                Escaninho.ListIndex = iIndice
                Exit For
            End If
        Next
        
        Observacao.Text = objInv.sOBS
        
        'preenche o grid de produtos c/ oq estiver gravado no bd
        lErro = Traz_GridProd_Tela(objInv)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    dtDataAnt = objInv.dtData
    iEscaninhoAnt = objInv.iEscaninho
    lCliFornAnt = objInv.lCliForn
    iFilialAnt = objInv.iFilial
    iTipoCliFornAnt = objInv.iTipoCliForn

    'zera as variaveis de alteração
    iAlterado = 0
    gbTrazendoDados = False

    Traz_InventarioTerc_Tela = SUCESSO

    Exit Function

Erro_Traz_InventarioTerc_Tela:

    gbTrazendoDados = False

    Traz_InventarioTerc_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209562)

    End Select

    Exit Function

End Function

Private Function Traz_GridProd_Tela(ByVal objInv As ClassInvCliForn) As Long
'traz os dados dos produtos de 3ºs p/ a tela

Dim lErro As Long
Dim iIndice As Integer
Dim objInvItem As ClassInvCliFornItens
Dim dQuantTotal As Double, sProdMask As String
            
On Error GoTo Erro_Traz_GridProd_Tela

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGridProd)
    
    'Exibe os dados da coleção na tela
    For Each objInvItem In objInv.colItens

        iIndice = iIndice + 1
        
        Call Mascara_RetornaProdutoEnxuto(objInvItem.sProduto, sProdMask)

        'insere o produto no controle
        Produto.PromptInclude = False
        Produto.Text = sProdMask
        Produto.PromptInclude = True
        
        'Insere o controle do produto no Grid de produtos
        GridProd.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text

        'busca a descricao do produto e a un. do produto
        lErro = Produto_Linha_Preenche(iIndice)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
        'Insere a quant no Grid de produtos
        GridProd.TextMatrix(iIndice, iGrid_QtdCliData_Col) = Formata_Estoque(objInvItem.dQtdCliData)
        GridProd.TextMatrix(iIndice, iGrid_QtdData_Col) = Formata_Estoque(objInvItem.dQtdData)
        GridProd.TextMatrix(iIndice, iGrid_QtdDistribData_Col) = Formata_Estoque(objInvItem.dQtdDistrib)
        GridProd.TextMatrix(iIndice, iGrid_Obs_Col) = objInvItem.sOBS
        
        If objInvItem.dQtdEncontCliData <> 0 Then
            GridProd.TextMatrix(iIndice, iGrid_QtdEncontCliData_Col) = Formata_Estoque(objInvItem.dQtdEncontCliData)
        Else
            GridProd.TextMatrix(iIndice, iGrid_QtdEncontCliData_Col) = ""
        End If
        If objInvItem.dQtdAcerto <> 0 Then
            GridProd.TextMatrix(iIndice, iGrid_QtdAcerto_Col) = Formata_Estoque(objInvItem.dQtdAcerto)
        Else
            GridProd.TextMatrix(iIndice, iGrid_QtdAcerto_Col) = ""
        End If
    Next

    'atualiza a qnt de linhas do grid
    objGridProd.iLinhasExistentes = iIndice
    
    Set gobjInv = objInv
    
    Traz_GridProd_Tela = SUCESSO
    
    Exit Function

Erro_Traz_GridProd_Tela:

    Traz_GridProd_Tela = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209563)
        
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoFechar_Click()
'fecha a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Fechar
    
    'se tiver alteração na tela, pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Unload Me
    
    Exit Sub
    
Erro_Botao_Fechar:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209564)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar

    'se houve alteração, pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'limpa a tela toda
    Call Limpa_Tela_InventarioTerc

    Exit Sub

Erro_BotaoLimpar:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209565)
            
    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_InventarioTerc()
'limpa a tela e o grid

    'limpa a tela(text e masks) e o grid
    Set gobjInv = New ClassInvCliForn
    
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridProd)
    Call Grid_Limpa(objGridProdCli)
    
    'limpa a combo e as labels
    Escaninho.ListIndex = -1
    Filial.Clear
    
    'zera as variáveis de alteração
    iAlterado = 0
    giEscaninho = -1
    
    iTipoCliFornAnt = TIPO_TERC_CLIENTE
    lCliFornAnt = 0
    dtDataAnt = DATA_NULA
    iEscaninhoAnt = 0
    iFilialAnt = 0
    OptionTipoTerc(1).Value = True
    
    NumIntDoc.Caption = ""
    
    Call DateParaMasked(Data, gdtDataAtual)

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Escaninho_Click()
'verifica o escaninho selecionado

On Error GoTo Erro_Escaninho_Click

    'se o escaninho não foi selecionado, sai da rotina (no caso de limpar a tela)
    If Escaninho.ListIndex = -1 Then Exit Sub
 
    'carrega em giEscaninho o cód. do escaninho correspondente
    If giEscaninho <> Escaninho.ItemData(Escaninho.ListIndex) Then
        giEscaninho = Escaninho.ItemData(Escaninho.ListIndex)
        
        Select Case giEscaninho
        
            Case 2 '2   Conserto - Nosso em Poder de Terceiros
                OptionTipoTerc(2).Value = True
            Case 3 '3   Consignação - Nosso em Poder de Terceiros
                OptionTipoTerc(1).Value = True
            Case 4 '4   Demonstração - Nosso em Poder de Terceiros
                OptionTipoTerc(1).Value = True
            Case 5 '5   Outros - Nosso em Poder de Terceiros
                OptionTipoTerc(2).Value = True
            Case 6 '6   Beneficiamento - Nosso em Poder de Terceiros
                OptionTipoTerc(2).Value = True
            Case 7 '7   Conserto - De Terceiros em Nosso Poder
                OptionTipoTerc(1).Value = True
            Case 8 '8   Consignação - De Terceiros em Nosso Poder
                OptionTipoTerc(2).Value = True
            Case 9 '9   Demonstração - De Terceiros em Nosso Poder
                OptionTipoTerc(2).Value = True
            Case 10 '10  Outros - De Terceiros em Nosso Poder
                OptionTipoTerc(2).Value = True
            Case 11 '11  Beneficiamento - De Terceiros em Nosso Poder
                OptionTipoTerc(1).Value = True
        End Select
        Call OptionTipoTerc_Trata
    End If

    Exit Sub

Erro_Escaninho_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209566)

    End Select

    Exit Sub

End Sub

Private Sub ClienteLabel_Click()
'chama o browser de clientes

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
        'Prenche o nome reduzido do Cliente
        objCliente.sNomeReduzido = Trim(Cliente.Text)
    End If
    
    'chama a tela de browser clienteslista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Filial_Validate(bSGECancelDummy)
End Sub

Private Sub Fornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
'verifica se o fornecedor selecionado é valido

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate
    
    'Verifica preenchimento de Fornecedor, se não foi preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
        'limpa a filial e sai da rotina
        Filial.Clear
    
    End If
    
    If objFornecedor.lCodigo <> lCliFornAnt Then

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
    
        'verifica se foi digitado nome ou cód. do fornecedor
        If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
            If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
            
        End If
    
        Call Trata_Troca_Dados
        
    End If

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209567)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridProd(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Total")
    objGridInt.colColuna.Add ("Encontrada")
    objGridInt.colColuna.Add ("Qtde Cli.")
    objGridInt.colColuna.Add ("Acerto")
    objGridInt.colColuna.Add ("Identif.")
    objGridInt.colColuna.Add ("Observação")

   'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (QtdData.Name)
    objGridInt.colCampo.Add (QtdEncontCliData.Name)
    objGridInt.colCampo.Add (QtdCliData.Name)
    objGridInt.colCampo.Add (QtdAcerto.Name)
    objGridInt.colCampo.Add (QtdDistribData.Name)
    objGridInt.colCampo.Add (Obs.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UM_Col = 3
    iGrid_QtdData_Col = 4
    iGrid_QtdEncontCliData_Col = 5
    iGrid_QtdCliData_Col = 6
    iGrid_QtdAcerto_Col = 7
    iGrid_QtdDistribData_Col = 8
    iGrid_Obs_Col = 9

    'passa o grid p/ o obj
    objGridInt.objGrid = GridProd
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8

    'largura da 1ª coluna
    GridProd.ColWidth(0) = 400

    'largura Manual das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'chama a rotina que inicializa o grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridProd = SUCESSO

End Function

Private Function Inicializa_GridProdCli(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Quantidade")

   'campos de edição do grid
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (ClienteGrid.Name)
    objGridInt.colCampo.Add (FilialGrid.Name)
    objGridInt.colCampo.Add (QtdCli.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Tipo_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_Filial_Col = 3
    iGrid_QtdCli_Col = 4

    'passa o grid p/ o obj
    objGridInt.objGrid = GridProdCli
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'largura da 1ª coluna
    GridProdCli.ColWidth(0) = 400

    'largura Manual das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'chama a rotina que inicializa o grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridProdCli = SUCESSO

End Function

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'se o cliente não foi preenchido, sai da rotina
    If Len(Trim(Cliente.Text)) <> 0 Then
    
        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'busca no bd a relação de filiais referentes ao cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
        'limpa a filial e sai da rotina
        Filial.Clear
    End If
    
    If objCliente.lCodigo <> lCliFornAnt Then
    
        'Preenche ComboBox de Filiais do cliente
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'verifica se foi digitado nome ou cód. do cliente
        If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
            If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
            
        End If
    
        Call Trata_Troca_Dados
        
    End If
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209568)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)
'verifica se a filial do cliente\fornecedor é válida

Dim lErro As Long
Dim objFilialCliente As ClassFilialCliente
Dim objFilialFornecedor As ClassFilialFornecedor
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

   'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.ListIndex = -1 Then

        'se o tipo de terc. for cliente
        If OptionTipoTerc(1).Value = True Then
        
            'verifica se o cliente foi preenchido
            If Len(Trim(Cliente.Text)) = 0 Then gError 209569
    
            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(Filial, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
        
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                'instancia o obj
                Set objFilialCliente = New ClassFilialCliente
        
                'passa o nº preenchido como código
                objFilialCliente.iCodFilial = iCodigo
        
                'Tentativa de leitura da Filial com esse código no BD
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
                If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM
        
                'Não encontrou Filial no  BD
                If lErro = 17660 Then gError 209570
        
                'Encontrou Filial no BD, coloca no Text da Combo
                Filial.Text = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
            End If
                
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then gError 209571
        
        'senão, é o fornecedor
        Else
    
            'verifica se o fornecedor foi preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 209572
    
            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(Filial, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
        
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                'instancia o obj
                Set objFilialFornecedor = New ClassFilialFornecedor
        
                'passa o nº preenchido como código
                objFilialFornecedor.iCodFilial = iCodigo
        
                'Tentativa de leitura da Filial com esse código no BD
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Trim(Fornecedor.Text), objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM
        
                'Não encontrou Filial no  BD
                If lErro = 18272 Then gError 209573
        
                'Encontrou Filial no BD, coloca no Text da Combo
                Filial.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        
            End If
                
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then gError 209574
    
        End If

    End If
    
    Call Trata_Troca_Dados

    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 209569
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 209570, 209571
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 209572
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 209573, 209574
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Fornecedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209575)

    End Select

    Exit Sub

End Sub


Private Sub UpDownData_DownClick()
'diminui a data em um dia

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a Data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    Call Trata_Troca_Dados

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209576)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data em um dia

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Aumenta a Data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    Call Trata_Troca_Dados

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209577)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchido
    If Len(Trim(Data.ClipText)) > 0 Then
        
        'Critica a Data informada
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Call Trata_Troca_Dados

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209578)

    End Select

    Exit Sub
    
End Sub

Private Sub FornecedorLabel_Click()
'chama o browser referente ao fornecedor

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'se o fornecedor estiver preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Preenche nomeReduzido com o fornecedor da tela
        objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
    End If
    
    'chama a tela c/ a lista de fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub OptionTipoTerc_Click(Index As Integer)
    Call OptionTipoTerc_Trata
End Sub

Private Sub OptionTipoTerc_Trata()
'verifica qual tipo de terc. foi selecionado
   
    'se foi cliente
    If OptionTipoTerc(1).Value = True Then
    
        'limpa a filial
        If Not Cliente.Visible Then
        
            Filial.Clear
              
            'desabilita a label e o textbox do fornecedor
            Fornecedor.Visible = False
            Fornecedor.Text = ""
            FornecedorLabel.Visible = False
            
            'habilita o cliente
            Cliente.Visible = True
            ClienteLabel.Visible = True
            
        End If
        
    'senão
    Else
        
        If Not Fornecedor.Visible Then
        
            Filial.Clear
        
            'desabilita a label e o textbox do cliente
            Cliente.Visible = False
            Cliente.Text = ""
            ClienteLabel.Visible = False
            
            'habilita o fornecedor
            Fornecedor.Visible = True
            FornecedorLabel.Visible = True
    
        End If
        
    End If
    
    Call Trata_Troca_Dados

End Sub

Private Sub GridProd_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProd, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridProd, iAlterado)

    End If
    
End Sub

Private Sub GridProd_GotFocus()

    Call Grid_Recebe_Foco(objGridProd)

End Sub

Private Sub GridProd_EnterCell()

    Call Grid_Entrada_Celula(objGridProd, iAlterado)

End Sub

Private Sub GridProd_LeaveCell()

    Call Saida_Celula(objGridProd)

End Sub

Private Sub GridProd_KeyDown(KeyCode As Integer, Shift As Integer)
    'Faz o tratamento correspondente à tecla que foi pressionada
    Call Grid_Trata_Tecla1(KeyCode, objGridProd)
End Sub

Private Sub GridProd_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProd, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProd, iAlterado)
    End If

End Sub

Private Sub GridProd_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProd)
End Sub

Private Sub GridProd_RowColChange()
Dim iLinhaAntiga As Integer
    iLinhaAntiga = objGridProd.iLinhaAntiga
    Call Grid_RowColChange(objGridProd)
    If GridProd.Row <> iLinhaAntiga Then
        Call Trata_Selecao_Prod(GridProd.Row)
    End If
End Sub

Private Sub GridProd_Scroll()
    Call Grid_Scroll(objGridProd)
End Sub

Private Sub GridProdCli_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdCli, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridProdCli, iAlterado)

    End If
    
End Sub

Private Sub GridProdCli_GotFocus()

    Call Grid_Recebe_Foco(objGridProdCli)

End Sub

Private Sub GridProdCli_EnterCell()

    Call Grid_Entrada_Celula(objGridProdCli, iAlterado)

End Sub

Private Sub GridProdCli_LeaveCell()

    Call Saida_Celula(objGridProdCli)

End Sub

Private Sub GridProdCli_KeyDown(KeyCode As Integer, Shift As Integer)
    'Faz o tratamento correspondente à tecla que foi pressionada
    Call Grid_Trata_Tecla1(KeyCode, objGridProdCli)
End Sub

Private Sub GridProdCli_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdCli, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdCli, iAlterado)
    End If

End Sub

Private Sub GridProdCli_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProdCli)
End Sub

Private Sub GridProdCli_RowColChange()
    Call Grid_RowColChange(objGridProdCli)
End Sub

Private Sub GridProdCli_Scroll()
    Call Grid_Scroll(objGridProdCli)
End Sub

Private Sub Obs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Obs_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProd)
End Sub

Private Sub Obs_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProd)
End Sub

Private Sub Obs_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProd.objControle = Obs
    lErro = Grid_Campo_Libera_Foco(objGridProd)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QtdEncontCliData_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QtdEncontCliData_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProd)
End Sub

Private Sub QtdEncontCliData_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProd)
End Sub

Private Sub QtdEncontCliData_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProd.objControle = QtdEncontCliData
    lErro = Grid_Campo_Libera_Foco(objGridProd)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        Select Case GridProd.Col

            Case iGrid_Obs_Col
                'faz a saida da celula do produto
                lErro = Saida_Celula_Padrao(objGridInt, Obs)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case iGrid_QtdEncontCliData_Col
                'faz a saida da celula da quantidade
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 209579
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209579
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209580)
    
    End Select
    
    Exit Function

End Function

Private Function Produto_Linha_Preenche(ByVal iLinha As Integer) As Long
'faz a validação do produto no grid, preenchendo a descricao e a un. de medida

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Linha_Preenche

    'Critica o Produto em relação a filial
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 51385 Then gError ERRO_SEM_MENSAGEM
    
    'erro produto não encontrado
    If lErro = 51381 Then gError 209580

    'se o produto existe
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'retorna o produto enxuto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'coloca o cód. do produto no controle
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridProd.iLinhasExistentes
            If iIndice <> GridProd.Row Then
                If GridProd.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 209581
            End If
        Next
        
        'preenche a Descricao Produto
        GridProd.TextMatrix(iLinha, iGrid_Descricao_Col) = objProduto.sDescricao
        
        'preenche a unidade de medida
        GridProd.TextMatrix(iLinha, iGrid_UM_Col) = objProduto.sSiglaUMEstoque
        
    End If

    Produto_Linha_Preenche = SUCESSO

    Exit Function

Erro_Produto_Linha_Preenche:

    Produto_Linha_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 209580 'produto não existente
            
            'pergunta se deseja criar novo produto
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            
            'se sim
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridProd)

                'chama a tela de produto
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 209581
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209582)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a saida da celula de quantidade do produto

Dim lErro As Long
Dim iLinha As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = QtdEncontCliData
    
    'Se a quantidade estiver preenchida
    If Len(Trim(QtdEncontCliData.Text)) > 0 Then
        
        'Critica o valor, não pode ser negativo
        lErro = Valor_NaoNegativo_Critica(QtdEncontCliData.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'coloca no grid o valor formatado
        QtdEncontCliData.Text = Formata_Estoque(QtdEncontCliData.Text)
        
        GridProd.TextMatrix(GridProd.Row, iGrid_QtdAcerto_Col) = Formata_Estoque(-StrParaDbl(GridProd.TextMatrix(GridProd.Row, iGrid_QtdCliData_Col)) + StrParaDbl(QtdEncontCliData.Text))
        
    Else
        GridProd.TextMatrix(GridProd.Row, iGrid_QtdAcerto_Col) = ""
    End If
                                               
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
    Saida_Celula_Quantidade = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209583)
    
    End Select
    
    Exit Function

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente com o cod. do Cliente selecionado
    Cliente.Text = objCliente.lCodigo
    
    'Dispara o Validate de Cliente p/ a validação do cliente e preencher a filial
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o cód. do fornecedor na Tela
    Fornecedor.Text = objFornecedor.lCodigo

    'dispara o validate do fornecedor p/ validar o fornecedor e preencher a filial
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        End If
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
    'finaliza os objs
    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoInventario = Nothing
    Set objGridProd = Nothing
    Set objGridProdCli = Nothing
    Set gobjInv = Nothing
    
End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Inventário de Estoque Em/De Terceiros"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "InventarioCliForn"
    
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

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub LabelEscaninho_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEscaninho, Source, X, Y)
End Sub

Private Sub LabelEscaninho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEscaninho, Button, Shift, X, Y)
End Sub

Private Sub Trata_Troca_Dados()

Dim lErro As Long
Dim objInv As New ClassInvCliForn
Dim objInvBD As New ClassInvCliForn
Dim objInvItem As ClassInvCliFornItens
Dim objInvItemBD As ClassInvCliFornItens

On Error GoTo Erro_Trata_Troca_Dados

    If Not gbTrazendoDados Then

        lErro = Move_Tela_Memoria(objInv)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
        'Se está com todos dados preenchidos
        If objInv.dtData <> DATA_NULA And objInv.iEscaninho >= 0 And objInv.lCliForn <> 0 And objInv.iFilial <> 0 Then
        
            'Testa para ver se tem algo diferente
            If objInv.dtData <> dtDataAnt Or objInv.iEscaninho <> iEscaninhoAnt Or objInv.lCliForn <> lCliFornAnt Or objInv.iFilial <> iFilialAnt Or objInv.iTipoCliForn <> iTipoCliFornAnt Then
            
                'Le todos os produtos com saldo na data\escaninho passados
                lErro = CF("InvCliForn_Le_Prod_Filtro", objInv)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                objInvBD.dtData = objInv.dtData
                objInvBD.iEscaninho = objInv.iEscaninho
                objInvBD.iFilial = objInv.iFilial
                objInvBD.iFilialEmpresa = objInv.iFilialEmpresa
                objInvBD.iTipoCliForn = objInv.iTipoCliForn
                objInvBD.lCliForn = objInv.lCliForn
                
                lErro = CF("InvCliForn_Le", objInvBD)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
                'Se já tem algo cadastrado para data tem que dar uma ajeitada para permitir
                'alteração sem perder os dados anteriores
                If lErro = SUCESSO Then
                
                    lErro = CF("InvCliFornItens_Le", objInvBD)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    For Each objInvItemBD In objInvBD.colItens
                        For Each objInvItem In objInv.colItens
                            If objInvItem.sProduto = objInvItemBD.sProduto Then
                                objInvItem.dQtdCliData = objInvItem.dQtdCliData - objInvItemBD.dQtdAcerto
                                objInvItem.dQtdEncontCliData = objInvItemBD.dQtdEncontCliData
                                objInvItem.dQtdAcerto = objInvItem.dQtdEncontCliData - objInvItem.dQtdCliData
                                objInvItem.sOBS = objInvItemBD.sOBS
                                Exit For
                            End If
                        Next
                    Next
                    
                End If
                
                lErro = Traz_GridProd_Tela(objInv)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
            
        Else
        
            Call Grid_Limpa(objGridProd)
            Set gobjInv = New ClassInvCliForn
        
        End If
        
        dtDataAnt = objInv.dtData
        iEscaninhoAnt = objInv.iEscaninho
        lCliFornAnt = objInv.lCliForn
        iFilialAnt = objInv.iFilial
        iTipoCliFornAnt = objInv.iTipoCliForn
    
    End If
   
    Exit Sub

Erro_Trata_Troca_Dados:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209584)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Trata_Selecao_Prod(ByVal iLinha As Integer)

Dim lErro As Long, iIndice As Integer
Dim objInvItem As ClassInvCliFornItens
Dim objEstTerc As ClassEstoqueTerc
Dim sTexto As String, objCli As ClassCliente, objForn As ClassFornecedor

On Error GoTo Erro_Trata_Selecao_Prod

    Call Grid_Limpa(objGridProdCli)

    If iLinha <> 0 And gobjInv.colItens.Count >= iLinha Then
    
        Set objInvItem = gobjInv.colItens.Item(iLinha)
    
        iIndice = 0
        For Each objEstTerc In objInvItem.colDistribuicao
            iIndice = iIndice + 1
            GridProdCli.TextMatrix(iIndice, iGrid_QtdCli_Col) = Formata_Estoque(objEstTerc.dQuantidade)
            If objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE Then
                GridProdCli.TextMatrix(iIndice, iGrid_Tipo_Col) = "C"
                Set objCli = New ClassCliente
                objCli.lCodigo = objEstTerc.lCliForn
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                sTexto = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido
            Else
                GridProdCli.TextMatrix(iIndice, iGrid_Tipo_Col) = "F"
                Set objForn = New ClassFornecedor
                objForn.lCodigo = objEstTerc.lCliForn
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
                sTexto = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
            End If
            GridProdCli.TextMatrix(iIndice, iGrid_Cliente_Col) = sTexto
            GridProdCli.TextMatrix(iIndice, iGrid_Filial_Col) = CStr(objEstTerc.iFilial)
        Next
    
    End If

    Exit Sub

Erro_Trata_Selecao_Prod:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209585)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoAtualizar_Click()
    dtDataAnt = DATA_NULA
    iEscaninhoAnt = 0
    lCliFornAnt = 0
    iFilialAnt = 0
    iTipoCliFornAnt = 0
    Call Trata_Troca_Dados
End Sub

Private Sub LabelData_Click()

Dim lErro As Long
Dim objInv As New ClassInvCliForn
Dim colSelecao As New Collection

On Error GoTo Erro_LabelData_Click

    objInv.dtData = StrParaDate(Data.Text)
    
    'chama a tela de browser InventarioTercProdLista
    Call Chama_Tela("InvTercLista", colSelecao, objInv, objEventoInventario, "", "Data")

    Exit Sub
    
Erro_LabelData_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209621)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelEscaninho_Click()

Dim lErro As Long
Dim objInv As New ClassInvCliForn
Dim colSelecao As New Collection

On Error GoTo Erro_LabelEscaninho_Click

    objInv.iEscaninho = giEscaninho
    
    'chama a tela de browser InventarioTercProdLista
    Call Chama_Tela("InvTercLista", colSelecao, objInv, objEventoInventario, "", "Escaninho")

    Exit Sub
    
Erro_LabelEscaninho_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209622)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoInventario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objInv As New ClassInvCliForn

On Error GoTo Erro_LabelEscaninho_Click

    Set objInv = obj1
    
    lErro = Traz_InventarioTerc_Tela(objInv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub
    
Erro_LabelEscaninho_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209623)

    End Select
    
    Exit Sub
    
End Sub
