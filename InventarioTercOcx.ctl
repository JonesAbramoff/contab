VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl InventarioTerc 
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   7470
   Begin VB.CommandButton BotaoInvCadastrados 
      Caption         =   "Inventários Cadastrados"
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
      Left            =   120
      TabIndex        =   11
      Top             =   5655
      Width           =   2430
   End
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
      Height          =   345
      Left            =   5445
      TabIndex        =   12
      Top             =   5655
      Width           =   1815
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   2715
      Left            =   120
      TabIndex        =   26
      Top             =   2820
      Width           =   7140
      Begin VB.TextBox UNMedida 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   4635
         TabIndex        =   29
         Top             =   585
         Width           =   630
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   1995
         MaxLength       =   50
         TabIndex        =   9
         Top             =   570
         Width           =   2600
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   210
         Left            =   5535
         TabIndex        =   10
         Top             =   555
         Width           =   1125
         _ExtentX        =   1984
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
      Begin MSFlexGridLib.MSFlexGrid GridProdutos 
         Height          =   2160
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3810
         _Version        =   393216
      End
      Begin VB.Label QuantTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5490
         TabIndex        =   28
         Top             =   2295
         Width           =   1125
      End
      Begin VB.Label LabelTotal 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4815
         TabIndex        =   27
         Top             =   2325
         Width           =   510
      End
   End
   Begin VB.Frame FrameTerc 
      Caption         =   "Tipo do Terceiro"
      Height          =   540
      Left            =   120
      TabIndex        =   25
      Top             =   735
      Width           =   3960
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
         Height          =   300
         Index           =   2
         Left            =   2115
         TabIndex        =   2
         Top             =   180
         Width           =   1380
      End
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
         Height          =   300
         Index           =   1
         Left            =   630
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   5250
      ScaleHeight     =   465
      ScaleWidth      =   1950
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   180
      Width           =   2010
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   90
         Picture         =   "InventarioTercOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1005
         Picture         =   "InventarioTercOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1470
         Picture         =   "InventarioTercOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   540
         Picture         =   "InventarioTercOcx.ctx":080A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   390
      End
   End
   Begin VB.ComboBox Escaninho 
      Height          =   315
      ItemData        =   "InventarioTercOcx.ctx":0994
      Left            =   1845
      List            =   "InventarioTercOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2370
      Width           =   4815
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   1815
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   735
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Frame FrameClienteFornec 
      Caption         =   "Dados do Cliente/Fornecedor"
      Height          =   780
      Left            =   120
      TabIndex        =   18
      Top             =   1425
      Width           =   7155
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4995
         TabIndex        =   5
         Top             =   300
         Width           =   1500
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   1395
         TabIndex        =   4
         Top             =   307
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   1395
         TabIndex        =   3
         Top             =   307
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   360
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
         Left            =   4500
         TabIndex        =   20
         Top             =   360
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
         Left            =   720
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   360
         Width           =   660
      End
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
      Left            =   180
      TabIndex        =   24
      Top             =   285
      Width           =   480
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
      Left            =   765
      TabIndex        =   17
      Top             =   2430
      Width           =   960
   End
End
Attribute VB_Name = "InventarioTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'numero máximo se linha do grid de produtos de 3ºs
Const NUM_MAXIMO_PRODUTOS_TERCEIROS = 100

'Property Variables:
Dim m_Caption As String
Event Unload()

'obj do grid
Dim objGridProdutos As AdmGrid

'obj c/ eventos dos browsers
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoInventario As AdmEvento
Attribute objEventoInventario.VB_VarHelpID = -1

'variaveis do grid
Dim iGridProdutos_Produto_Col As Integer
Dim iGridProdutos_DescricaoItem_Col As Integer
Dim iGridProdutos_UNMedida_Col As Integer
Dim iGridProdutos_Quantidade_Col As Integer

'variaveis de alteração
Dim iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iClienteAlterado As Integer

'variavel de controle de cód. do escaninho
Dim giEscaninho As Integer

Public Function Trata_Parametros(Optional objInventarioTerc As ClassInventarioTerc) As Long
'traz os dados p/ a tela de acordo c/ os parametros passados
    
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
        
    'se foi passado algum parametro
    If Not objInventarioTerc Is Nothing Then
        
        'traz os dados p/ a tela de acordo c/ os parametros
        lErro = Traz_InventarioTerc_Tela(objInventarioTerc)
        If lErro <> SUCESSO Then gError 119643
    
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 119643
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161964)
            
    End Select
            
    Exit Function
        
End Function

Public Sub Form_Load()
'Carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
      
    'Inicializa o objGridProdutos
    Set objGridProdutos = New AdmGrid

    'Inicializa os eventos dos browser
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoInventario = New AdmEvento

    'inicializa a mask. do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 119644
        
    'Le todos os escaninhos que podem ser de terceiro \ nosso
    lErro = Carrega_Escaninhos
    If lErro <> SUCESSO Then gError 119645
    
    'inicializa o grid
    lErro = Inicializa_GridProdutos(objGridProdutos)
    If lErro <> SUCESSO Then gError 119646

    'zera as variáveis de alteração
    iAlterado = 0
    giEscaninho = -1
    iClienteAlterado = 0
    iFornecedorAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 119644, 119645, 119646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161965)

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
    If lErro <> SUCESSO And lErro <> 119709 Then gError 119775
    
    'sem dados
    If lErro = 119709 Then gError 119776
    
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
    
        Case 119775
        
        Case 119776
            Call Rotina_Erro(vbOKOnly, "ERRO_ESCANINHOS_INEXISTENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161966)
            
    End Select
        
    Exit Function
        
End Function

Private Function Move_Tela_Memoria(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'move os dados da tela p/ a memória, no objInventarioTerc

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria
            
    'move os dados de fora do grid p/ o obj
    lErro = Move_Tela_Memoria1(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119777
        
    'move os dados do grid p/ a collection em objInventarioTerc
    lErro = Move_GridProdutos_Memoria(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119650

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 119650, 119777
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161967)

    End Select
    
    Exit Function

End Function

Private Function Move_Tela_Memoria1(ByVal objInventarioTerc As ClassInventarioTerc, Optional bTrataTercInexistente As Boolean = True) As Long
'carrega o obj c/ os dados de fora do grid

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria1

    'passa a filial da empresa que está sendo usada
    objInventarioTerc.iFilialEmpresa = giFilialEmpresa
    
    'carrega o obj c/ a data preenchida
    objInventarioTerc.dtData = StrParaDate(Data.Text)
    
    'carrega o obj c/ o tipo de terc, cód e filial
    lErro = Move_TipoTerc_Memoria(objInventarioTerc)
    If lErro <> SUCESSO And lErro <> 119666 And lErro <> 119665 Then gError 119647
        
    'se for p/ mostrar o erro de cliente/filial inexistente
    If bTrataTercInexistente = True Then
        
        'cliente não cadastrado
        If lErro = 119666 Then gError 119648
            
        'fornecedor não cadastrado
        If lErro = 119665 Then gError 119649
    
    End If
    
    'carrega o obj c/ o cód. do escaninho que está em giescaninho
    objInventarioTerc.iCodEscaninho = giEscaninho

    Move_Tela_Memoria1 = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria1:

    Move_Tela_Memoria1 = gErr
    
    Select Case gErr

        Case 119647

        Case 119648
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)
        
        Case 119649
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor.Text)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'inicia a etapa de exclusão de registros

Dim lErro As Long
Dim objInventarioTerc As New ClassInventarioTerc
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'ponteiro p/ ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'verifica se os campos obrigatórios estão preenchidos
    lErro = Verifica_Preenchimento_Obrigatorio
    If lErro <> SUCESSO Then gError 119651
    
    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_INVENTARIOTERCPROD")
    
    'se não, sai da rotina
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'move os dados de fora do grid p/ o obj
    lErro = Move_Tela_Memoria1(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119652
    
    'exclui os dados de terceiros cadastrados
    lErro = CF("InventarioTerc_Exclui", objInventarioTerc)
    If lErro <> SUCESSO Then gError 119757
    
    'Limpa a Tela
    Call Limpa_Tela_InventarioTerc
    
    'ampulheta p/ padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 119651, 119652, 119757
            
        Case 119771
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor.Text)
                    
        Case 119772
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161968)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'inicia a gravacao

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 119654

    'Limpa a Tela
    Call Limpa_Tela_InventarioTerc

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 119654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161969)

    End Select

    Exit Sub
End Sub

Public Function Gravar_Registro() As Long
'verifica o preenchimento da tela e chama a rotina de gravação

Dim lErro As Long
Dim objInventarioTerc As New ClassInventarioTerc

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro um ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica o preenchimento dos campos obrigatórios de fora do grid
    lErro = Verifica_Preenchimento_Obrigatorio
    If lErro <> SUCESSO Then gError 119655
       
    'carrega o obj c/ os dados da tela
    lErro = Move_Tela_Memoria(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119656

    'verifica se ja existe algum registro, se existir, pergunta se deseja atualizar o registro existente
    lErro = Trata_Alteracao(objInventarioTerc, objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
    If lErro <> SUCESSO Then gError 119642

    'guarda o registro de 3ºs em InventarioTerc
    lErro = CF("InventarioTerc_Grava", objInventarioTerc)
    If lErro <> SUCESSO Then gError 119758
    
    'volta o ponteiro no padrao
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 119656, 119655, 119758, 119642
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161970)

    End Select

    Exit Function

End Function

Private Function Verifica_Preenchimento_Obrigatorio() As Long
'veririfica se os campos obrigatórios estão preenchidos

On Error GoTo Erro_Verifica_Preenchimento_Obrigatorio

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
    
        Case 119725
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
    
        Case 119658
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 119659
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 119660
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 119661
            Call Rotina_Erro(vbOKOnly, "ERRO_ESCANINHO_NAO_SELECIONADO", gErr)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161971)

    End Select

    Exit Function

End Function

Private Function Move_TipoTerc_Memoria(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'carrega o obj passado como parametro c/ o tipo, o cód., e a filial de terceiros

Dim lErro As Long
Dim objCliente As ClassCliente
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Move_TipoTerc_Memoria

    'se for cliente
    If OptionTipoTerc(1).Value = True Then
        
        'instancia o obj cliente
        Set objCliente = New ClassCliente
        
        'carrega o obj c/ o tipo de terceiro (cliente)
        objInventarioTerc.iTipoTerc = TIPO_TERC_CLIENTE
        
        'preenche o objcliente c/ o nomered do cliente na tela
        objCliente.sNomeReduzido = Trim(Cliente.Text)
        
        'busca o cód. do cliente a apartir do nomereduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 119663
        
        'cliente não cadastrado
        If lErro = 12348 Then gError 119666
        
        'preenche o obj c/ o cód do cliente
        objInventarioTerc.lCodTerc = objCliente.lCodigo
        
    'senão, é fornecedor
    Else
    
        'instancia o obj fornecedor
        Set objFornecedor = New ClassFornecedor
        
        'carrega o obj c/ o tipo de terceiro (fornecedor)
        objInventarioTerc.iTipoTerc = TIPO_TERC_FORNECEDOR
    
        'carrega o objforncedor com o nomered do fornecedor da tela
        objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
        
        'busca o cód. do fornecedor a apartir do nomereduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 119664
    
        'fornecedor não cadastrado
        If lErro = 6681 Then gError 119665
    
        'preenche o obj c/ o cód. do fornecedor
        objInventarioTerc.lCodTerc = objFornecedor.lCodigo
    
    End If

    'passa o código da filial do cliente / fornecedor p/ o obj
    objInventarioTerc.iFilialTerc = Codigo_Extrai(Filial.Text)

    Move_TipoTerc_Memoria = SUCESSO

    Exit Function

Erro_Move_TipoTerc_Memoria:

    Move_TipoTerc_Memoria = gErr
    
    Select Case gErr

        Case 119663, 119664

        Case 119666 'cliente não cadastrado
        
        Case 119665 'fornecedor não cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161972)

    End Select
    
    Exit Function

End Function

Private Function Move_GridProdutos_Memoria(objInventarioTerc As ClassInventarioTerc) As Long
'move os dados do grid p/ o obj passado como parametro

Dim lErro As Long
Dim iPreenchido As Integer
Dim sProduto As String
Dim objInventarioTercProd As ClassInventarioTercProd
Dim iIndice As Integer

On Error GoTo Erro_Move_GridProdutos_Memoria

    'verifica se tem ao menos uma linha preenchida no grid
    If objGridProdutos.iLinhasExistentes = 0 Then gError 119662

    'preenche uma colecao com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridProdutos.iLinhasExistentes

        'instancia o obj p/ receber os dados do grid
        Set objInventarioTercProd = New ClassInventarioTercProd

        'formata o produto
        lErro = CF("Produto_Formata", GridProdutos.TextMatrix(iIndice, iGridProdutos_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 119667
        
        'Se o produto não estiver preenchido => erro
        If iPreenchido = PRODUTO_VAZIO Then gError 119668
        
        'verifica se a quantidade está preenchida
        If Len(Trim(GridProdutos.TextMatrix(iIndice, iGridProdutos_Quantidade_Col))) = 0 Then gError 119669
                        
        'preenche o obj c/ o produto formatado
        objInventarioTercProd.sProduto = sProduto
        
        'carrega o obj c/ a quant de cada produto
        objInventarioTercProd.dQuantTotal = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGridProdutos_Quantidade_Col))
            
        'adiciona o objInventarioTercProd na col. do objInventarioTerc que foi passada como parametro
        objInventarioTerc.colInventarioTercProd.Add objInventarioTercProd

    Next
            
    Move_GridProdutos_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridProdutos_Memoria:
    
    Move_GridProdutos_Memoria = gErr
    
    Select Case gErr
        
        Case 119667

        Case 119662
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
            
        Case 119669
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, GridProdutos.TextMatrix(iIndice, iGridProdutos_Produto_Col))

        Case 119668
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161973)

    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objInventarioTerc As New ClassInventarioTerc

On Error GoTo Erro_Tela_Extrai

    'Informa a view associada à Tela
    sTabela = "InventarioTerc"
    
    'carrega o obj c/ o tipo de terc, cód e filial (se estiverem preenchidos)
    lErro = Move_Tela_Memoria1(objInventarioTerc, False)
    If lErro <> SUCESSO Then gError 119670
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Data", objInventarioTerc.dtData, 0, "Data"
    colCampoValor.Add "TipoTerc", objInventarioTerc.iTipoTerc, 0, "TipoTerc"
    colCampoValor.Add "CodTerc", objInventarioTerc.lCodTerc, 0, "CodTerc"
    colCampoValor.Add "FilialTerc", objInventarioTerc.iFilialTerc, 0, "FilialTerc"
    colCampoValor.Add "CodEscaninho", objInventarioTerc.iCodEscaninho, 0, "CodEscaninho"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 119670

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161974)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objInventarioTerc As New ClassInventarioTerc

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objInventarioTerc
    objInventarioTerc.dtData = colCampoValor.Item("Data").vValor
    objInventarioTerc.lCodTerc = colCampoValor.Item("CodTerc").vValor
    objInventarioTerc.iTipoTerc = colCampoValor.Item("TipoTerc").vValor
    objInventarioTerc.iFilialTerc = colCampoValor.Item("Filialterc").vValor
    objInventarioTerc.iCodEscaninho = colCampoValor.Item("CodEscaninho").vValor
    objInventarioTerc.iFilialEmpresa = giFilialEmpresa

    'preenche a tela c/ os dados carregados em objInventarioTerc
    lErro = Traz_InventarioTerc_Tela(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119671

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 119671

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161975)

    End Select

    Exit Sub

End Sub

Private Function Traz_InventarioTerc_Tela(objInventarioTerc As ClassInventarioTerc) As Long
'traz os dados do da view InventarioTerc p/ a tela, que foram carregados no obj

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_InventarioTerc_Tela
        
    'preenche a data da tela c/ a data gravada no obj
    Data.PromptInclude = False
    Data.Text = Format(objInventarioTerc.dtData, "dd/mm/yy")
    Data.PromptInclude = True
    
    'se for cliente
    If objInventarioTerc.iTipoTerc = TIPO_TERC_CLIENTE Then
        
        'coloca a opt. cliente como marcada
        OptionTipoTerc(1).Value = True
        
        'coloca o cód. do cliente na tela e chama o validate
        Cliente.Text = objInventarioTerc.lCodTerc
        Call Cliente_Validate(bSGECancelDummy)
        
        'coloca a filial do cliente na tela
        Filial.Text = objInventarioTerc.iFilialTerc
        Call Filial_Validate(bSGECancelDummy)
               
    'senão, é o fornecedor
    ElseIf objInventarioTerc.iTipoTerc = TIPO_TERC_FORNECEDOR Then
    
        'coloca a opt. do fornecedor marcada
        OptionTipoTerc(2).Value = True
    
        'coloca o cód. do fornecedor na tela e chama o validate
        Fornecedor.Text = objInventarioTerc.lCodTerc
        Call Fornecedor_Validate(bSGECancelDummy)
        
        'coloca a filial do fornecedor na tela
        Filial.Text = objInventarioTerc.iFilialTerc
        Call Filial_Validate(bSGECancelDummy)
        
    End If
        
    'traz o escaninho gravado no bd p/ a tela
    For iIndice = 0 To (Escaninho.ListCount - 1)
        If Escaninho.ItemData(iIndice) = objInventarioTerc.iCodEscaninho Then
            Escaninho.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'le na tabela InventarioTercProd os produtos relacionados c/ os 3ºs
    lErro = CF("InventarioTercProd_Le_InventarioTerc", objInventarioTerc)
    If lErro <> SUCESSO And lErro <> 119714 Then gError 119672

    'erro sem dados
    If lErro = 119714 Then gError 119784
    
    'preenche o grid de produtos c/ oq estiver gravado no bd
    lErro = Traz_GridProdutos_Tela(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119783

    'zera as variaveis de alteração
    iAlterado = 0
    iClienteAlterado = 0
    iFornecedorAlterado = 0

    Traz_InventarioTerc_Tela = SUCESSO

    Exit Function

Erro_Traz_InventarioTerc_Tela:

    Traz_InventarioTerc_Tela = gErr

    Select Case gErr

        Case 119672, 119783

        Case 119784
            Call Rotina_Erro(vbOKOnly, "ERRO_INVENTARIOTERCPROD_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161976)

    End Select

    Exit Function

End Function

Private Function Traz_GridProdutos_Tela(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'traz os dados dos produtos de 3ºs p/ a tela

Dim lErro As Long
Dim iIndice As Integer
Dim objInventarioTercProd As ClassInventarioTercProd
Dim dQuantTotal As Double
            
On Error GoTo Erro_Traz_GridProdutos_Tela

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGridProdutos)
    
    'Exibe os dados da coleção na tela
    For Each objInventarioTercProd In objInventarioTerc.colInventarioTercProd

        iIndice = iIndice + 1

        'insere o produto no controle
        Produto.PromptInclude = False
        Produto.Text = objInventarioTercProd.sProduto
        Produto.PromptInclude = True
        
        'Insere o controle do produto no Grid de produtos
        GridProdutos.TextMatrix(iIndice, iGridProdutos_Produto_Col) = Produto.Text
        
        'diz qual é a linha corrente
        GridProdutos.Row = iIndice

        'busca a descricao do produto e a un. do produto
        lErro = Produto_Linha_Preenche
        If lErro <> SUCESSO Then gError 119673

        'Insere a quant no Grid de produtos
        GridProdutos.TextMatrix(iIndice, iGridProdutos_Quantidade_Col) = Formata_Estoque(objInventarioTercProd.dQuantTotal)

        'calcula a qnt total de produtos
        dQuantTotal = dQuantTotal + objInventarioTercProd.dQuantTotal
                
    Next

    'atualiza a qnt de linhas do grid
    objGridProdutos.iLinhasExistentes = iIndice

    'mosta a qnt total de itens
    QuantTotal.Caption = Formata_Estoque(dQuantTotal)
    
    Traz_GridProdutos_Tela = SUCESSO

Erro_Traz_GridProdutos_Tela:

    Traz_GridProdutos_Tela = gErr
    
    Select Case gErr

        Case 119674, 119673
        
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoFechar_Click()
'fecha a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Fechar
    
    'se tiver alteração na tela, pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 119675
    
    Unload Me
    
    Exit Sub
    
Erro_Botao_Fechar:

    Select Case gErr
    
        Case 119675
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161977)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoInvCadastrados_Click()
'chama o browser de inventarios cadastrados

Dim lErro As Long
Dim objInventarioTerc As New ClassInventarioTerc
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoInvCadastrados_Click

    lErro = Move_Tela_Memoria1(objInventarioTerc, False)
    If lErro <> SUCESSO Then gError 119785
    
    'chama a tela de browser InventarioTercProdLista
    Call Chama_Tela("InventarioTercProdLista", colSelecao, objInventarioTerc, objEventoInventario)

    Exit Sub
    
Erro_BotaoInvCadastrados_Click:

    Select Case gErr

        Case 119785
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161978)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar

    'se houve alteração, pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 119676

    'limpa a tela toda
    Call Limpa_Tela_InventarioTerc

    Exit Sub

Erro_BotaoLimpar:

    Select Case gErr

        Case 119676
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161979)
            
    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_InventarioTerc()
'limpa a tela e o grid

    'limpa a tela(text e masks) e o grid
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridProdutos)
    
    'limpa a combo e as labels
    Escaninho.ListIndex = -1
    Filial.Clear
    QuantTotal.Caption = ""
    
    'zera as variáveis de alteração
    iAlterado = 0
    giEscaninho = -1
    iClienteAlterado = 0
    iFornecedorAlterado = 0

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Escaninho_Click()
'verifica o escaninho selecionado

On Error GoTo Erro_Escaninho_Click

    'se o escaninho não foi selecionado, sai da rotina (no caso de limpar a tela)
    If Escaninho.ListIndex = -1 Then Exit Sub

    'se o escaninho selecionado foi o mesmo que estava, sai da rotina
    If giEscaninho = Escaninho.ItemData(Escaninho.ListIndex) Then Exit Sub
 
    'carrega em giEscaninho o cód. do escaninho correspondente
    giEscaninho = Escaninho.ItemData(Escaninho.ListIndex)

    Exit Sub

Erro_Escaninho_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161980)

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

Private Sub Fornecedor_Change()
    iFornecedorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
'verifica se o fornecedor selecionado é valido

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'verifica se o fornecedor foi alterado, se não foi ==> sai da rotina
    If iFornecedorAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Verifica preenchimento de Fornecedor, se não foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then
    
        'limpa a filial e sai da rotina
        Filial.Clear
        Exit Sub
        
    End If

    'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then gError 119677

    'Lê coleção de códigos, nomes de Filiais do Fornecedor
    lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
    If lErro <> SUCESSO Then gError 119678

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'verifica se foi digitado nome ou cód. do fornecedor
    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
            
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    End If

    'zera avariavel de alteração do fornecedor
    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 119677, 119678

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161981)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridProdutos(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")

   'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UNMedida.Name)
    objGridInt.colCampo.Add (Quantidade.Name)

    'Indica onde estão situadas as colunas do grid
    iGridProdutos_Produto_Col = 1
    iGridProdutos_DescricaoItem_Col = 2
    iGridProdutos_UNMedida_Col = 3
    iGridProdutos_Quantidade_Col = 4

    'passa o grid p/ o obj
    objGridInt.objGrid = GridProdutos
    
    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PRODUTOS_TERCEIROS

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'largura da 1ª coluna
    GridProdutos.ColWidth(0) = 400

    'largura Manual das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'chama a rotina que inicializa o grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridProdutos = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
'habilita/ desabilita o campo produto do grid

Dim lErro As Long
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'É p/ verificar se produto está preenchido
    sCodProduto = GridProdutos.TextMatrix(iLinha, iGridProdutos_Produto_Col)

    'formata o produto
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 119759
    
    'se o controle o grid for o do produto
    If objControl.Name = "Produto" Then

        'verifica se ele está preenchido, se sim
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
           
           'desabilita a célula do produto
           objControl.Enabled = False
           
        'senão
        Else
        
            'habilita a célula do produo
            objControl.Enabled = True
        
        End If
    
    'se for algum dos demais controles
    Else
        'verifica se o produto está preenchido, se ñ estiver
        If Len(Trim(GridProdutos.TextMatrix(iLinha, iGridProdutos_Produto_Col))) = 0 Then
            'desabilita a quantidade
            Quantidade.Enabled = False
            
        Else
            'habilita a qunt
            Quantidade.Enabled = True
        
        End If
    
    End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 119759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161982)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()
        iClienteAlterado = REGISTRO_ALTERADO
        iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate
      
    'se o cliente não foi alterado, sai da rotina
    If iClienteAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'se o cliente não foi preenchido, sai da rotina
    If Len(Trim(Cliente.Text)) = 0 Then
        'limpa a filial e sai da rotina
        Filial.Clear
        Exit Sub
    End If

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 119679

    'busca no bd a relação de filiais referentes ao cliente
    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 119680
    
    'Preenche ComboBox de Filiais do cliente
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    
    'verifica se foi digitado nome ou cód. do cliente
    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
            
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    End If
    
    iClienteAlterado = 0
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 119679, 119680
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161983)
    
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
    If Filial.ListIndex >= 0 Then Exit Sub

    'se o tipo de terc. for cliente
    If OptionTipoTerc(1).Value = True Then
    
        'verifica se o cliente foi preenchido
        If Len(Trim(Cliente.Text)) = 0 Then gError 119681

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 119683
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialCliente = New ClassFilialCliente
    
            'passa o nº preenchido como código
            objFilialCliente.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 119684
    
            'Não encontrou Filial no  BD
            If lErro = 17660 Then gError 119685
    
            'Encontrou Filial no BD, coloca no Text da Combo
            Filial.Text = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 119686
    
    'senão, é o fornecedor
    Else

        'verifica se o fornecedor foi preenchido
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 119682

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 119752
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialFornecedor = New ClassFilialFornecedor
    
            'passa o nº preenchido como código
            objFilialFornecedor.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Trim(Fornecedor.Text), objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 119753
    
            'Não encontrou Filial no  BD
            If lErro = 18272 Then gError 119754
    
            'Encontrou Filial no BD, coloca no Text da Combo
            Filial.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 119755

    End If

    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 119683, 119684, 119752, 119753

        Case 119682
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 119681
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 119685, 119686
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 119754, 119755
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Fornecedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161984)

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
        If lErro <> SUCESSO Then gError 119687

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 119687

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161985)

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
        If lErro <> SUCESSO Then gError 119688

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 119688

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161986)

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
        If lErro <> SUCESSO Then gError 119689

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case 119689

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161987)

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

Private Sub BotaoProdutos_Click()
'traz o borwser de produtos p/ a tela

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
    
    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridProdutos.Row = 0 Then gError 119690

        'carrega a string a ser passada como parametro c/ o produto do grid
        sProduto1 = GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_Produto_Col)

    End If

    'formata o produto
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 119691

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoLista_Consulta
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 119690
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 119691

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161988)

    End Select

    Exit Sub
    
End Sub

Private Sub OptionTipoTerc_Click(Index As Integer)
'verifica qual tipo de terc. foi selecionado

    'limpa a filial
    Filial.Clear
    
    'se foi cliente
    If OptionTipoTerc(1).Value = True Then
          
        'desabilita a label e o textbox do fornecedor
        Fornecedor.Visible = False
        Fornecedor.Text = ""
        FornecedorLabel.Visible = False
        
        'habilita o cliente
        Cliente.Visible = True
        ClienteLabel.Visible = True
        
    'senão
    Else
        
        'desabilita a label e o textbox do cliente
        Cliente.Visible = False
        Cliente.Text = ""
        ClienteLabel.Visible = False
        
        'habilita o fornecedor
        Fornecedor.Visible = True
        FornecedorLabel.Visible = True
    
    End If

End Sub

Private Sub GridProdutos_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)

    End If
    
End Sub

Private Sub GridProdutos_GotFocus()

    Call Grid_Recebe_Foco(objGridProdutos)

End Sub

Private Sub GridProdutos_EnterCell()

    Call Grid_Entrada_Celula(objGridProdutos, iAlterado)

End Sub

Private Sub GridProdutos_LeaveCell()

    Call Saida_Celula(objGridProdutos)

End Sub

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim dQuantLinhaAtual As Double

On Error GoTo Erro_GridProdutos_KeyDown

    'Guarda o número de linhas existentes antes do tratamento da tecla
    iLinhasExistentesAnterior = objGridProdutos.iLinhasExistentes
    
    'Guarda a quantidade da linha atual
    dQuantLinhaAtual = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_Quantidade_Col))
    
    'Faz o tratamento correspondente à tecla que foi pressionada
    Call Grid_Trata_Tecla1(KeyCode, objGridProdutos)
    
    'se a tecla pressionada foi delete e o número de linhas atuais é menor do que
    'o número de linhas antes de tratar a tecla pressionada
    If (KeyCode = vbKeyDelete) And (objGridProdutos.iLinhasExistentes < iLinhasExistentesAnterior) Then
    
        'Significa que uma linha foi exlucída e é necessário recalcular a quantidade total
        'Subtrai da quantidade total a quantidade que estava na linha excluída
        QuantTotal.Caption = Formata_Estoque(QuantTotal.Caption - dQuantLinhaAtual)
    
    End If

    Exit Sub

Erro_GridProdutos_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161989)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProdutos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProdutos, iAlterado)
    End If

End Sub

Private Sub GridProdutos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridProdutos)
End Sub

Private Sub GridProdutos_RowColChange()
    Call Grid_RowColChange(objGridProdutos)
End Sub

Private Sub GridProdutos_Scroll()
    Call Grid_Scroll(objGridProdutos)
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProdutos)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridProdutos)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProdutos)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProdutos.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridProdutos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridProdutos.Col

            Case iGridProdutos_Produto_Col
                'faz a saida da celula do produto
                lErro = Saida_Celula_Produto(objGridInt)
                If lErro <> SUCESSO Then gError 119692

            Case iGridProdutos_Quantidade_Col
                'faz a saida da celula da quantidade
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 119693

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 119694
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case 119692, 119693
        
        Case 119694
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161990)
    
    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula do produto

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
    
    'verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then
        
        'busca a validação do produto e a descricao
        lErro = Produto_Linha_Preenche
        If lErro <> SUCESSO Then gError 119695
        
    End If
            
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119696

    Saida_Celula_Produto = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr
    
    Select Case gErr
    
        Case 119695, 119696
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161991)
    
    End Select
    
    Exit Function

End Function

Private Function Produto_Linha_Preenche() As Long
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
    If lErro <> SUCESSO And lErro <> 51381 Then gError 119697
    
    'erro produto não encontrado
    If lErro = 51381 Then gError 119698

    'se o produto existe
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'retorna o produto enxuto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 119699

        'coloca o cód. do produto no controle
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridProdutos.iLinhasExistentes
            If iIndice <> GridProdutos.Row Then
                If GridProdutos.TextMatrix(iIndice, iGridProdutos_Produto_Col) = Produto.Text Then gError 119700
            End If
        Next
        
        'preenche a Descricao Produto
        GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_DescricaoItem_Col) = objProduto.sDescricao
        
        'preenche a unidade de medida
        GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_UNMedida_Col) = objProduto.sSiglaUMEstoque
        
        'se necessário, cria + uma linha
        If GridProdutos.Row - GridProdutos.FixedRows = objGridProdutos.iLinhasExistentes Then objGridProdutos.iLinhasExistentes = objGridProdutos.iLinhasExistentes + 1

    End If

    Produto_Linha_Preenche = SUCESSO

    Exit Function

Erro_Produto_Linha_Preenche:

    Produto_Linha_Preenche = gErr

    Select Case gErr

        Case 119697, 119699

        Case 119698 'produto não existente
            
            'pergunta se deseja criar novo produto
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            
            'se sim
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridProdutos)

                'chama a tela de produto
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 119700
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161992)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a saida da celula de quantidade do produto

Dim lErro As Long
Dim iLinha As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'Se a quantidade estiver preenchida
    If Len(Trim(Quantidade.Text)) > 0 Then
        
        'Critica o valor, não pode ser 0 e nem negativo
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 119701

        'coloca no grid o valor formatado
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
        
    End If
                                               
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 119702
                                
    'limpa a label do total
    QuantTotal.Caption = ""

    'p/ cada lina do grid
    For iLinha = 1 To objGridProdutos.iLinhasExistentes

        'calcula a qntd total
        dQuantTotal = dQuantTotal + StrParaDbl(GridProdutos.TextMatrix(iLinha, iGridProdutos_Quantidade_Col))

    Next

    'coloca a qnt total dos produtos na tela
    QuantTotal.Caption = Formata_Estoque(dQuantTotal)

    Saida_Celula_Quantidade = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr
    
    Select Case gErr
    
        Case 119701, 119702
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161993)
    
    End Select
    
    Exit Function

End Function

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento que traz p/ tela o item selecionado do browser

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridProdutos.Row < 1 Then Exit Sub

    'formata o produto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 119703

    'inclui no controle
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
        
    'se o foco não estiver no produto
    If Not (Me.ActiveControl Is Produto) Then
    
        'inclui no controle do grid o cód. do produto
        GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_Produto_Col) = Produto.Text
    
        'Faz o Tratamento do produto
        lErro = Produto_Linha_Preenche
        If lErro <> SUCESSO Then gError 119704
        
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
            
        Case 119703
        
        Case 119704
            'caso retorne erro, limpa a célula do produto corrente
            Produto.PromptInclude = False
            Produto.Text = ""
            Produto.PromptInclude = True
            GridProdutos.TextMatrix(GridProdutos.Row, iGridProdutos_Produto_Col) = Produto.ClipText
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161994)

    End Select

    Exit Sub

End Sub

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

Private Sub objEventoInventario_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim lErro As Long
Dim objInventarioTerc As ClassInventarioTerc

On Error GoTo Erro_objEventoInventario_evSelecao

    Set objInventarioTerc = obj1

    'traz os dados p/ a tela de acordo c/ oq foi selecionado no browser
    lErro = Traz_InventarioTerc_Tela(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119756

    Me.Show

    Exit Sub

Erro_objEventoInventario_evSelecao:

    Select Case gErr
    
        Case 119756
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161995)

    End Select
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        End If
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    'finaliza os objs
    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoInventario = Nothing
    Set objEventoProduto = Nothing
    Set objGridProdutos = Nothing
    
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

    Name = "InventarioTerc"
    
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
