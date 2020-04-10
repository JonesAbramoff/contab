VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TabPreco 
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8070
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   1
      Left            =   180
      TabIndex        =   20
      Top             =   930
      Width           =   7590
      Begin VB.TextBox TextObservacao 
         Height          =   330
         Left            =   1740
         MaxLength       =   255
         TabIndex        =   11
         Top             =   3915
         Width           =   4980
      End
      Begin VB.TextBox TextOrigem 
         Height          =   330
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1505
         Width           =   3255
      End
      Begin VB.ComboBox ComboUFOrigem 
         Height          =   315
         ItemData        =   "TabPrecoGR.ctx":0000
         Left            =   6030
         List            =   "TabPrecoGR.ctx":0002
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1500
         Width           =   675
      End
      Begin VB.ComboBox ComboUFDestino 
         Height          =   315
         ItemData        =   "TabPrecoGR.ctx":0004
         Left            =   6030
         List            =   "TabPrecoGR.ctx":0006
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2115
         Width           =   672
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   270
         Left            =   2520
         Picture         =   "TabPrecoGR.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   330
         Width           =   300
      End
      Begin VB.TextBox TextDestino 
         Height          =   330
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2115
         Width           =   3270
      End
      Begin MSMask.MaskEdBox MaskCodigo 
         Height          =   300
         Left            =   1740
         TabIndex        =   0
         Top             =   315
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskPedagio 
         Height          =   315
         Left            =   1740
         TabIndex        =   8
         Top             =   2725
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskAdValoren 
         Height          =   315
         Left            =   1740
         TabIndex        =   10
         Top             =   3320
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskCliente 
         Height          =   330
         Left            =   1740
         TabIndex        =   3
         Top             =   900
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskDataVigencia 
         Height          =   300
         Left            =   4710
         TabIndex        =   2
         Top             =   345
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataVigencia 
         Height          =   300
         Left            =   5880
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AD - Valoren:"
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
         Left            =   480
         TabIndex        =   30
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   540
         TabIndex        =   34
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "U.F.:"
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
         Left            =   5520
         TabIndex        =   33
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "U.F.:"
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
         Left            =   5520
         TabIndex        =   32
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label LabelCliente 
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
         Left            =   975
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label LabelCodigo 
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
         Left            =   975
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   368
         Width           =   660
      End
      Begin VB.Label LabelOrigem 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   975
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label LabelDestino 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Left            =   915
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vigência:"
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
         Index           =   5
         Left            =   3120
         TabIndex        =   22
         ToolTipText     =   "Data Inicio de Vigência"
         Top             =   375
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pedágio:"
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
         Left            =   870
         TabIndex        =   21
         Top             =   2775
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   2
      Left            =   225
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   7590
      Begin MSMask.MaskEdBox MaskPreco 
         Height          =   225
         Left            =   4725
         TabIndex        =   18
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   397
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox TextDescricaoServico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1635
         MaxLength       =   50
         TabIndex        =   17
         Top             =   570
         Width           =   3960
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços"
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
         Left            =   5550
         TabIndex        =   19
         Top             =   4020
         Width           =   1965
      End
      Begin MSMask.MaskEdBox MaskServico 
         Height          =   225
         Left            =   375
         TabIndex        =   16
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTabPrecoItens 
         Height          =   3900
         Left            =   135
         TabIndex        =   9
         Top             =   105
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   6879
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   29
      Top             =   615
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tabela de Preço"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preço Serviços"
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
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   15
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   28
      Top             =   30
      Width           =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5775
      ScaleHeight     =   495
      ScaleWidth      =   2145
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   45
      Width           =   2205
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "TabPrecoGR.ctx":00F2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "TabPrecoGR.ctx":0270
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "TabPrecoGR.ctx":07A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TabPrecoGR.ctx":092C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "TabPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Definicoes do grid de servicos
Public objGridTabPrecoItens As New AdmGrid

Dim iGrid_TipoServico_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Preco_Col As Integer

'Definições dos TABS
Private Const TAB_TabelaPreco = 1
Private Const TAB_PrecoServico = 2
Private Const iCria = 1
Private Const STRING_ESTADO_SIGLA = 2

Private Const NUM_MAX_SERVICOS = 100

Public iFrameAtual As Integer
Public iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoTabPreco As AdmEvento
Attribute objEventoTabPreco.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoOrigem As AdmEvento
Attribute objEventoOrigem.VB_VarHelpID = -1
Private WithEvents objEventoDestino As AdmEvento
Attribute objEventoDestino.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(MaskCliente.Text)) > 0 Then objCliente.sNomeReduzido = MaskCliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelOrigem_Click()

Dim objOrigemDestino As New ClassOrigemDestino
Dim colSelecao As Collection

    'Preenche objOrigemDestino com o Nome da tela
    If Len(Trim(TextOrigem.Text)) > 0 Then objOrigemDestino.sOrigemDestino = TextOrigem.Text

    'Chama Tela OrigemDestino
    Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoOrigem)

End Sub

Private Sub LabelDestino_Click()

Dim objOrigemDestino As New ClassOrigemDestino
Dim colSelecao As Collection

    'Preenche objOrigemDestino com o Nome da tela
    If Len(Trim(TextDestino.Text)) > 0 Then objOrigemDestino.sOrigemDestino = TextDestino.Text

    'Chama Tela OrigemDestino
    Call Chama_Tela("OrigemDestinoLista", colSelecao, objOrigemDestino, objEventoDestino)

End Sub

Private Sub LabelCodigo_Click()

Dim objTabPreco As New ClassTabPreco
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(MaskCodigo.Text)) > 0 Then objTabPreco.lCodigo = StrParaLong(MaskCodigo.Text)
        
    'Chama Tela TabPrecoLista
    Call Chama_Tela("TabPrecoLista", colSelecao, objTabPreco, objEventoTabPreco)

End Sub

Private Sub objEventoTabPreco_evSelecao(obj1 As Object)

Dim objTabPreco As ClassTabPreco
Dim lErro As Long

On Error GoTo Erro_objEventoTabPreco_evSelecao

    Set objTabPreco = obj1

    'Move os dados para a tela
    lErro = Traz_TabPreco_Tela(objTabPreco)
    If lErro <> SUCESSO And lErro <> 96765 And lErro <> 96779 Then gError 96810
        
    If lErro = 96765 Then gError 96811
    If lErro = 96779 Then gError 96812
       
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoTabPreco_evSelecao:

    Select Case gErr

        Case 96810
            
        Case 96811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_CADASTRADA", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        
        Case 96812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECOITENS_NAO_CADASTRADA", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objCliente = obj1

    'Move o nomereduzido do cliente para a tela
    MaskCliente.Text = objCliente.sNomeReduzido
    
    Me.Show

    Exit Sub
    
Erro_objEventoCliente_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoDestino_evSelecao(obj1 As Object)

Dim objOrigemDestino As ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_objEventoDestino_evSelecao

    Set objOrigemDestino = obj1

    'Move Origem e UF para a tela
    TextDestino.Text = objOrigemDestino.sOrigemDestino
    ComboUFDestino.Text = objOrigemDestino.sUF
    
    Me.Show

    Exit Sub
    
Erro_objEventoDestino_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoOrigem_evSelecao(obj1 As Object)

Dim objOrigemDestino As ClassOrigemDestino
Dim lErro As Long

On Error GoTo Erro_objEventoOrigem_evSelecao

    Set objOrigemDestino = obj1

    'Move Origem e UF para a tela
    TextOrigem.Text = objOrigemDestino.sOrigemDestino
    ComboUFOrigem.Text = objOrigemDestino.sUF
    
    Me.Show

    Exit Sub
    
Erro_objEventoOrigem_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sServico As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridTabPrecoItens.Row < 1 Then Exit Sub
    
    MaskServico.PromptInclude = False
    MaskServico.Text = sServico
    MaskServico.PromptInclude = True
    
    'Faz o Tratamento do produto
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 96982 Then gError 96980

    If lErro = 96982 Then gError 96995

    'Verifica se o browser está sendo chamado pelo botão, se for
    'joga no grid a descrição e o produto
    If Not (Me.ActiveControl Is MaskServico) Then
        GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_TipoServico_Col) = MaskServico.Text
        GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr
    
        Case 96980
            GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_TipoServico_Col) = ""
  
        Case 96995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub
       
Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iServicoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sServico As String

On Error GoTo Erro_Traz_Produto_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", MaskServico.Text, objProduto, iServicoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 96981
    If lErro = 51381 Then gError 96982

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sServico)
    If lErro <> SUCESSO Then gError 96882
       
    MaskServico.PromptInclude = False
    MaskServico.Text = sServico
    MaskServico.PromptInclude = True

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGridTabPrecoItens.iLinhasExistentes
        If iIndice <> GridTabPrecoItens.Row Then
            If GridTabPrecoItens.TextMatrix(iIndice, iGrid_TipoServico_Col) = MaskServico.Text Then gError 96983
        End If
    Next

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 96984
    
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 96882
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sNomeReduzido)
        
        Case 96981, 96982
        
        Case 96983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE2", gErr, MaskServico.Text, iIndice)
        
        Case 96984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_TabelaPreco

    Set objEventoCliente = New AdmEvento
    Set objEventoOrigem = New AdmEvento
    Set objEventoTabPreco = New AdmEvento
    Set objEventoDestino = New AdmEvento
    Set objEventoServico = New AdmEvento

    'Inicializa o Grid de TabPrecoItens
    Call Inicializa_Grid_TabPrecoItens(objGridTabPrecoItens)
    
    lErro = Carrega_Estados()
    If lErro <> SUCESSO Then gError 96969
    
    'Inicializa a máscara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaskServico)
    If lErro <> SUCESSO Then gError 96999
    
    'Data de vigência inicializada com a Data Atual
    MaskDataVigencia.PromptInclude = False
    MaskDataVigencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataVigencia.PromptInclude = True
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 96969, 96999
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Function Carrega_Estados() As Long

Dim colCodigo As New Collection
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Carrega_Estados

    'Lê cada código da tabela Estados e poe na coleção colCodigo
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADO_SIGLA)
    If lErro <> SUCESSO Then gError 96968

    For iIndice = 1 To colCodigo.Count
        ComboUFOrigem.AddItem colCodigo(iIndice)
        ComboUFDestino.AddItem colCodigo(iIndice)
    Next
    
    Carrega_Estados = SUCESSO
    
    Exit Function
    
Erro_Carrega_Estados:

    Carrega_Estados = gErr

    Select Case gErr

        Case 96968

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function
   
Function Inicializa_Grid_TabPrecoItens(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Preço")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (MaskServico.Name)
    objGridInt.colCampo.Add (TextDescricaoServico.Name)
    objGridInt.colCampo.Add (MaskPreco.Name)

    'Colunas do Grid
    iGrid_TipoServico_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_Preco_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridTabPrecoItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_SERVICOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 14

    'Largura da primeira coluna
    GridTabPrecoItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_TabPrecoItens = SUCESSO

End Function

Public Function Trata_Parametros(Optional objTabPreco As ClassTabPreco) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma tabela de preço foi passada por parâmetro
    If Not (objTabPreco Is Nothing) Then

        'Tenta ler a tabela de Preço passada por parâmetro
        lErro = Traz_TabPreco_Tela1(objTabPreco)
        If lErro <> SUCESSO And lErro <> 96767 And lErro <> 96800 Then gError 96763

        'Se Código não está cadastrado
        If lErro = 96767 Or lErro = 96800 Then

            Call Limpa_TabPreco

            'Coloca o Código na tela
            MaskCodigo.Text = objTabPreco.lCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 96763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_TabPreco_Tela(objTabPreco As ClassTabPreco) As Long

Dim lErro As Long
Dim objTabPrecoItens As New ClassTabPrecoItens

On Error GoTo Erro_Traz_TabPreco_Tela

    Call Limpa_TabPreco

    lErro = CF("TabPreco_Le", objTabPreco)
    If lErro <> SUCESSO And lErro <> 96771 Then gError 96764
    
    'Se não existe tabela de preço com o Código passado --> Erro
    If lErro = 96771 Then gError 96765
    
    lErro = CF("TabPrecoItens_Le", objTabPreco)
    If lErro <> SUCESSO And lErro <> 96774 Then gError 96978

    'Se não existe tabela de preço de Itens com o Código passado --> Erro
    If lErro = 96774 Then gError 96779
    
    Call Carrega_Tela(objTabPreco)
    
    Call Carrega_GridTabPrecoItens(objTabPreco)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Exit Function

Erro_Traz_TabPreco_Tela:

    Traz_TabPreco_Tela = gErr

    Select Case gErr

        Case 96764, 96765, 96779, 96978

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Traz_TabPreco_Tela1(objTabPreco As ClassTabPreco) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TabPreco_Tela1

    Call Limpa_TabPreco

    If objTabPreco.dtDataVigencia <> DATA_NULA Then objTabPreco.dtDataVigencia = gdtDataAtual

    lErro = CF("TabPrecoAntData_Le", objTabPreco)
    If lErro <> SUCESSO And lErro <> 96784 Then gError 96766

    'Se não existe tabela de preço com o Código passado --> Erro
    If lErro = 96784 Then gError 96767
    
    Call Carrega_Tela(objTabPreco)
    
    lErro = CF("TabPrecoItens_Le", objTabPreco)
    If lErro <> SUCESSO And lErro <> 96774 Then gError 96980
    
    'Se não existe tabela de preço de Itens com o Código passado --> Erro
    If lErro = 96774 Then gError 96800
        
    Call Carrega_GridTabPrecoItens(objTabPreco)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Traz_TabPreco_Tela1 = SUCESSO
    
    Exit Function

Erro_Traz_TabPreco_Tela1:

    Traz_TabPreco_Tela1 = gErr

    Select Case gErr

        Case 96766, 96767, 96780, 96800

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Sub Carrega_Tela(objTabPreco As ClassTabPreco)

    MaskCodigo.Text = objTabPreco.lCodigo
    TextObservacao.Text = objTabPreco.sObservacao

    MaskDataVigencia.PromptInclude = False
    MaskDataVigencia.Text = Format(objTabPreco.dtDataVigencia, "dd/mm/yy")
    MaskDataVigencia.PromptInclude = True
    
    TextOrigem.Text = objTabPreco.iOrigem
    Call textOrigem_Validate(False)
    
    TextDestino.Text = objTabPreco.iDestino
    Call textDestino_Validate(False)
    
    MaskCliente.Text = objTabPreco.lCliente
    Call MaskCliente_Validate(False)
    
    MaskPedagio.Text = objTabPreco.dPedagio
    Call maskPedagio_Validate(False)

    MaskAdValoren.Text = objTabPreco.dAdValoren * 100
    Call maskAdValoren_Validate(False)
    
End Sub

Private Sub Carrega_GridTabPrecoItens(objTabPreco As ClassTabPreco)

Dim iLinha As Integer
Dim objTabPrecoItens As ClassTabPrecoItens
Dim sProdutoEnxuto As String
Dim lErro As Long

On Error GoTo Erro_Carrega_GridTabPrecoItens

    'Limpa o Grid de TabPrecoItens
    Call Grid_Limpa(objGridTabPrecoItens)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de TabPrecoItens
    For Each objTabPrecoItens In objTabPreco.colTabPrecoItens
    
        iLinha = iLinha + 1

        lErro = Mascara_RetornaProdutoEnxuto(objTabPrecoItens.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 96997
        
        'Mascara o produto enxuto
        MaskServico.PromptInclude = False
        MaskServico.Text = sProdutoEnxuto
        MaskServico.PromptInclude = True

        GridTabPrecoItens.TextMatrix(iLinha, iGrid_TipoServico_Col) = MaskServico.Text
        GridTabPrecoItens.TextMatrix(iLinha, iGrid_Descricao_Col) = objTabPrecoItens.sDescricao
        GridTabPrecoItens.TextMatrix(iLinha, iGrid_Preco_Col) = Format(objTabPrecoItens.dPreco, gobjFAT.sFormatoPrecoUnitario)
       
    Next
    
    objGridTabPrecoItens.iLinhasExistentes = iLinha

    Exit Sub

Erro_Carrega_GridTabPrecoItens:

    Select Case gErr

        Case 96997
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objTabPrecoItens.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Sub Limpa_TabPreco()

    'Limpa Tela
    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridTabPrecoItens)

    MaskDataVigencia.PromptInclude = False
    MaskDataVigencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataVigencia.PromptInclude = True
    
    ComboUFOrigem.Text = ""
    ComboUFDestino.Text = ""

End Sub

Public Function TP_OrigemDestino_Le(objOrigemDestino As ClassOrigemDestino, sOrigemDestino As String, sUF As String) As Long
'Lê a Origem ou o Destino com Código ou NomeRed

Dim eTipoOrigemDestino As enumTipo
Dim lErro As Long

On Error GoTo TP_OrigemDestino_Le

    eTipoOrigemDestino = Tipo_OrigemDestino(sOrigemDestino)

    Select Case eTipoOrigemDestino

    Case TIPO_STRING
        
        objOrigemDestino.sOrigemDestino = sOrigemDestino
        
        If Len(Trim(sUF)) <> 0 Then
        
            objOrigemDestino.sUF = sUF
            
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 96846
            
            'Não existe OrigemDestino com este Nome e UF
            If lErro = 96860 Then gError 96847
                        
        Else
        
            lErro = CF("OrigemDestino_Le_Nome", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96864 Then gError 96848
            
            If lErro = 96864 Then gError 96849
            
        End If
                      
    Case TIPO_CODIGO

        objOrigemDestino.iCodigo = StrParaInt(sOrigemDestino)
                
        lErro = CF("OrigemDestino_Le", objOrigemDestino)
        If lErro <> SUCESSO And lErro <> 96567 Then gError 96852
            
        If lErro = 96567 Then gError 96853
        
    Case TIPO_DECIMAL

        gError 96855

    Case TIPO_NAO_POSITIVO

        gError 96856

    End Select

    TP_OrigemDestino_Le = SUCESSO

    Exit Function

TP_OrigemDestino_Le:

    TP_OrigemDestino_Le = gErr

    Select Case gErr
        
        Case 96846, 96848, 96850, 96852 'Tratados nas rotinas chamadas
        
        Case 96847
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO_UF", objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
                
        Case 96849
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO", objOrigemDestino.sOrigemDestino)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
                
        Case 96851
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO_UF", objOrigemDestino.iCodigo, objOrigemDestino.sUF)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
                
        Case 96853
            'Envia aviso que OrigemDestino não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ORIGEMDESTINO", objOrigemDestino.iCodigo)
    
                If lErro = vbYes Then
                    'Chama tela de OrigemDestino
                    lErro = Chama_Tela("OrigemDestino", objOrigemDestino)
                End If
        
        Case 96855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sOrigemDestino)

        Case 96856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sOrigemDestino)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

Private Function Tipo_OrigemDestino(ByVal sText As String) As enumTipo

If Not IsNumeric(sText) Then
    Tipo_OrigemDestino = TIPO_STRING
ElseIf Int(CDbl(sText)) <> CDbl(sText) Then
    Tipo_OrigemDestino = TIPO_DECIMAL
ElseIf CDbl(sText) <= 0 Then
    Tipo_OrigemDestino = TIPO_NAO_POSITIVO
Else
    Tipo_OrigemDestino = TIPO_CODIGO
End If

End Function

Private Sub UpDownDataVigencia_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDownDataVigencia_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_DownClick

    'Diminui um dia em DataVigencia
    lErro = Data_Up_Down_Click(MaskDataVigencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 96844

    Exit Sub

Erro_UpDownDataVigencia_DownClick:

    Select Case gErr

        Case 96844
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVigencia_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_UpClick

    'Aumenta um dia em DataVigencia
    lErro = Data_Up_Down_Click(MaskDataVigencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 96845

    Exit Sub

Erro_UpDownDataVigencia_UpClick:

    Select Case gErr

        Case 96845
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskAdValoren_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskAdValoren_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_maskAdValoren_Validate

    If Len(Trim(MaskAdValoren.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica(MaskAdValoren.Text)
    If lErro <> SUCESSO Then gError 96802

    Exit Sub

Erro_maskAdValoren_Validate:

    Cancel = True

    Select Case gErr

        Case 96802
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ComboUFOrigem_Change()

    iAlterado = REGISTRO_ALTERADO
        
End Sub

Private Sub ComboUFOrigem_Click()

    iAlterado = REGISTRO_ALTERADO
        
End Sub

Private Sub ComboUFOrigem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_ComboUFOrigem_Validate
    
    'Se não estiver preenchido sai da Sub
    If Len(Trim(ComboUFOrigem.Text)) = 0 Then Exit Sub
    
    'Verifica se existe o ítem na ComboUFOrigem, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(ComboUFOrigem)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 96853

    'Não existe o ítem na ComboUFOrigem
    If lErro = 58583 Then gError 96854
        
    'Se o campo Origem estiver preenchido...
    If Len(Trim(TextOrigem.Text)) <> 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, TextOrigem.Text, ComboUFOrigem.Text)
        If lErro <> SUCESSO Then gError 96873
        
    End If
    
    Exit Sub

Erro_ComboUFOrigem_Validate:

    Cancel = True

    Select Case gErr

        Case 96853, 96873

        Case 96854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, ComboUFOrigem.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub

Private Sub ComboUFDestino_Change()

    iAlterado = REGISTRO_ALTERADO
        
End Sub

Private Sub ComboUFDestino_Click()

    iAlterado = REGISTRO_ALTERADO
        
End Sub

Private Sub ComboUFDestino_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_ComboUFDestino_Validate
    
    'Se não estiver preenchido sai da Sub
    If Len(Trim(ComboUFDestino.Text)) = 0 Then Exit Sub
    
    'Verifica se existe o ítem na ComboUFDestino, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(ComboUFDestino)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 96875

    'Não existe o ítem na ComboUFDestino
    If lErro = 58583 Then gError 96876
        
    'Se o campo Destino estiver preenchido...
    If Len(Trim(TextDestino.Text)) <> 0 Then
    
        lErro = TP_OrigemDestino_Le(objOrigemDestino, TextDestino.Text, ComboUFDestino.Text)
        If lErro <> SUCESSO Then gError 96877
        
    End If
    
    Exit Sub

Erro_ComboUFDestino_Validate:

    Cancel = True

    Select Case gErr

        Case 96875, 96877

        Case 96876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, ComboUFDestino.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub
      
Private Sub MaskCliente_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_MaskCliente_Validate
    
    'Se não estiver preenchido sai da Sub
    If Len(Trim(MaskCliente.Text)) = 0 Then Exit Sub

    'Faz a leitura do Cliente
    lErro = TP_Cliente_Le(MaskCliente, objCliente, iCodFilial, iCria)
    If lErro <> SUCESSO Then gError 96803
    
    'Preenche o Campo Cliente
    MaskCliente.Text = objCliente.sNomeReduzido
                   
    Exit Sub
    
Erro_MaskCliente_Validate:

    Cancel = True

    Select Case gErr

    Case 96803
                
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDataVigencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataVigencia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataVigencia_Validate

    'Verifica se a data de Vigencia foi digitada
    If Len(Trim(MaskDataVigencia.Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(MaskDataVigencia.Text)
    If lErro <> SUCESSO Then gError 96805
    
    Exit Sub

Erro_MaskDataVigencia_Validate:

    Cancel = True

    Select Case gErr

        Case 96805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate

    If Len(Trim(MaskCodigo.Text)) = 0 Then Exit Sub

    lErro = Long_Critica(MaskCodigo.Text)
    If lErro <> SUCESSO Then gError 96832

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 96832

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub TextOrigem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub textOrigem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTabPreco As New ClassTabPreco
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_TextOrigem_Validate

    'Se Origem foi alterado,
    If Len(Trim(TextOrigem.Text)) > 0 Then
        
            lErro = TP_OrigemDestino_Le(objOrigemDestino, TextOrigem.Text, ComboUFOrigem.Text)
            If lErro <> SUCESSO Then gError 96833
            
    End If
    
    TextOrigem.Text = objOrigemDestino.sOrigemDestino
    ComboUFOrigem.Text = objOrigemDestino.sUF

    Exit Sub

Erro_TextOrigem_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 96833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Sub

End Sub

Public Sub TextDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub textDestino_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTabPreco As New ClassTabPreco
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_TextDestino_Validate

    'Se Destino foi alterado,
    If Len(Trim(TextDestino.Text)) > 0 Then
        
        lErro = TP_OrigemDestino_Le(objOrigemDestino, TextDestino.Text, ComboUFDestino.Text)
        If lErro <> SUCESSO Then gError 96836
            
        TextDestino.Text = objOrigemDestino.sOrigemDestino
        ComboUFDestino.Text = objOrigemDestino.sUF
        
    End If

    Exit Sub

Erro_TextDestino_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 96836

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Sub

End Sub

Private Sub TextObservacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskPedagio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskPedagio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskPedagio_Validate

    If Len(Trim(MaskPedagio.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(MaskPedagio.Text)
    If lErro <> SUCESSO Then gError 96813

    MaskPedagio.Text = Format(MaskPedagio.Text, "STANDARD")

    Exit Sub

Erro_MaskPedagio_Validate:

    Cancel = True

    Select Case gErr

        Case 96813
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()
'Faz a Troca dos Frames visíveis em função do click nos Tabs correspondentes

    'Se frame atual não corresponde ao tab selecionado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True

        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If
    
End Sub

Public Sub GridTabPrecoItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridTabPrecoItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTabPrecoItens, iAlterado)
    End If

End Sub

Public Sub GridTabPrecoItens_EnterCell()

    Call Grid_Entrada_Celula(objGridTabPrecoItens, iAlterado)

End Sub

Public Sub GridTabPrecoItens_GotFocus()

    Call Grid_Recebe_Foco(objGridTabPrecoItens)

End Sub

Public Sub GridTabPrecoItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTabPrecoItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTabPrecoItens, iAlterado)
    End If

End Sub

Public Sub GridTabPrecoItens_LeaveCell()

    Call Saida_Celula(objGridTabPrecoItens)

End Sub

Public Sub GridTabPrecoItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridTabPrecoItens)

End Sub

Public Sub GridTabPrecoItens_RowColChange()

    Call Grid_RowColChange(objGridTabPrecoItens)

End Sub

Public Sub GridTabPrecoItens_Scroll()

    Call Grid_Scroll(objGridTabPrecoItens)

End Sub

Private Sub GridTabPrecoItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridTabPrecoItens)

End Sub

Private Sub maskServico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskServico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTabPrecoItens)

End Sub

Private Sub MaskServico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTabPrecoItens)
    
End Sub

Private Sub maskServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTabPrecoItens.objControle = MaskServico
    lErro = Grid_Campo_Libera_Foco(objGridTabPrecoItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub MaskPreco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskPreco_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTabPrecoItens)

End Sub

Private Sub maskPreco_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTabPrecoItens)
    
End Sub

Private Sub maskPreco_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTabPrecoItens.objControle = MaskPreco
    lErro = Grid_Campo_Libera_Foco(objGridTabPrecoItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96814

    Select Case objGridInt.objGrid.Col

        Case iGrid_TipoServico_Col
            lErro = Saida_Celula_MaskServico(objGridInt)
            If lErro <> SUCESSO Then gError 96815

        Case iGrid_Preco_Col
            lErro = Saida_Celula_MaskPreco(objGridInt)
            If lErro <> SUCESSO Then gError 96816

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96817

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 96814 To 96817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskServico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_MaskServico

    Set objGridInt.objControle = MaskServico
    
    If Len(Trim(MaskServico.ClipText)) > 0 Then
    
        'Faz o Tratamento do produto
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO And lErro <> 96982 Then gError 96819
    
        If lErro = 96982 Then gError 96996
    
        If GridTabPrecoItens.Row - GridTabPrecoItens.FixedRows = objGridTabPrecoItens.iLinhasExistentes Then
            objGridTabPrecoItens.iLinhasExistentes = objGridTabPrecoItens.iLinhasExistentes + 1
        End If
    
        GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    
    Else
        
        GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_Descricao_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96824

    Saida_Celula_MaskServico = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskServico:

    Saida_Celula_MaskServico = gErr

    Select Case gErr

        Case 96819, 96824
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 96996
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MaskServico.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskPreco(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskPreco

    Set objGridInt.objControle = MaskPreco

    If Len(Trim(MaskPreco.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(MaskPreco.Text)
        If lErro <> SUCESSO Then gError 96829

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96830

    Saida_Celula_MaskPreco = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskPreco:

    Saida_Celula_MaskPreco = gErr

    Select Case gErr

        Case 96829, 96830
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'gera um novo número de tabela de preço automaticamente

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_TABPRECO", "TabPrecoGR", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 96839

    MaskCodigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 96839

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 96879

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 96879

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Chama a função de gravação e limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 96880

    'Limpa a Tela
    Call Limpa_TabPreco
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 96880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'pergunta se o usuário deseja salvar as alterações e limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 96881

    'Limpa a Tela
    Call Limpa_TabPreco
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 96881

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sServico As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sServico1 As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is MaskServico Then
    
        sServico1 = MaskServico.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridTabPrecoItens.Row = 0 Then gError 96883

        sServico1 = GridTabPrecoItens.TextMatrix(GridTabPrecoItens.Row, iGrid_TipoServico_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sServico1, sServico, iPreenchido)
    If lErro <> SUCESSO Then gError 96884
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sServico = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sServico
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico)

    Exit Sub
        
Erro_BotaoServicos_Click:
    
    Select Case gErr
        
        Case 96883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 96884 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui a tabela de preço do código passado

Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult
Dim lCodigo As Long
Dim objTabPreco As New ClassTabPreco

On Error GoTo Erro_BotaoExcluir_Click
      
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos Código e Data Vigencia foram informados, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96899
    If Len(Trim(MaskDataVigencia.ClipText)) = 0 Then gError 96900
        
    objTabPreco.lCodigo = StrParaLong(MaskCodigo.Text)
    objTabPreco.dtDataVigencia = StrParaDate(MaskDataVigencia.Text)
    
    lErro = CF("TabPrecoItens_Le1", objTabPreco)
    If lErro <> SUCESSO And lErro <> 98005 Then gError 96901
    
    'Se não está cadastrado --> Erro
    If lErro = 98005 Then gError 96902

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TABPRECO", objTabPreco.lCodigo, objTabPreco.dtDataVigencia)

    If vbMsgRet = vbYes Then

        'exclui a Tabela de Preço
        lErro = CF("TabPreco_Exclui", objTabPreco)
        If lErro <> SUCESSO Then gError 96904
        
        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
        
        'Limpa a Tela
        Call Limpa_TabPreco

        iAlterado = 0

    End If
    
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
                
        Case 96899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_NAO_PREENCHIDA", gErr)

        Case 96901, 96904
        
        Case 96902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGODATA_NAO_ENCONTRADO", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
   
    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Chama as funções de recolhimento de dados da tela e Gravação

Dim objTabPreco As New ClassTabPreco
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios foram informados, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96885
    If Len(Trim(MaskDataVigencia.ClipText)) = 0 Then gError 96886
    If Len(Trim(MaskCliente.Text)) = 0 Then gError 96887
    If Len(Trim(TextOrigem.Text)) = 0 Then gError 96888
    If Len(Trim(TextDestino.Text)) = 0 Then gError 96889
    If Len(Trim(ComboUFOrigem.Text)) = 0 Then gError 96890
    If Len(Trim(ComboUFDestino.Text)) = 0 Then gError 96891
    If Len(Trim(MaskPedagio.Text)) = 0 Then gError 96892
    If Len(Trim(MaskAdValoren.Text)) = 0 Then gError 96893
            
    'Se não houver pelo menos uma linha do grid preenchida, ERRO.
    If objGridTabPrecoItens.iLinhasExistentes <= 0 Then gError 96894

    'Se o Preço não está preenchido
    For iIndice = 1 To objGridTabPrecoItens.iLinhasExistentes
        If Len(Trim(GridTabPrecoItens.TextMatrix(iIndice, iGrid_Preco_Col))) = 0 Then gError 96977
        If Len(Trim(GridTabPrecoItens.TextMatrix(iIndice, iGrid_TipoServico_Col))) = 0 Then gError 98000
    Next
    
    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objTabPreco)
    If lErro <> SUCESSO Then gError 96895
   
    'Verifica se o Código já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objTabPreco, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> SUCESSO Then gError 96896

    'Grava no BD os dados da Tela
    lErro = CF("TabPreco_Grava", objTabPreco)
    If lErro <> SUCESSO Then gError 96897

     'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_NAO_PREENCHIDA", gErr)

        Case 96887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 96888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", gErr)
            
        Case 96889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_PREENCHIDO", gErr)

        Case 96890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UFORIGEM_NAO_PREENCHIDO", gErr)

        Case 96891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UFDESTINO_NAO_PREENCHIDO", gErr)

        Case 96892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDAGIO_NAO_PREENCHIDO", gErr)

        Case 96893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ADVALOREN_NAO_PREENCHIDO", gErr)
        
        Case 96894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
                
        Case 96895, 96896, 96897

        Case 96977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRECO_NAO_PREENCHIDO", gErr, iIndice)
                
        Case 98000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr, iIndice)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Function Move_Tela_Memoria(objTabPreco As ClassTabPreco) As Long
'Move os dados da tela para a memória

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados do Tab Principal para a memória
    objTabPreco.lCodigo = StrParaLong(MaskCodigo.Text)
    objTabPreco.dAdValoren = PercentParaDbl(MaskAdValoren.FormattedText)
    objTabPreco.dPedagio = StrParaDbl(MaskPedagio.Text)
    objTabPreco.dtDataVigencia = StrParaDate(MaskDataVigencia.Text)
    objTabPreco.sObservacao = TextObservacao.Text
    
    If Len(Trim(MaskCliente.Text)) > 0 Then
        'Faz a leitura do Cliente
        objCliente.sNomeReduzido = MaskCliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 96970
        
        If lErro = 12348 Then gError 96971
        objTabPreco.lCliente = objCliente.lCodigo
    End If
        
    If Len(Trim(TextOrigem.Text)) > 0 Then
                
        objOrigemDestino.sOrigemDestino = TextOrigem.Text
        If Len(Trim(ComboUFOrigem.Text)) <> 0 Then
            
            objOrigemDestino.sUF = ComboUFOrigem.Text
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 98005
        
            If lErro = 96860 Then gError 98006
            
        Else
        
            lErro = CF("OrigemDestino_Le_Nome", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96864 Then gError 96972
        
            If lErro = 96864 Then gError 96973
                        
        End If
        objTabPreco.iOrigem = objOrigemDestino.iCodigo
    End If
        
    If Len(Trim(TextDestino.Text)) > 0 Then
        objOrigemDestino.sOrigemDestino = TextDestino.Text
        If Len(Trim(ComboUFDestino.Text)) <> 0 Then
            
            objOrigemDestino.sUF = ComboUFDestino.Text
            lErro = CF("OrigemDestino_Le_NomeUF", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96860 Then gError 98007
        
            If lErro = 96860 Then gError 98008
            
        Else
               
            lErro = CF("OrigemDestino_Le_Nome", objOrigemDestino)
            If lErro <> SUCESSO And lErro <> 96864 Then gError 96974
        
            If lErro = 96864 Then gError 96975
            
        End If
        objTabPreco.iDestino = objOrigemDestino.iCodigo
    End If
    
    'Move os dados do GridTabPrecoItens para a Memória
    Call Move_GridTabPrecoItens_Memoria(objTabPreco)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 96970, 96972, 96974, 98005, 98007
        
        Case 96971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, MaskCliente.Text)
            
        Case 96973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_EXISTENTE", gErr, TextOrigem.Text)
            
        Case 96975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_EXISTENTE", gErr, TextDestino.Text)
            
        Case 98006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_EXISTENTE1", gErr, TextOrigem.Text, ComboUFOrigem.Text)
            
        Case 98008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_EXISTENTE1", gErr, TextDestino.Text, ComboUFDestino.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub Move_GridTabPrecoItens_Memoria(objTabPreco As ClassTabPreco)
'Move os dados do GridTabPrecoItens para a memória

Dim iIndice As Integer
Dim objTabPrecoItens As ClassTabPrecoItens
Dim sServico As String
Dim iServicoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridTabPrecoItens_Memoria

    'Para cada Serviço do grid
    For iIndice = 1 To objGridTabPrecoItens.iLinhasExistentes

        Set objTabPrecoItens = New ClassTabPrecoItens
        
        Call CF("Produto_Formata", Trim(GridTabPrecoItens.TextMatrix(iIndice, iGrid_TipoServico_Col)), sServico, iServicoPreenchido)
        If lErro <> SUCESSO Then gError 98305
        
        'recolhe os dados do grid de Serviços e adiciona na coleção
        objTabPrecoItens.sProduto = sServico
        objTabPrecoItens.sDescricao = Trim(GridTabPrecoItens.TextMatrix(iIndice, iGrid_Descricao_Col))
        objTabPrecoItens.dPreco = StrParaDbl(GridTabPrecoItens.TextMatrix(iIndice, iGrid_Preco_Col))
        objTabPrecoItens.dtDataVigencia = objTabPreco.dtDataVigencia
        objTabPrecoItens.lCodTabela = objTabPreco.lCodigo
        'Adiciona o obj já carregado na coleção
        objTabPreco.colTabPrecoItens.Add objTabPrecoItens

    Next
    
    Exit Sub
    
Erro_Move_GridTabPrecoItens_Memoria:

    Select Case gErr
        
        Case 98305
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTabPreco As New ClassTabPreco

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TabPrecoGR"

    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objTabPreco)
    If lErro <> SUCESSO Then gError 96985
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objTabPreco.lCodigo, 0, "Codigo"
    colCampoValor.Add "DataVigencia", objTabPreco.dtDataVigencia, 0, "DataVigencia"
    colCampoValor.Add "Cliente", objTabPreco.lCliente, 0, "Cliente"
    colCampoValor.Add "Origem", objTabPreco.iOrigem, 0, "Origem"
    colCampoValor.Add "Destino", objTabPreco.iDestino, 0, "Destino"
    colCampoValor.Add "Pedagio", objTabPreco.dPedagio, 0, "Pedagio"
    colCampoValor.Add "ADValoren", objTabPreco.dAdValoren, 0, "ADValoren"
    colCampoValor.Add "Observacao", objTabPreco.sObservacao, STRING_TABPRECO_OBSERVACAO, "Observacao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        
        Case 96985
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTabPreco As New ClassTabPreco

On Error GoTo Erro_Tela_Preenche

    objTabPreco.lCodigo = colCampoValor.Item("Codigo").vValor
    objTabPreco.dtDataVigencia = colCampoValor.Item("DataVigencia").vValor
    objTabPreco.lCliente = colCampoValor.Item("Cliente").vValor
    objTabPreco.iOrigem = colCampoValor.Item("Origem").vValor
    objTabPreco.iDestino = colCampoValor.Item("Destino").vValor
    objTabPreco.dPedagio = colCampoValor.Item("Pedagio").vValor
    objTabPreco.dAdValoren = colCampoValor.Item("ADValoren").vValor
    objTabPreco.sObservacao = colCampoValor.Item("Observacao").vValor
    
    If objTabPreco.lCodigo <> 0 Then

        'Move os dados para a tela
        lErro = Traz_TabPreco_Tela(objTabPreco)
        If lErro <> SUCESSO And lErro <> 96765 And lErro <> 96779 Then gError 96841
        
        If lErro = 96765 Then gError 96842
        
        If lErro = 96779 Then gError 96843
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 96841
        
        Case 96842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_CADASTRADO", gErr, objTabPreco.lCodigo)
    
        Case 96843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITENS_NAO_CADASTRADO", gErr, objTabPreco.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais
    Set objEventoTabPreco = Nothing
    Set objEventoCliente = Nothing
    Set objEventoOrigem = Nothing
    Set objEventoDestino = Nothing
    Set objEventoServico = Nothing

    'Libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
            
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is TextOrigem Then Call LabelOrigem_Click
            If Me.ActiveControl Is TextDestino Then Call LabelDestino_Click
            If Me.ActiveControl Is MaskCliente Then Call LabelCliente_Click
            If Me.ActiveControl Is MaskCodigo Then Call LabelCodigo_Click
            If Me.ActiveControl Is MaskServico Then Call BotaoServicos_Click
                    
    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Tabela de Preço"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "TabPreco"

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
'    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

