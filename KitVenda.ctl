VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl KitVendaOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8085
   Begin VB.TextBox Observacao 
      Height          =   345
      Left            =   1350
      MaxLength       =   255
      TabIndex        =   5
      Top             =   1680
      Width           =   6570
   End
   Begin VB.Frame Frame7 
      Caption         =   "Componentes"
      Height          =   3600
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   2220
      Width           =   7755
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
         Left            =   210
         TabIndex        =   24
         Top             =   3150
         Width           =   1365
      End
      Begin VB.ComboBox UMComp 
         Height          =   315
         Left            =   4335
         TabIndex        =   14
         Top             =   1035
         Width           =   1410
      End
      Begin MSMask.MaskEdBox DescricaoComp 
         Height          =   315
         Left            =   3405
         TabIndex        =   15
         Top             =   1725
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoComp 
         Height          =   315
         Left            =   495
         TabIndex        =   16
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantComp 
         Height          =   315
         Left            =   2070
         TabIndex        =   17
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2550
         Left            =   180
         TabIndex        =   6
         Top             =   345
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   4498
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
   End
   Begin VB.CommandButton BotaoKits 
      Caption         =   "Kits de Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6810
      TabIndex        =   7
      Top             =   720
      Width           =   1140
   End
   Begin VB.ComboBox UM 
      Height          =   315
      Left            =   3270
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5880
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   105
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "KitVenda.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "KitVenda.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "KitVenda.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "KitVenda.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1215
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1365
      TabIndex        =   0
      Top             =   735
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   1200
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   315
      Left            =   5445
      TabIndex        =   4
      Top             =   1200
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observações:"
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
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   1725
      Width           =   1185
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   855
      TabIndex        =   22
      Top             =   1230
      Width           =   480
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2865
      TabIndex        =   21
      Top             =   735
      Width           =   3855
   End
   Begin VB.Label ProdutoLbl 
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
      Left            =   615
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   20
      Top             =   795
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "UM:"
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
      Height          =   315
      Index           =   2
      Left            =   2895
      TabIndex        =   19
      Top             =   1260
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade:"
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
      Height          =   315
      Index           =   3
      Left            =   4365
      TabIndex        =   18
      Top             =   1260
      Width           =   1035
   End
End
Attribute VB_Name = "KitVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Itens do grid
Dim objGridItens As AdmGrid
Dim iGrid_ProdutoComp_Col As Integer
Dim iGrid_DescricaoComp_Col As Integer
Dim iGrid_UMComp_Col As Integer
Dim iGrid_QuantidadeComp_Col As Integer

Dim iAlterado As Integer

'Eventos da tela
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoKits As AdmEvento
Attribute objEventoKits.VB_VarHelpID = -1
Private WithEvents objEventoComponente As AdmEvento
Attribute objEventoComponente.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Kits de Venda"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "KitVenda"

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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Preenche os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing

    Set objEventoProduto = Nothing
    Set objEventoKits = Nothing
    Set objEventoComponente = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177541)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento
    Set objEventoComponente = New AdmEvento
    Set objEventoKits = New AdmEvento

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 177445

    'Inicializa as máscaras de Produto da tela
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 177495
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoComp)
    If lErro <> SUCESSO Then gError 177496
            
    'Coloca o valor Formatado na tela
    QuantComp.Format = FORMATO_ESTOQUE
    Quantidade.Format = FORMATO_ESTOQUE
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 177445, 177495, 177496

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177542)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objKitVenda As ClassKitVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objKitVenda Is Nothing) Then

        lErro = Traz_KitVenda_Tela(objKitVenda)
        If lErro <> SUCESSO Then gError 177446

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 177446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177543)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Move_Tela_Memoria(objKitVenda As ClassKitVenda) As Long
'Move os dados da tela para a memória

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    'Coloca o Produto no formato do banco
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177513
        
    objKitVenda.sProduto = sProduto
    objKitVenda.dtData = StrParaDate(Data.Text)
    objKitVenda.sUM = UM.Text
    objKitVenda.dQuantidade = StrParaDbl(Quantidade.Text)
    objKitVenda.sObservacao = Observacao.Text
    
    Call Move_GridItens_Memoria(objKitVenda)
        
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 177513
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177544)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "KitVenda"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objKitVenda)
    If lErro <> SUCESSO Then gError 177447

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objKitVenda.sProduto, STRING_PRODUTO, "Produto"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 177447

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177545)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Tela_Preenche

    objKitVenda.sProduto = colCampoValor.Item("Produto").vValor

    If Len(Trim(objKitVenda.sProduto)) > 0 Then
        lErro = Traz_KitVenda_Tela(objKitVenda)
        If lErro <> SUCESSO Then gError 177448
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 177448

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177546)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Produto.Text)) = 0 Then gError 177449
    If Len(Trim(Data.Text)) = 0 Then gError 177524
    If Len(Trim(Quantidade.Text)) = 0 Then gError 177525
    If Len(Trim(UM.Text)) = 0 Then gError 177526

    'Se não houver pelo menos uma linha do grid preenchida, ERRO.
    If objGridItens.iLinhasExistentes <= 0 Then gError 177527
    
    'Valida o Grid de Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantidadeComp_Col)) = 0 Then gError 177528
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UMComp_Col))) = 0 Then gError 177529
    
    Next
    
    'Preenche o objKitVenda
    lErro = Move_Tela_Memoria(objKitVenda)
    If lErro <> SUCESSO Then gError 177450

    lErro = Trata_Alteracao(objKitVenda, objKitVenda.sProduto)
    If lErro <> SUCESSO Then gError 177451

    'Grava o/a KitVenda no Banco de Dados
    lErro = CF("KitVenda_Grava", objKitVenda)
    If lErro <> SUCESSO Then gError 177452

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177449
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus

        Case 177450, 177451, 177452
        
        Case 177524
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
        
        Case 177525
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDO1", gErr)
            Quantidade.SetFocus

        Case 177526
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ROTEIROSDEFABRICACAO_NAO_PREENCHIDA", gErr)
            UM.SetFocus
            
        Case 177527
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)

        Case 177528
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 177529
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177547)

    End Select

    Exit Function

End Function

Function Limpa_Tela_KitVenda() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_KitVenda

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    Descricao.Caption = ""
    
    UM.Clear
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Limpa o grid da tela
    Call Grid_Limpa(objGridItens)
    
    iAlterado = 0

    Limpa_Tela_KitVenda = SUCESSO

    Exit Function

Erro_Limpa_Tela_KitVenda:

    Limpa_Tela_KitVenda = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177548)

    End Select

    Exit Function

End Function

Function Traz_KitVenda_Tela(objKitVenda As ClassKitVenda) As Long

Dim lErro As Long
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_KitVenda_Tela

    'Lê o KitVenda que está sendo Passado
    lErro = CF("KitVenda_Le", objKitVenda)
    If lErro <> SUCESSO And lErro <> 177426 Then gError 177453

    If lErro = SUCESSO Then

        'Preenche a tela com os dados passados no obj
        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objKitVenda.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 177516
            
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        objProduto.sCodigo = objKitVenda.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177515
        
        Descricao.Caption = objProduto.sDescricao
        
        lErro = Preenche_Combo_UMs(objProduto.iClasseUM, UM, objKitVenda.sUM)
        If lErro <> SUCESSO Then gError 177518
        
        If objKitVenda.dQuantidade <> 0 Then
            Quantidade.Text = Formata_Estoque(objKitVenda.dQuantidade)
        End If

        If objKitVenda.dtData <> DATA_NULA Then
            Data.PromptInclude = False
            Data.Text = Format(objKitVenda.dtData, "dd/mm/yy")
            Data.PromptInclude = True
        End If

        Observacao.Text = objKitVenda.sObservacao
        
        lErro = Preenche_GridItens(objKitVenda)
        If lErro <> SUCESSO Then gError 177519

    End If

    iAlterado = 0

    Traz_KitVenda_Tela = SUCESSO

    Exit Function

Erro_Traz_KitVenda_Tela:

    Traz_KitVenda_Tela = gErr

    Select Case gErr

        Case 177453, 177515, 177516, 177518, 177519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177549)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 177454

    'Limpa Tela
    Call Limpa_Tela_KitVenda

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 177454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177550)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177551)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 177455

    Call Limpa_Tela_KitVenda

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 177455

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177552)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objKitVenda As New ClassKitVenda
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Produto.ClipText)) = 0 Then gError 177456

    'Coloca o Produto no formato do banco
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177522
    
    objKitVenda.sProduto = sProduto

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_KITVENDA", objKitVenda.sProduto)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("KitVenda_Exclui", objKitVenda)
        If lErro <> SUCESSO Then gError 177457

        'Limpa Tela
        Call Limpa_Tela_KitVenda

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177456
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus

        Case 177457, 177522

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177553)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_Produto_Validate

    'Se o Produto está Preenchido...
    If Len(Trim(Produto.ClipText)) <> 0 Then
            
        'Faz a Crítica do produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 177502
        
        'Se o produto não existe --> erro
        If lErro = 51381 Then gError 177503
        
        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 177504
        
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
                
        'Verifica se é produto de venda, se for --> erro
        If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 177506
        
        'Verifica se pode ser produto pai
        If objProduto.iKitVendaComp <> MARCADO Then gError 177507
        
        'Verifica se é produto gerencial, se for --> erro
        If objProduto.iGerencial <> PRODUTO_GERENCIAL Then gError 177505
        
        lErro = Preenche_Combo_UMs(objProduto.iClasseUM, UM, objProduto.sSiglaUMVenda)
        If lErro <> SUCESSO Then gError 177508
        
        'Preenche a Descrição do Produto
        Descricao.Caption = objProduto.sDescricao
                
    Else

        'Limpa a descricao
        Descricao.Caption = ""

    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 177502

        Case 177503 'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                'Limpa Descricao
                Descricao.Caption = ""

            End If
        
        Case 177504
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 177505
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_GERENCIAL", gErr, objProduto.sCodigo)
 
        Case 177507
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_PAI", gErr, objProduto.sCodigo)
            
        Case 177508
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177554)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UM_Validate

    If Len(UM.Text) > 0 Then

        lErro = Combo_Item_Igual(UM)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 177511
        
        If lErro = 12253 Then gError 177512

    End If

    Exit Sub

Erro_UM_Validate:

    Cancel = True

    Select Case gErr

        Case 177511

        Case 177512
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_UM_INEXISTENTE", gErr, UM.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177556)

    End Select

    Exit Sub

End Sub

Private Sub UM_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

       'Critica a Quantidade
       lErro = Valor_Positivo_Critica(Quantidade.Text)
       If lErro <> SUCESSO Then gError 177458

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 177458

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177557)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 177459

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 177459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177558)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 177460

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 177460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177559)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 177461

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 177461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177560)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177561)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    objKitVenda.sProduto = objProduto.sCodigo

    'Mostra os dados do KitVenda na tela
    lErro = Traz_KitVenda_Tela(objKitVenda)
    If lErro <> SUCESSO Then gError 177462

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 177462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177562)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLbl_Click()
'Browse chamado só deve exibir Produtos que PODEM ser RAIZ de KIT VENDA.

Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLbl_Click
    
    If Len(Trim(Produto.ClipText)) > 0 Then
        
        'Coloca o Produto no formato do banco
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 177497
        
        'Preenche com o Produto da tela
        objProduto.sCodigo = sProduto
        
     End If
        
    'Chama Tela ProdutoKitVendaPaiLista
    Call Chama_Tela("ProdutoKitVendaLista", colSelecao, objProduto, objEventoProduto)
        
    Exit Sub
    
Erro_ProdutoLbl_Click:

    Select Case gErr
        
        Case 177497
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177563)

    End Select
    
    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Unidade Med")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoComp.Name)
    objGrid.colCampo.Add (DescricaoComp.Name)
    objGrid.colCampo.Add (UMComp.Name)
    objGrid.colCampo.Add (QuantComp.Name)

    'Colunas do Grid
    iGrid_ProdutoComp_Col = 1
    iGrid_DescricaoComp_Col = 2
    iGrid_UMComp_Col = 3
    iGrid_QuantidadeComp_Col = 4

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_KITVENDA + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 6
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridItens, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If
End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub ProdutoComp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoComp_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ProdutoComp_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ProdutoComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ProdutoComp
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoComp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescricaoComp_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DescricaoComp_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DescricaoComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoComp
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMComp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UMComp_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UMComp_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UMComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UMComp
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantComp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QuantComp_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub QuantComp_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub QuantComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantComp
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoComp_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177530
    
    If objControl.Name = "ProdutoComp" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
        
    ElseIf objControl.Name = "DescricaoComp" Then

        objControl.Enabled = False
        
    ElseIf objControl.Name = "QuantComp" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
        
    ElseIf objControl.Name = "UMComp" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objControl.Enabled = True

            Set objProdutos = New ClassProduto

            objProdutos.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProdutos)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 177531

            lErro = Preenche_Combo_UMs(objProdutos.iClasseUM, UMComp, UMComp.Text)
            If lErro <> SUCESSO Then gError 177532
            
        Else
            objControl.Enabled = False
        End If
            
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 177530 To 177532

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177564)

    End Select

    Exit Sub

End Sub

Private Function Preenche_GridItens(objKitVenda As ClassKitVenda) As Long
'Preenche os dados do grid no obj para a tela

Dim iLinha As Integer
Dim objProdutoKitVenda As New ClassProdutoKitVenda
Dim sProdutoMascarado As String
Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_Preenche_GridItens

    'Limpa o Grid de Itens
    Call Grid_Limpa(objGridItens)

    iLinha = 0
    
    'Preenche o grid com os objetos da coleção de itens
    For Each objProdutoKitVenda In objKitVenda.colComponentes
    
        Set objProduto = New ClassProduto

        iLinha = iLinha + 1
        
        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objProdutoKitVenda.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 177520
        
        'Coloca o produto já com a máscara no controle
        ProdutoComp.PromptInclude = False
        ProdutoComp.Text = sProdutoMascarado
        ProdutoComp.PromptInclude = True
        
        objProduto.sCodigo = objProdutoKitVenda.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177521
    
        'Preenche o grid de itens com os dados
        GridItens.TextMatrix(iLinha, iGrid_ProdutoComp_Col) = ProdutoComp.Text
        GridItens.TextMatrix(iLinha, iGrid_DescricaoComp_Col) = objProduto.sDescricao
        GridItens.TextMatrix(iLinha, iGrid_QuantidadeComp_Col) = Formata_Estoque(objProdutoKitVenda.dQuantidade)
        GridItens.TextMatrix(iLinha, iGrid_UMComp_Col) = objProdutoKitVenda.sUM
        
    Next

    'Preenche com o número atual de linhas existentes no grid
    objGridItens.iLinhasExistentes = iLinha
    
    Preenche_GridItens = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 177520
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProdutoKitVenda.sProdutoKit)
            
        Case 177521
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177565)

    End Select

    Exit Function

End Function

Private Sub Move_GridItens_Memoria(objKitVenda As ClassKitVenda)
'Move os dados do GridItens para a memória

Dim iIndice As Integer
Dim objProdutoKitVenda As ClassProdutoKitVenda
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada Componente do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'inicializa o obj
        Set objProdutoKitVenda = New ClassProdutoKitVenda
        
        'Coloca o produto no formato de banco de dados
        lErro = CF("Produto_Formata", Trim(GridItens.TextMatrix(iIndice, iGrid_ProdutoComp_Col)), sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 177514
        
        'recolhe os dados do grid de Itens e adiciona na coleção
        objProdutoKitVenda.sProdutoKit = objKitVenda.sProduto
        objProdutoKitVenda.sProduto = sProduto
        objProdutoKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantidadeComp_Col))
        objProdutoKitVenda.sUM = GridItens.TextMatrix(iIndice, iGrid_UMComp_Col)
        objProdutoKitVenda.iSeq = iIndice
        
        'Adiciona o obj já Preenchedo na coleção
        objKitVenda.colComponentes.Add objProdutoKitVenda

    Next
    
    Exit Sub
    
Erro_Move_GridItens_Memoria:

    Select Case gErr
        
        Case 177514
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177566)

    End Select

    Exit Sub

End Sub

Sub BotaoKits_Click()
'Exibe os kits cadastrados no sistema

Dim objKitVenda As New ClassKitVenda
Dim colSelecao As Collection
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoKits_Click

    'Se o produto estiver preenchido na tela
    If Len(Trim(Produto.ClipText)) > 0 Then
        
        'Coloca o Produto no formato do banco
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 177498
        
        'Preenche com o Produto da tela
        objKitVenda.sProduto = sProduto
    
    End If
     
    'Chama Tela KitVendaLista
    Call Chama_Tela("KitVendaLista", colSelecao, objKitVenda, objEventoKits)
        
    Exit Sub
    
Erro_BotaoKits_Click:

    Select Case gErr
        
        Case 177498
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177567)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoKits_evSelecao(obj1 As Object)

Dim objKitVenda As ClassKitVenda
Dim lErro As Long

On Error GoTo Erro_objEventoKits_evSelecao

    Set objKitVenda = obj1
    
    'Move os dados para a tela
    lErro = Traz_KitVenda_Tela(objKitVenda)
    If lErro <> SUCESSO Then gError 177499
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoKits_evSelecao:

    Select Case gErr

        Case 177499
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177568)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoComponente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoComponente_evSelecao

    Set objProduto = obj1
    
    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 177539

    ProdutoComp.PromptInclude = False
    ProdutoComp.Text = sProduto
    ProdutoComp.PromptInclude = True

    'Incluído por Luiz Nogueira em 29/04/03
    If Not (Me.ActiveControl Is ProdutoComp) Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoComp_Col) = ProdutoComp.Text
    
        'Faz o Tratamento do produto
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 177540

    End If
    
    Me.Show

    Exit Sub
    
Erro_objEventoComponente_evSelecao:

    Select Case gErr
    
        Case 177539 To 177540
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177569)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_BotaoProdutos_Click
    
    If Me.ActiveControl Is ProdutoComp Then
        sProduto = ProdutoComp.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 177500
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoComp_Col)
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 177501
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
        
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoComponente)
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 177500
        
        Case 177501
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177570)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Combo_UMs(ByVal iClasseUM As Integer, ByVal objComboUM As ComboBox, ByVal sUM As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objClasse As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUnidadeDeMedida As ClassUnidadeDeMedida

On Error GoTo Erro_Preenche_Combo_UMs

    objClasse.iClasse = iClasseUM

    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasse, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then gError 177516

    If lErro = 22539 Then gError 177517

    objComboUM.Clear

    For Each objUnidadeDeMedida In colSiglas
        objComboUM.AddItem objUnidadeDeMedida.sSigla
    Next
    
    'Tento selecionar na Combo a Unidade anterior
    If objComboUM.ListCount <> 0 Then

        For iIndice = 0 To objComboUM.ListCount - 1

            If objComboUM.List(iIndice) = sUM Then
                objComboUM.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    Preenche_Combo_UMs = SUCESSO

    Exit Function

Erro_Preenche_Combo_UMs:

    Preenche_Combo_UMs = gErr

    Select Case gErr

        Case 177516

        Case 177517
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_INEXISTENTE", gErr, objClasse.iClasse)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177571)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
            
    'Clique em F3
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Produto Then Call ProdutoLbl_Click
        If Me.ActiveControl Is ProdutoComp Then Call BotaoProdutos_Click
            
    End If

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
       
    If lErro = SUCESSO Then
        
        'OperacaoInsumos
        If objGridInt.objGrid.Name = GridItens.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_ProdutoComp_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 177526

                Case iGrid_QuantidadeComp_Col

                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 177527

                Case iGrid_UMComp_Col

                    lErro = Saida_Celula_UM(objGridInt)
                    If lErro <> SUCESSO Then gError 177528

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 177529

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 177526 To 177528

        Case 177529
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177572)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = ProdutoComp

    If Len(Trim(ProdutoComp.ClipText)) > 0 Then
    
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 177530

    End If
       
    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoComp_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177532

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 177530, 177532
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 177531

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177573)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Saida_Celula

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", ProdutoComp.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 177533
    
    'Se o produto é gerencial  ==> Erro
    If lErro = 86295 Then gError 177538
     
    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 177534

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_ProdutoComp_Col) = Produto.Text Then gError 177536
        End If
    Next

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = 0 Then gError 177537
       
    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UMComp_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoComp_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr
                            
        Case 177533

        Case 177534
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 177535
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case 177536
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, Produto.Text, iIndice)

        Case 177537
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)
            
        Case 177538
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177574)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = QuantComp
    
    'Se o campo foi preenchido
    If Len(Trim(QuantComp.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(QuantComp.Text)
        If lErro <> SUCESSO Then gError 177535
        
        QuantComp.Text = Formata_Estoque(QuantComp.Text)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177536

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 177535, 177536
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177575)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UMComp
    
    'Se o campo foi preenchido
    If Len(Trim(UMComp.Text)) > 0 Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_UMComp_Col) = UMComp.Text
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177534

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 177534
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177576)

    End Select

    Exit Function

End Function

