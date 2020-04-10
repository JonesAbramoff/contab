VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ProdutoFornecedorOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.OptionButton OptForn 
      Caption         =   "Fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5265
      TabIndex        =   14
      Top             =   780
      Width           =   1800
   End
   Begin VB.OptionButton optProduto 
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
      Height          =   270
      Left            =   45
      TabIndex        =   13
      Top             =   780
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Acão na Gravação"
      Height          =   615
      Left            =   75
      TabIndex        =   12
      Top             =   45
      Width           =   7545
      Begin VB.OptionButton OptSobrepor 
         Caption         =   "Sobrepor associações existentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4245
         TabIndex        =   16
         Top             =   255
         Width           =   3120
      End
      Begin VB.OptionButton OptAdicionar 
         Caption         =   "Somente adicionar associações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   3810
      End
   End
   Begin MSComctlLib.TreeView Produtos 
      Height          =   4815
      Left            =   45
      TabIndex        =   10
      Top             =   1095
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   8493
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7770
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ProdutoFornecedorOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ProdutoFornecedorOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ProdutoFornecedorOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoForAssociados 
      Caption         =   "Fornecedores Associadas"
      Height          =   900
      Left            =   4080
      Picture         =   "ProdutoFornecedorOcx.ctx":080A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1095
      Width           =   1155
   End
   Begin VB.CommandButton BotaoProdAssociados 
      Caption         =   "Produtos Associados"
      Height          =   900
      Left            =   8310
      Picture         =   "ProdutoFornecedorOcx.ctx":1274
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1095
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarTodosFor 
      Caption         =   "Mar.Todos"
      Height          =   555
      Left            =   8310
      Picture         =   "ProdutoFornecedorOcx.ctx":1CDE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4695
      Width           =   1155
   End
   Begin VB.CommandButton BotaoDesTodosFor 
      Caption         =   "Des.Todos"
      Height          =   555
      Left            =   8310
      Picture         =   "ProdutoFornecedorOcx.ctx":2CF8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5340
      Width           =   1155
   End
   Begin VB.CommandButton BotaoDesTodosProd 
      Caption         =   "Des.Todos"
      Height          =   555
      Left            =   4080
      Picture         =   "ProdutoFornecedorOcx.ctx":3EDA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5355
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarTodosProd 
      Caption         =   "Mar.Todos"
      Height          =   555
      Left            =   4080
      Picture         =   "ProdutoFornecedorOcx.ctx":50BC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4710
      Width           =   1155
   End
   Begin MSComctlLib.TreeView Fornecedores 
      Height          =   4815
      Left            =   5280
      TabIndex        =   11
      Top             =   1065
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   8493
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "ProdutoFornecedorOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolProduto As Collection
Dim gcolFilialFornecedor As Collection

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim colProduto As New Collection
Dim colFilialFornecedor As New Collection
Dim objNode As Node
Dim vbMsgRet As VbMsgBoxResult
Dim iAcao As Integer
Dim iOrigem As Integer

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    iIndice = 0
    For Each objNode In Produtos.Nodes
        
        iIndice = iIndice + 1
        
        If objNode.Checked = True Then
            
            If gcolProduto.Item(iIndice).iGerencial <> PRODUTO_GERENCIAL Then
                colProduto.Add gcolProduto.Item(iIndice)
            End If
            
        End If
               
    Next

    'se nenhum produto estiver selecionado ==> erro
    If colProduto.Count = 0 Then gError 180175
    
    iIndice = 0
    For Each objNode In Fornecedores.Nodes
        
        iIndice = iIndice + 1
        
        If objNode.Checked = True Then
            
            If Not (gcolFilialFornecedor.Item(iIndice) Is Nothing) Then
                colFilialFornecedor.Add gcolFilialFornecedor.Item(iIndice)
            End If
            
        End If
               
    Next

    'se nenhum produto estiver selecionado ==> erro
    If colFilialFornecedor.Count = 0 Then gError 180176
    
    If OptAdicionar.Value = True Then
        iAcao = PRODUTOFORNECEDOR_ACAO_ADICIONAR
        vbMsgRet = vbYes
    Else
        iAcao = PRODUTOFORNECEDOR_ACAO_SOBREPOR
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_PRODUTOFORNECEDOR")
    End If
    
    If optProduto.Value = True Then
        iOrigem = PRODUTOFORNECEDOR_ORIGEM_PRODUTO
    Else
        iOrigem = PRODUTOFORNECEDOR_ORIGEM_FORNECEDOR
    End If
        
    If vbMsgRet = vbYes Then
    
        lErro = CF("ProdutoFornecedor_Grava", giFilialEmpresa, colProduto, colFilialFornecedor, iAcao, iOrigem)
        If lErro <> SUCESSO Then gError 180176
        
        'limpa a tela
        Call BotaoLimpar_Click
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 180174
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_FILIALFORN_MARCADO", gErr)
        
        Case 180175
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PRODUTO_MARCADO", gErr)
        
        Case 180176
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180177)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
    Call BotaoDesTodosFor_Click
    Call BotaoDesTodosProd_Click
    
    OptAdicionar.Value = True
    optProduto.Value = True
    
End Sub

Private Sub BotaoMarTodosProd_Click()
    Call Marca_Desmarca_TreeView(Produtos, True)
End Sub

Private Sub BotaoDesTodosProd_Click()
    Call Marca_Desmarca_TreeView(Produtos, False)
End Sub

Private Sub BotaoMarTodosFor_Click()
    Call Marca_Desmarca_TreeView(Fornecedores, True)
End Sub

Private Sub BotaoDesTodosFor_Click()
    Call Marca_Desmarca_TreeView(Fornecedores, False)
End Sub

Private Sub Produtos_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call Marca_Filhos_TreeView(Produtos, Node)
End Sub

Private Sub Fornecedores_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call Marca_Filhos_TreeView(Fornecedores, Node)
End Sub

Private Sub BotaoForAssociados_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim colFornecedoresPFF As New Collection
Dim objProduto As New ClassProduto
Dim objFornecedorPFF As ClassFornecedorProdutoFF
Dim objFilialFornecedor As ClassFilialFornecedor
    
On Error GoTo Erro_BotaoForAssociadas_Click
        
    'verifica se há algum centro de custo selecionado
    If Produtos.SelectedItem Is Nothing Then gError 180178
    
    Set objProduto = gcolProduto.Item(Produtos.SelectedItem.Index)
    
    'desmarca todos os Fornecedores
    Call Marca_Desmarca_TreeView(Fornecedores, False)
       
    lErro = CF("FornecedoresProdutoFF_Le", colFornecedoresPFF, objProduto)
    If lErro <> SUCESSO And lErro <> 63156 Then gError 180179
    
    For Each objFornecedorPFF In colFornecedoresPFF
       
        iIndice = 0
        For Each objFilialFornecedor In gcolFilialFornecedor
        
            iIndice = iIndice + 1
       
            If Not (objFilialFornecedor Is Nothing) Then
                If objFornecedorPFF.iFilialForn = objFilialFornecedor.iCodFilial And objFornecedorPFF.lFornecedor = objFilialFornecedor.lCodFornecedor Then
                    Fornecedores.Nodes.Item(iIndice).Checked = True
                    Exit For
                End If
            End If
       
        Next
        
    Next
    
    Exit Sub
    
Erro_BotaoForAssociadas_Click:

    Select Case gErr
    
        Case 180178
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_SELECIONADO", gErr)
        
        Case 180179
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180180)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProdAssociados_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim colFornecedoresPFF As New Collection
Dim objProduto As New ClassProduto
Dim objFornecedorPFF As ClassFornecedorProdutoFF
Dim objFilialFornecedor As ClassFilialFornecedor
    
On Error GoTo Erro_BotaoProdAssociadas_Click
        
    'verifica se há algum centro de custo selecionado
    If Fornecedores.SelectedItem Is Nothing Then gError 180181
    
    Set objFilialFornecedor = gcolFilialFornecedor.Item(Fornecedores.SelectedItem.Index)
    
    'Tem que clicar em uma filial e não no fornecedor
    If objFilialFornecedor Is Nothing Then gError 180182
    
    'desmarca todos os Fornecedores
    Call Marca_Desmarca_TreeView(Produtos, False)
    
    'Le os produtos ligados as fornecedores
    lErro = CF("FornecedoresProdutoFF_Le_Produto", colFornecedoresPFF, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 180194 Then gError 180183
    
    'Para cada Ligação do Produto com o Fornecedor
    For Each objFornecedorPFF In colFornecedoresPFF
       
        iIndice = 0
        For Each objProduto In gcolProduto
        
            iIndice = iIndice + 1

            If objFornecedorPFF.sProduto = objProduto.sCodigo Then
                Produtos.Nodes.Item(iIndice).Checked = True
                Exit For
            End If
       
        Next
        
    Next
    
    Exit Sub
    
Erro_BotaoProdAssociadas_Click:

    Select Case gErr
    
        Case 180181, 180182
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFOR_NAO_SELECIONADO", gErr)
        
        Case 180183
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180184)

    End Select
    
    Exit Sub

End Sub

Private Sub Produtos_DblClick()

    Call BotaoForAssociados_Click

End Sub

Private Sub Fornecedores_DblClick()

    Call BotaoProdAssociados_Click

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gcolProduto = New Collection
    Set gcolFilialFornecedor = New Collection

    Call Carrega_Arvore_Produtos
    
    Call Carrega_Arvore_Fornecedores
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180185)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolProduto = Nothing
    Set gcolFilialFornecedor = Nothing
    
End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros


    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180186)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    Parent.HelpContextID = IDH_ASSOCIACAO_CONTA_CENTRO_CUSTO_LUCRO_CONTABIL
    Set Form_Load_Ocx = Me
    Caption = "Associação Produto x Fornecedor"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ProdutoFornecedor"
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
'Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub

Private Sub Marca_Desmarca_TreeView(objTreeView As TreeView, ByVal bFlag As Boolean)

Dim objNode As Node

    'Para cada Nó da Árvore
    For Each objNode In objTreeView.Nodes
        'Replica o Flag
        objNode.Checked = bFlag
    Next
    
End Sub

Private Function Carrega_Arvore_Produtos() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colProdutos As New Collection
Dim objProduto As ClassProduto
Dim objNode As Node
Dim colIndexPai As Collection
Dim vIndexPai As Variant
Dim iNivelAnt As Integer

On Error GoTo Erro_Carrega_Arvore_Produtos

    lErro = CF("Produto_Le_Todos", colProdutos)
    If lErro <> SUCESSO Then gError 180186
    
    Set gcolProduto = colProdutos
    
    'Para cada produto
    For Each objProduto In colProdutos
    
        'Se for o nível mais alto coloca como raiz
        If objProduto.iNivel = 1 Then
        
            Set colIndexPai = New Collection
            
            'Atualiza a árvore com o Nó Pai
            Set objNode = Produtos.Nodes.Add(, tvwFirst, "X" & CStr(objProduto.sCodigo), objProduto.sCodigo & SEPARADOR & objProduto.sDescricao)
            Produtos.Nodes.Item(objNode.Index).Expanded = True
            objNode.Tag = "X" & CStr(objProduto.sCodigo)
            
            iNivelAnt = objProduto.iNivel
            
            'Adiciona a coleção de Index Pais
            colIndexPai.Add objNode.Index
            
        Else
            
            'Se o nível atual é menor ou igual ao anterior, é sinal que o ramo antigo
            'da árvore já terminou -> Tem que retirar da coleção os indices desse ramo
            For iIndice = iNivelAnt To objProduto.iNivel Step -1
                colIndexPai.Remove (iIndice)
            Next
        
            'Pega o indice do nó pai
            vIndexPai = colIndexPai.Item(objProduto.iNivel - 1)
            
            'Adiciona o Filho a árvore
            Set objNode = Produtos.Nodes.Add(vIndexPai, tvwChild, "X" & CStr(objProduto.sCodigo), objProduto.sCodigo & SEPARADOR & objProduto.sDescricao)
            objNode.Tag = "X" & CStr(objProduto.sCodigo)
        
            'Atualiza a coleção de indice
            colIndexPai.Add objNode.Index
            
            iNivelAnt = objProduto.iNivel
        
        End If
                
    Next

    Carrega_Arvore_Produtos = SUCESSO
    
    Exit Function

Erro_Carrega_Arvore_Produtos:

    Carrega_Arvore_Produtos = gErr

    Select Case gErr
    
        Case 180186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180187)
    
    End Select
    
    Exit Function

End Function

Private Function Carrega_Arvore_Fornecedores() As Long

Dim lErro As Long
Dim colFornecedores As New Collection
Dim colFiliais As New AdmColCodigoNome
Dim objFornecedor As ClassFornecedor
Dim objNodePai As Node
Dim objNode As Node
Dim objFilial As AdmCodigoNome
Dim objFilialFornecedor As ClassFilialFornecedor

On Error GoTo Erro_Carrega_Arvore_Fornecedores

    Set gcolFilialFornecedor = New Collection

    'Le todos ois fornecedores
    lErro = CF("Fornecedor_Le_Todos", colFornecedores)
    If lErro <> SUCESSO Then gError 180188
    
    'Para cada fornecedor
    For Each objFornecedor In colFornecedores
    
        'Inicializa coleção de Filiais
        Set colFiliais = New AdmColCodigoNome
    
        'Le as filiais do fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colFiliais)
        If lErro <> SUCESSO Then gError 180189
    
        'Aciona nó pai
        Set objNodePai = Fornecedores.Nodes.Add(, tvwFirst, "X" & CStr(objFornecedor.lCodigo), objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido)
        Fornecedores.Nodes.Item(objNodePai.Index).Expanded = True
        objNodePai.Tag = "X" & CStr(objFornecedor.lCodigo)
        
        gcolFilialFornecedor.Add Nothing
        
        'Para cada filial
        For Each objFilial In colFiliais
        
            Set objFilialFornecedor = New ClassFilialFornecedor
            
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            objFilialFornecedor.iCodFilial = objFilial.iCodigo
            objFilialFornecedor.sNome = objFilial.sNome

            gcolFilialFornecedor.Add objFilialFornecedor

            'Adicona o nó filho
            Set objNode = Fornecedores.Nodes.Add(objNodePai.Index, tvwChild, "X" & CStr(objFornecedor.lCodigo) & SEPARADOR & CStr(objFilial.iCodigo), objFilial.iCodigo & SEPARADOR & objFilial.sNome)
            objNode.Tag = "X" & CStr(objFornecedor.lCodigo) & SEPARADOR & CStr(objFilial.iCodigo)
            
        Next
                
    Next

    Carrega_Arvore_Fornecedores = SUCESSO
    
    Exit Function

Erro_Carrega_Arvore_Fornecedores:

    Carrega_Arvore_Fornecedores = gErr

    Select Case gErr
    
        Case 180188, 180189

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180190)
    
    End Select
    
    Exit Function

End Function

Private Function Marca_Filhos_TreeView(objTreeView As TreeView, objNode As Node)
'Marca e desmarca os filhos de um nó pai conforme sua marcação

Dim iIndice As Integer

    'Para cada nó
    For iIndice = 1 To objTreeView.Nodes.Count
        'Se o nó atual tem Pai
        If Not (objTreeView.Nodes.Item(iIndice).Parent Is Nothing) Then
            'Se o Pai do Nó atual é o no do Loop
            If objTreeView.Nodes.Item(iIndice).Parent = objNode.Text Then
                'Replica a marcação do Pai para o Filho
                objTreeView.Nodes.Item(iIndice).Checked = objNode.Checked
                'Busca os netos recursivamente
                Call Marca_Filhos_TreeView(objTreeView, objTreeView.Nodes.Item(iIndice))
            End If
        End If
    
    Next
    
End Function

