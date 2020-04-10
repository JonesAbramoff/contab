VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ProdutoEmbalagem 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   KeyPreview      =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9930
   Begin VB.Frame Frame2 
      Caption         =   "Produto x Embalagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   9495
      Begin VB.CheckBox FixarGrid 
         Caption         =   "Fixar Conteúdo do Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   6645
         TabIndex        =   24
         Top             =   252
         Width           =   2610
      End
      Begin VB.ComboBox UMPeso 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1620
         Width           =   975
      End
      Begin VB.ComboBox UMEmbalagem 
         Height          =   288
         Left            =   6912
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1044
         Width           =   975
      End
      Begin VB.TextBox PesoEmbalagem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   7020
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox PesoBruto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8160
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox PesoLiqTotal 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5640
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton BotaoEmbalagens 
         Height          =   540
         Left            =   7560
         Picture         =   "ProdutoEmbalagem.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2952
         Width           =   1815
      End
      Begin VB.TextBox Capacidade 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8112
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox NomeProdEmb 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   2364
      End
      Begin VB.TextBox Embalagem 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.OptionButton EmbalagemPadrao 
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   825
      End
      Begin MSFlexGridLib.MSFlexGrid GridEmbalagens 
         Height          =   2292
         Left            =   120
         TabIndex        =   2
         Top             =   576
         Width           =   9252
         _ExtentX        =   16325
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   9495
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   855
         Width           =   5520
      End
      Begin VB.Label LabelUMEstoque 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4920
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "U.M. Estoque:"
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
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label LabelProduto 
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
         Left            =   630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   390
         Width           =   735
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
         Height          =   195
         Left            =   420
         TabIndex        =   19
         Top             =   885
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7560
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ProdutoEmbalagem.ctx":11EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ProdutoEmbalagem.ctx":1344
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ProdutoEmbalagem.ctx":14CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ProdutoEmbalagem.ctx":1A00
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "ProdutoEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variável de controle de alteração nos campos da tela
Dim iAlterado As Integer

'Variável de controle de alteração no campo produto
Dim iProdutoAlterado As Integer

'Obj para utilização do grid
Dim objGrid As AdmGrid

'Variável para guardar o nome reduzido do produto
Dim gsNomeRedProd As String

'Variável para guardar a classe UM do produto
Dim giClasseUM As Integer

'Declaração utilizada para evento LabelProduto_Click
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Declaração utilizada para evento BotaoEmbalagens_Click
Private WithEvents objEventoEmbalagens As AdmEvento
Attribute objEventoEmbalagens.VB_VarHelpID = -1

'Variáveis que indicam os índices de cada coluna do grid
Dim iGridEmbalagens_EmbalagemPadrao_Col As Integer
Dim iGridEmbalagens_Embalagem_Col As Integer
Dim iGridEmbalagens_NomeProdEmb_Col As Integer
Dim iGridEmbalagens_UMEmbalagem_Col As Integer
Dim iGridEmbalagens_Capacidade_Col As Integer
Dim iGridEmbalagens_UMPeso_Col As Integer
Dim iGridEmbalagens_PesoLiqTotal_Col As Integer
Dim iGridEmbalagens_PesoEmbalagem_Col As Integer
Dim iGridEmbalagens_PesoBruto_Col As Integer

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    '*** Fernando, Criar IDH
    'Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Produto x Embalagens"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ProdutoEmbalagem"

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

Private Sub BotaoEmbalagens_Click()

Dim objEmbalagem As New ClassEmbalagem
Dim colSelecao As New Collection
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoEmbalagens_Click

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95121

    'Verifica se há uma linha selecionada
    If GridEmbalagens.Row = 0 Then gError 95108
        
    'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
    faz uma nova leitura em busca do codigo
    objEmbalagem.sSigla = GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_Embalagem_Col)
    lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 95088 Then gError 95468
    
    'chama a tela de browser
    Call Chama_Tela("EmbalagensLista", colSelecao, objEmbalagem, objEventoEmbalagens)
    
    iAlterado = 1
    
    Exit Sub
    
Erro_BotaoEmbalagens_Click:
    
    Select Case gErr
    
        Case 95468
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case 95469
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM1", objEmbalagem.sSigla)
            
            If vbMsgRes = vbYes Then
                'Chama a tela de Embalagens
                Call Chama_Tela("Embalagem", objEmbalagem)
            Else
                Embalagem.SetFocus
            End If
        
        Case 95108
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 95121
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165608)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoEmbalagens_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEmbalagem As ClassEmbalagem
Dim iLinha As Integer
Dim sSiglaAux As String
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoEmbalagens_evSelecao
           
    'Verifica se a embalagem da linha atual está preenchida
    If Len(Trim(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_Embalagem_Col))) = 0 Then
            
        'Define o tipo de obj recebido (Tipo Embalagem)
        Set objEmbalagem = obj1
        
        'Verifica se há alguma embalagem repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridEmbalagens.Row Then
                
                If GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col) = objEmbalagem.sSigla Then
                    Embalagem.Text = ""
                    gError 95120
                End If
                    
            End If
                           
        Next
        
        If ActiveControl Is Embalagem Then
            Embalagem.Text = objEmbalagem.sSigla
        Else
        
            'Preenche o grid
            GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_Embalagem_Col) = objEmbalagem.sSigla
            GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_NomeProdEmb_Col) = Mid(gsNomeRedProd & " " & objEmbalagem.sDescricao, 1, 20)
            GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col) = objEmbalagem.dPeso
            
            For iIndice = 0 To UMPeso.ListCount - 1
            
                If UCase(UMPeso.List(iIndice)) = "KG" Then Exit For
            
            Next
        
            'UM Default da embalagem
            GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMPeso_Col) = UMPeso.List(iIndice)
        
            'Cria mais uma linha no grid
            If GridEmbalagens.Row - GridEmbalagens.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoEmbalagens_evSelecao:

    Select Case gErr

        Case 95464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case 95120
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_REPETIDA", gErr, objEmbalagem.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGrid)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165609)
              
    End Select
    
    Exit Sub
        
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    'Traz para a tela
    lErro = Traz_ProdutoEmbalagem_Tela(objProduto)
    If lErro <> SUCESSO Then gError 95106
    
    Me.Show
    
    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case 95106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165610)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Se o produto está preenchido...
    If (Len(Trim(Produto.ClipText)) > 0) Then
        
        'Formata o código do produto para o BD
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 95012
        
        'guarda o código já formatado em objProduto
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'chama a tela de browser
    Call Chama_Tela("ProdutosSubstLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub
    
Erro_LabelProduto_Click:

    Select Case gErr
    
        Case 95012
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165611)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = 1
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim vbMsgRes As VbMsgBoxResult
Dim iLinha As Integer

On Error GoTo Erro_Produto_Validate

    'Se nao houve alteracao
    If iProdutoAlterado = 0 Then Exit Sub
    
    'Limpa os campos
    LabelUMEstoque.Caption = ""
    Descricao.Caption = ""
    
    'Inicializa o NomeReduzido do Produto com vazio
    gsNomeRedProd = ""
    
    'Se nao estiver marcado => Limpa o Grid
    If FixarGrid.Value = vbUnchecked Then
        
        'Se existir linha válida ...
        If objGrid.iLinhasExistentes > 0 Then
            '... Limpa o grid
            lErro = Grid_Limpa(objGrid)
            If lErro <> SUCESSO Then gError 95065
        End If
        
    End If

    'Verifica se o produto esta preeenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
                
        'Critica o formato do produto e se existe no BD
        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 95005
        
        'Se nao encontrou o produto => erro
        If lErro = 25041 Then gError 95010
        
        'O Produto é Gerencial
        If lErro = 25043 Then gError 95011
        
        objProdutoFilial.sProduto = objProduto.sCodigo
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Le o produto para a giFilialEmpresa
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 95006
        
        'se nao encontrou para essa filial => erro
        If lErro = 28261 Then gError 95007
        
        If giClasseUM <> objProduto.iClasseUM Then
            For iLinha = 1 To objGrid.iLinhasExistentes
                    GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMEmbalagem_Col) = " "
            Next
        End If
        
        'Exibe os dados na tela
        Produto.PromptInclude = False
        Produto.Text = objProduto.sCodigo
        Produto.PromptInclude = True
        Descricao.Caption = objProduto.sDescricao
        LabelUMEstoque.Caption = objProduto.sSiglaUMEstoque
        gsNomeRedProd = objProduto.sNomeReduzido
        giClasseUM = objProduto.iClasseUM
               
        'Carrega a combo
        lErro = Carrega_UMEmbalagem(objProduto)
        If lErro <> SUCESSO Then gError 95050
        
        'Le as embalagens já relacionadas com o produto
        lErro = CF("ProdutoEmbalagem_Le_Produto", objProduto)
        If lErro <> SUCESSO And lErro <> 95016 Then gError 95008
        
        If lErro = SUCESSO Then
            
            'Se existir linhas válidas no grid ...
            If objGrid.iLinhasExistentes <> 0 Then
                
                '... Avisa que nao foi possivel fixar o conteudo do grid pois há emb. associadas
                Call Rotina_Aviso(vbOKOnly, "AVISO_FIXAR_GRIDEMB_NAO_POSSIVEL")
                
            End If
            
            'Desmarca a opcao de fixar o grid
            FixarGrid.Value = vbUnchecked
            
            'Limpa o grid
            lErro = Grid_Limpa(objGrid)
            If lErro <> SUCESSO Then gError 95144
            
            'Carrega o grid com os dados do Obj
            lErro = Preenche_GridEmbalagens(objProduto)
            If lErro <> SUCESSO Then gError 95009
            
        End If
        
        iAlterado = 0
        iProdutoAlterado = 0
        
    End If
        
    Exit Sub
        
Erro_Produto_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 95005, 95006, 95009, 95008, 95050, 95144
        
        Case 95011
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
                    
        Case 95010
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Descricao.Caption = ""
                LabelUMEstoque.Caption = ""
                gsNomeRedProd = ""
                Call Limpa_Tela_ProdutoEmbalagem
                Produto.SetFocus
            End If
        
        Case 95007
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objProduto.sCodigo, giFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165612)
            
    End Select
    
    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa a máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 95003
    
    'Inicializa o grid
    Set objGrid = New AdmGrid
    Call GridEmbalagens_Inicializa(objGrid)
    
    'Inicializa o objEventoProduto
    Set objEventoProduto = New AdmEvento
    
    'Inicializa o objEventoEmbalagens
    Set objEventoEmbalagens = New AdmEvento

    'Le a unidade de medida
    lErro = Carrega_UMPeso()
    If lErro <> SUCESSO Then gError 95038
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
    
        Case 95038
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165613)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objProduto Is Nothing Then
        lErro = Traz_ProdutoEmbalagem_Tela(objProduto)
        If lErro <> SUCESSO Then gError 95004
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 95004
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165614)
    End Select
    
    Exit Function

End Function

Private Sub EmbalagemPadrao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub EmbalagemPadrao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub EmbalagemPadrao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub EmbalagemPadrao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = EmbalagemPadrao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub Embalagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Embalagem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Embalagem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Embalagem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Embalagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Embalagem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomeProdEmb_Change()
    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub NomeProdEmb_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub NomeProdEmb_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub NomeProdEmb_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub NomeProdEmb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = NomeProdEmb
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMEmbalagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub UMEmbalagem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub UMEmbalagem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub UMEmbalagem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub UMEmbalagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UMEmbalagem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Capacidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Capacidade_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Capacidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Capacidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Capacidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Capacidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub UMPeso_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub UMPeso_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub UMPeso_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub UMPeso_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub UMPeso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UMPeso
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PesoLiqTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoLiqTotal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoLiqTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub PesoLiqTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub PesoLiqTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PesoLiqTotal
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PesoEmbalagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoEmbalagem_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoEmbalagem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub PesoEmbalagem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub PesoEmbalagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PesoEmbalagem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub PesoBruto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoBruto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoBruto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub PesoBruto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PesoBruto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Gravar o registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 95018
    
    'Limpa a tela
    Call Limpa_Tela_ProdutoEmbalagem
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
        
        Case 95018
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165615)
            
    End Select
    
    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProdutoEmbalagem As ClassProdutoEmbalagem

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
       
    'Critica os dados da tela
    lErro = ProdutoEmbalagem_Critica()
    If lErro <> SUCESSO Then gError 95089
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then gError 95021
    
    'Verifica se o registro já existe e confirma se deseja alterar
    lErro = Trata_Alteracao(objProdutoEmbalagem, objProduto.sCodigo)
    If lErro <> SUCESSO Then gError 95023
    
    'Grava
    lErro = CF("ProdutoEmbalagem_Grava", objProduto)
    If lErro <> SUCESSO Then gError 95022

    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Select Case gErr
        
        Case 95021, 95022, 95023, 95089
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165616)
            
    End Select
            
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95033
    
    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PRODUTOEMBALAGEM", objProduto.sCodigo)

    'Se a resposta for não
    If vbMsgRes = vbNo Then
        
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    
    End If
    
    'Formata o código do produto para o BD
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 95113
    
    'guarda o código já formatado em objProduto
    objProduto.sCodigo = sProdutoFormatado
    
    'Confirmou a exclusao, logo...
    lErro = CF("ProdutoEmbalagem_Exclui", objProduto)
    If lErro <> SUCESSO Then gError 95036
    
    'limpa a tela
    lErro = Limpa_Tela_ProdutoEmbalagem
    If lErro <> SUCESSO Then gError 95051
    
    iAlterado = 0
    iProdutoAlterado = 0
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoExcluir_Click:

    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
        
    Select Case gErr
        
        Case 95033
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
                
        Case 95034, 95036, 95051, 95113
                
        Case 95035
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOEMBALAGEM_INEXISTENTE", gErr, objProduto.sCodigo)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165617)
            
    End Select
    
    Exit Sub

End Sub

Function Move_Tela_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'se estiver preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Formata o código do produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 95040
        
        'Move para o objProduto o código do produto no formato do BD
        objProduto.sCodigo = sProdutoFormatado
        
        'Move os dados do grid para o ObjProduto
        lErro = Move_GridEmbalagens_Memoria(objProduto)
        If lErro <> SUCESSO Then gError 95041
    
    End If

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 95040, 95041
        
        Case 95039
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165618)
            
    End Select

    Exit Function

End Function

Function Move_GridEmbalagens_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutoEmb As ClassProdutoEmbalagem
Dim objEmbalagem As New ClassEmbalagem
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Move_GridEmbalagens_Memoria

    'Para cada linha do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        'Verifica se a Embalagem foi preenchida
        If Len(Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_Embalagem_Col))) <> 0 Then
                        
            'Inicializa o Obj
            Set objProdutoEmb = New ClassProdutoEmbalagem
            Set objEmbalagem = New ClassEmbalagem
            
            'Se for a embalagem padrao => guarda a constante
            If StrParaInt(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_EmbalagemPadrao_Col)) = vbChecked Then objProdutoEmb.iPadrao = PRODUTOEMBALAGEM_PADRAO
            
            'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
            faz uma nova leitura em busca do codigo
            objEmbalagem.sSigla = GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_Embalagem_Col)
            lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 95088 Then gError 95462
            
            'Se nao achou => ERRO
            If lErro = 95088 Then gError 95463
                                                                        
            'Preencher o obj
            objProdutoEmb.iEmbalagem = objEmbalagem.iCodigo
            objProdutoEmb.dCapacidade = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_Capacidade_Col))
            objProdutoEmb.dPesoBruto = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_PesoBruto_Col))
            objProdutoEmb.dPesoLiqTotal = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_PesoLiqTotal_Col))
            objProdutoEmb.sUMPeso = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_UMPeso_Col))
            objProdutoEmb.sNomeProdEmb = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_NomeProdEmb_Col))
            objProdutoEmb.sUMEmbalagem = Trim(GridEmbalagens.TextMatrix(iIndice, iGridEmbalagens_UMEmbalagem_Col))
            objProdutoEmb.iSeqGrid = iIndice
            
            objProduto.colProdutoEmbalagem.Add objProdutoEmb

        End If

    Next
        
    Move_GridEmbalagens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridEmbalagens_Memoria:

    Move_GridEmbalagens_Memoria = gErr

    Select Case gErr
        
        Case 95462
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case 95643
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM1", objEmbalagem.sSigla)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Embalagens
                Call Chama_Tela("Embalagem", objEmbalagem)
            Else
                Embalagem.SetFocus
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165619)
        
    End Select
    
    Exit Function

End Function

Function Carrega_UMPeso() As Long

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As New ClassUnidadeDeMedida

On Error GoTo Erro_Carrega_UMPeso

    'Carrega o código da classe de UM
    objClasseUM.iClasse = UM_PESO_CLASSE
    
    'Le as unidades de medida
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 27499 Then gError 95044
    
    'Se nao encontrou => erro!
    If lErro = 27499 Then gError 95045
    
    'Limpa o Combo de UM
    UMPeso.Clear
    
    'Adiciona a unidade de medida no Combo
    For Each objUM In colSiglas
        UMPeso.AddItem objUM.sSigla
    Next
    
    UMPeso.AddItem " "
    
    Carrega_UMPeso = SUCESSO
    
    Exit Function
    
Erro_Carrega_UMPeso:

    Carrega_UMPeso = gErr
    
    Select Case gErr
    
        Case 95044
        
        Case 95045
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSE_UMPESO_NAO_ENCONTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165620)
        
    End Select
    
    Exit Function

End Function

Function Carrega_UMEmbalagem(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim objUM As ClassUnidadeDeMedida
Dim colUM As New Collection

On Error GoTo Erro_Carrega_UMEmbalagem

    'Carrega o código da classe de UM
    objClasseUM.iClasse = objProduto.iClasseUM
        
    'Le as unidades de medida
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colUM)
    If lErro <> SUCESSO And lErro <> 27499 Then gError 95046
    
    'Se nao encontrou => erro!
    If lErro = 27499 Then gError 95047
    
    'Limpa o Combo de UM
    UMEmbalagem.Clear
    
    'Adiciona a unidade de medida no Combo
    For Each objUM In colUM
        UMEmbalagem.AddItem objUM.sSigla
    Next
    
    UMEmbalagem.AddItem " "
    
    Carrega_UMEmbalagem = SUCESSO
    
    Exit Function
    
Erro_Carrega_UMEmbalagem:

    Carrega_UMEmbalagem = gErr
    
    Select Case gErr
    
        Case 95046
        
        Case 95047
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSE_UMPESO_NAO_ENCONTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165621)
        
    End Select
    
    Exit Function
    
End Function

Function Preenche_GridEmbalagens(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim iLinha As Integer
Dim objEmbalagem As New ClassEmbalagem
Dim sPeso As String
Dim dFator As Double

On Error GoTo Erro_Preenche_GridEmbalagens
   
    For Each objProdutoEmbalagem In objProduto.colProdutoEmbalagem
        
        'Incrementa o contador
        iLinha = iLinha + 1
        
        'Se for a embalagem padrao => marca como padrao
        If objProdutoEmbalagem.iPadrao = PRODUTOEMBALAGEM_PADRAO Then GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_EmbalagemPadrao_Col) = vbChecked
                
        'Guarda em objEmbalagem o código da embalagem que está em objProdutoEmbalagem
        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
        
        'Le do BD a descricao da embalagem
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 95048
        
        'Se nao encontrou => erro
        If lErro = 82763 Then gError 95049
        
        'Exibe a embalagem
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col) = objEmbalagem.sSigla
        
        'Exibir o restante
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Capacidade_Col) = Formata_Estoque(objProdutoEmbalagem.dCapacidade)
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoLiqTotal_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMPeso_Col) = objProdutoEmbalagem.sUMPeso
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_NomeProdEmb_Col) = objProdutoEmbalagem.sNomeProdEmb
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMEmbalagem_Col) = objProdutoEmbalagem.sUMEmbalagem
            
        '??? converter peso embalagem para umpeso
        lErro = CF("UM_Conversao", UM_PESO_CLASSE, "Kg", objProdutoEmbalagem.sUMPeso, dFator)
        If lErro <> SUCESSO Then gError 95143
        
        '???Multiplica o peso embalagem pelo fator de conversao
        sPeso = CStr(dFator * objEmbalagem.dPeso)
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoEmbalagem_Col) = Formata_Estoque(CDbl(sPeso))
                
    
    Next
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    objGrid.iLinhasExistentes = iLinha

    Preenche_GridEmbalagens = SUCESSO

    Exit Function

Erro_Preenche_GridEmbalagens:

    Preenche_GridEmbalagens = gErr
    
    Select Case gErr
        
        Case 95048
        
        Case 95049
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_ENCONTRADA", gErr, objEmbalagem.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165622)
    
    End Select

    Exit Function
    
End Function

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    'Testa se houve alguma alteracao
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 95058
    
    'Limpa a tela
    Call Limpa_Tela_ProdutoEmbalagem
    
    FixarGrid.Value = vbUnchecked
        
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 95058
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165623)
    
    End Select
    
    Exit Sub
    
End Sub

Function Traz_ProdutoEmbalagem_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoEnxuto As String

On Error GoTo Erro_Traz_ProdutoEmbalagem_Tela
            
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 95107

    'Coloca o Codigo na tela
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Critica os dados
    Call Produto_Validate(bSGECancelDummy)
    
    Traz_ProdutoEmbalagem_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ProdutoEmbalagem_Tela:

    Traz_ProdutoEmbalagem_Tela = gErr
    
    Select Case gErr
    
        Case 95107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165624)
    
    End Select
    
    Exit Function
    
End Function

Function Limpa_Tela_ProdutoEmbalagem() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ProdutoEmbalagem

    'Limpa as TextBox
    Call Limpa_Tela(Me)
    
    'Limpa os campos
    LabelUMEstoque.Caption = ""
    Descricao.Caption = ""
        
    'Inicializa o NomeReduzido do Produto com vazio
    gsNomeRedProd = ""
    
    'Limpa o grid
    lErro = Grid_Limpa(objGrid)
    If lErro <> SUCESSO Then gError 95065
    
    iProdutoAlterado = 0
    iAlterado = 0
    FixarGrid.Value = vbUnchecked
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 95066

    Limpa_Tela_ProdutoEmbalagem = SUCESSO

    Exit Function
    
Erro_Limpa_Tela_ProdutoEmbalagem:

    Limpa_Tela_ProdutoEmbalagem = gErr
    
    Select Case gErr
    
        Case 95065, 95066
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165625)
    
    End Select
    
    Exit Function

End Function

Sub GridEmbalagens_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If
    
End Sub

Sub GridEmbalagens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Sub GridEmbalagens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridEmbalagens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridEmbalagens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridEmbalagens_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    Exit Sub

Erro_GridEmbalagens_KeyDown:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165626)

    End Select

    Exit Sub

End Sub

Sub GridEmbalagens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridEmbalagens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Sub GridEmbalagens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Sub GridEmbalagens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridEmbalagens.Col

            Case iGridEmbalagens_EmbalagemPadrao_Col

                lErro = Saida_Celula_EmbalagemPadrao(objGridInt)
                If lErro <> SUCESSO Then gError 95068

            Case iGridEmbalagens_Capacidade_Col
            
                lErro = Saida_Celula_Capacidade(objGridInt)
                If lErro <> SUCESSO Then gError 95112
            
            Case iGridEmbalagens_Embalagem_Col

                lErro = Saida_Celula_Embalagem(objGridInt)
                If lErro <> SUCESSO Then gError 95069

            Case iGridEmbalagens_NomeProdEmb_Col

                lErro = Saida_Celula_NomeProdEmb(objGridInt)
                If lErro <> SUCESSO Then gError 95070

            Case iGridEmbalagens_UMEmbalagem_Col

                lErro = Saida_Celula_UMEmbalagem(objGridInt)
                If lErro <> SUCESSO Then gError 95071

            Case iGridEmbalagens_UMPeso_Col

                lErro = Saida_Celula_UMPeso(objGridInt)
                If lErro <> SUCESSO Then gError 95072

            Case iGridEmbalagens_PesoLiqTotal_Col

                lErro = Saida_Celula_PesoLiqTotal(objGridInt)
                If lErro <> SUCESSO Then gError 95073

             Case iGridEmbalagens_PesoBruto_Col

                lErro = Saida_Celula_PesoBruto(objGridInt)
                If lErro <> SUCESSO Then gError 95075
        
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 95076
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function


Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case 95068 To 95073, 95075, 95112
        
        Case 95076
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165627)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_EmbalagemPadrao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_EmbalagemPadrao

    Set objGridInt.objControle = EmbalagemPadrao
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95074

    Saida_Celula_EmbalagemPadrao = SUCESSO
    
    Exit Function

Erro_Saida_Celula_EmbalagemPadrao:

    Saida_Celula_EmbalagemPadrao = gErr
    
    Select Case gErr
    
        Case 95074
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165628)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_Embalagem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem
Dim iLinha As Integer
Dim sSiglaAux As Integer
Dim objEmbalagemTextBox As Object
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Embalagem
    
    Set objGridInt.objControle = Embalagem

    'Verifica se a embalagem foi preenchida
    If Len(Trim(Embalagem.Text)) > 0 Then
        
        'Joga o controle Embalagem dentro de um obj
        Set objEmbalagemTextBox = Embalagem
        
        'Le os dados da embalagem
        lErro = CF("TP_Embalagem_Le_Grid", objEmbalagemTextBox, objEmbalagem)
        If lErro <> SUCESSO Then gError 95077
        
        'Verifica se há alguma embalagem repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridEmbalagens.Row Then
                                                    
                If GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col) = Embalagem.Text Then
                
                    Embalagem.Text = ""
                    gError 95081
                    
                End If
                    
            End If
                           
        Next
        
        'Preenche o grid
        Embalagem.Text = objEmbalagem.sSigla
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_NomeProdEmb_Col) = Mid(gsNomeRedProd & " " & objEmbalagem.sSigla, 1, 20)
         
        For iIndice = 0 To UMPeso.ListCount - 1
        
            If UCase(UMPeso.List(iIndice)) = "KG" Then Exit For
        
        Next
        
        'UM Default da embalagem
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMPeso_Col) = UMPeso.List(iIndice)
        
        GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoEmbalagem_Col) = Formata_Estoque(CDbl(objEmbalagem.dPeso))
        
        'Se necessário cria uma nova linha no Grid
        If GridEmbalagens.Row - GridEmbalagens.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
    
    End If
             
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Embalagem = SUCESSO

    Exit Function

Erro_Saida_Celula_Embalagem:

    Saida_Celula_Embalagem = gErr

    Select Case gErr

        Case 95465
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case 95466
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM1", objEmbalagem.sSigla)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Embalagens
                Call Chama_Tela("Embalagem", objEmbalagem)
            Else
                Embalagem.SetFocus
            End If
        
        Case 95081
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_REPETIDA", gErr, objEmbalagem.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95077, 95082
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165629)

    End Select

    Exit Function

End Function

Function Saida_Celula_NomeProdEmb(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NomeProdEmb

    Set objGridInt.objControle = NomeProdEmb
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95078

    Saida_Celula_NomeProdEmb = SUCESSO
    
    Exit Function

Erro_Saida_Celula_NomeProdEmb:

    Saida_Celula_NomeProdEmb = gErr
    
    Select Case gErr
    
        Case 95078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165630)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_UMEmbalagem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UMEmbalagem

    Set objGridInt.objControle = UMEmbalagem
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95079

    Saida_Celula_UMEmbalagem = SUCESSO
    
    Exit Function

Erro_Saida_Celula_UMEmbalagem:

    Saida_Celula_UMEmbalagem = gErr
    
    Select Case gErr
    
        Case 95079
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165631)
    
    End Select
    
    Exit Function

End Function


Function Saida_Celula_UMPeso(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem
Dim dFator As Double
Dim sPeso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_UMPeso

    Set objGridInt.objControle = UMPeso
    
    'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
    faz uma nova leitura em busca do codigo
    objEmbalagem.sSigla = GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_Embalagem_Col)
    lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 95088 Then gError 95470
    
    'Se nao achou => ERRO
    If lErro = 95088 Then gError 95471
'''
'''    '???Le a embalagem a partir do codigo
'''    lErro = CF("Embalagem_Le", objEmbalagem)
'''    If lErro <> SUCESSO Then gError 95141
'''
'''    '???Se nao encontrou => erro
'''    If lErro = 82763 Then gError 95142
        
    If Len(Trim(UMPeso.Text)) > 0 Then
        '??? converter peso embalagem para umpeso
        lErro = CF("UM_Conversao", UM_PESO_CLASSE, "Kg", UMPeso.Text, dFator)
        If lErro <> SUCESSO Then gError 95143
      
        '???Multiplica o peso embalagem pelo fator de conversao
        sPeso = CStr(dFator * objEmbalagem.dPeso)
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col) = Formata_Estoque(CDbl(sPeso))
    Else
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col) = Formata_Estoque(0)
    End If
      
    If Len(Trim(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoLiqTotal_Col))) > 0 Then
        
        'Calcula o peso bruto
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoBruto_Col) = Formata_Estoque((StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoLiqTotal_Col)) + StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col))))
    
    End If
        
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95080

    Saida_Celula_UMPeso = SUCESSO
    
    Exit Function

Erro_Saida_Celula_UMPeso:

    Saida_Celula_UMPeso = gErr
    
    Select Case gErr
    
        Case 95080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 95470
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case 95471
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM1", objEmbalagem.sSigla)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Embalagens
                Call Chama_Tela("Embalagem", objEmbalagem)
            Else
                Embalagem.SetFocus
            End If
        
        Case 95143
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165632)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_Capacidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Capacidade

    Set objGridInt.objControle = Capacidade
           
    'Se estiver preenchida
    If Len(Trim(Capacidade.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(Capacidade.Text)
        If lErro <> SUCESSO Then gError 95115

        'Coloca o valor Formatado na tela
        Capacidade.Text = Formata_Estoque(CDbl(Capacidade.Text))

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95111
        
    Saida_Celula_Capacidade = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Capacidade:

    Saida_Celula_Capacidade = gErr
    
    Select Case gErr
    
        Case 95111, 95115
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165633)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_PesoLiqTotal(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PesoLiqTotal

    Set objGridInt.objControle = PesoLiqTotal
    
    'Se estiver preenchida
    If Len(Trim(PesoLiqTotal.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(PesoLiqTotal.Text)
        If lErro <> SUCESSO Then gError 95116

        'Coloca o valor Formatado na tela
        PesoLiqTotal.Text = Formata_Estoque(CDbl(PesoLiqTotal.Text))

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95110
    
    If Len(Trim(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoLiqTotal_Col))) > 0 Then
        
        'Calcula o peso bruto
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoBruto_Col) = Formata_Estoque((StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoLiqTotal_Col)) + StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col))))
    
    End If
    
    Saida_Celula_PesoLiqTotal = SUCESSO
    
    Exit Function

Erro_Saida_Celula_PesoLiqTotal:

    Saida_Celula_PesoLiqTotal = gErr
    
    Select Case gErr
    
        Case 95110, 95116
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165634)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_PesoBruto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PesoBruto

    Set objGridInt.objControle = PesoBruto
    
    'Se estiver preenchida
    If Len(Trim(PesoBruto.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(PesoBruto.Text)
        If lErro <> SUCESSO Then gError 95117

        'Critica se PesoEmbalagem é maior do que PesoBruto
        If (StrParaDbl(PesoEmbalagem.Text) > StrParaDbl(PesoBruto.Text)) Then gError 95140
    
        'Coloca o valor Formatado na tela
        PesoBruto.Text = Formata_Estoque(CDbl(PesoBruto.Text))

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95118
    
    If Len(Trim(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoBruto_Col))) > 0 Then
        
        'Calcula o peso liquido total
        GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoLiqTotal_Col) = Formata_Estoque((StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoBruto_Col)) - StrParaDbl(GridEmbalagens.TextMatrix(GridEmbalagens.Row, iGridEmbalagens_PesoEmbalagem_Col))))
        
    End If
    
    Saida_Celula_PesoBruto = SUCESSO
    
    Exit Function

Erro_Saida_Celula_PesoBruto:

    Saida_Celula_PesoBruto = gErr
    
    Select Case gErr
    
        Case 95118, 95117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95140
            Call Rotina_Erro(vbOKOnly, "ERRO_PESOBRUTO_MENOR_PESOEMBALAGEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165635)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGrid = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEmbalagens = Nothing
     
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Function ProdutoEmbalagem_Critica() As Long

Dim iLinha As Integer
Dim iPadraoMarcado As Integer
Dim dPeso As Double
Dim dPesoBruto As Double

On Error GoTo Erro_ProdutoEmbalagem_Critica

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95090

    'Verifica se o grid está vazio
    If objGrid.iLinhasExistentes = 0 Then gError 95091

    'Para cada linha do grid
    For iLinha = 1 To objGrid.iLinhasExistentes
            
        'Verifica se todas as colunas da linha iLinha do grid estao preenchidas
        If StrParaInt(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_EmbalagemPadrao_Col)) = vbChecked Then iPadraoMarcado = PRODUTOEMBALAGEM_PADRAO
                
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col))) = 0 Then gError 95093
                
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_NomeProdEmb_Col))) = 0 Then gError 95094
        
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMEmbalagem_Col))) = 0 Then gError 95098
            
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Capacidade_Col))) = 0 Then gError 95092
        
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_UMPeso_Col))) = 0 Then gError 95099
                
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoLiqTotal_Col))) = 0 Then gError 95097
                
        If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoBruto_Col))) = 0 Then gError 95095
                       
        'Verifica consistencia do peso total = PesoLiqTotal + PesoEmbalagem - PesoBruto
        dPeso = StrParaDbl(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoLiqTotal_Col)) + StrParaDbl(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoEmbalagem_Col))
        dPesoBruto = StrParaDbl(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_PesoBruto_Col))
        If Abs(dPeso - dPesoBruto) > QTDE_ESTOQUE_DELTA Then gError 95101
                
    Next
    
    'Se nao tiver um padrao => erro
    If iPadraoMarcado <> PRODUTOEMBALAGEM_PADRAO Then gError 95100
    
    ProdutoEmbalagem_Critica = SUCESSO
    
    Exit Function
    
Erro_ProdutoEmbalagem_Critica:

    ProdutoEmbalagem_Critica = gErr
    
    Select Case gErr
    
        Case 95090
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 95091
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
            
        Case 95092
            Call Rotina_Erro(vbOKOnly, "ERRO_CAPACIDADE_NAO_PREENCHIDA1", gErr, iLinha)
            
        Case 95093
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 95094
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEPRODEMB_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 95095
            Call Rotina_Erro(vbOKOnly, "ERRO_PESOBRUTO_NAO_PREENCHIDO", gErr, iLinha)
            
        Case 95097
            Call Rotina_Erro(vbOKOnly, "ERRO_PESOLIQTOTAL_NAO_PREENCHIDO", gErr, iLinha)
            
        Case 95098
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEMBALAGEM_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 95099
            Call Rotina_Erro(vbOKOnly, "ERRO_UMPESO_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 95100
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEMPADRAO_OBRIGATORIA", gErr)
            
        Case 95101
            Call Rotina_Erro(vbOKOnly, "ERRO_PESOTOTAL_INCONSISTENTE", gErr, iLinha, dPesoBruto, dPeso)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165636)
            
    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProdEmb"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 95102

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Produto", objProduto.sCodigo, STRING_PRODUTO, "Produto"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 95102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165637)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objProduto.sCodigo = colCampoValor.Item("Produto").vValor

    lErro = Traz_ProdutoEmbalagem_Tela(objProduto)
    If lErro <> SUCESSO Then gError 95103

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165638)

    End Select

    Exit Sub

End Sub

Private Function GridEmbalagens_Inicializa(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Seq.")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Embalagem")
    objGridInt.colColuna.Add ("Desc. Produto X Embalagem")
    objGridInt.colColuna.Add ("U.M. Emb.")
    objGridInt.colColuna.Add ("Capacidade")
    objGridInt.colColuna.Add ("U.M. Peso")
    objGridInt.colColuna.Add ("Peso Liq.Total")
    objGridInt.colColuna.Add ("Peso Emb.")
    objGridInt.colColuna.Add ("Peso Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (EmbalagemPadrao.Name)
    objGridInt.colCampo.Add (Embalagem.Name)
    objGridInt.colCampo.Add (NomeProdEmb.Name)
    objGridInt.colCampo.Add (UMEmbalagem.Name)
    objGridInt.colCampo.Add (Capacidade.Name)
    objGridInt.colCampo.Add (UMPeso.Name)
    objGridInt.colCampo.Add (PesoLiqTotal.Name)
    objGridInt.colCampo.Add (PesoEmbalagem.Name)
    objGridInt.colCampo.Add (PesoBruto.Name)

    iGridEmbalagens_EmbalagemPadrao_Col = 1
    iGridEmbalagens_Embalagem_Col = 2
    iGridEmbalagens_NomeProdEmb_Col = 3
    iGridEmbalagens_UMEmbalagem_Col = 4
    iGridEmbalagens_Capacidade_Col = 5
    iGridEmbalagens_UMPeso_Col = 6
    iGridEmbalagens_PesoLiqTotal_Col = 7
    iGridEmbalagens_PesoEmbalagem_Col = 8
    iGridEmbalagens_PesoBruto_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridEmbalagens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 50

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridEmbalagens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    'Habilita a embalagem
    Embalagem.Enabled = True
    
    GridEmbalagens_Inicializa = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
        
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        
        'Embalagem
        Case Embalagem.Name
            If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col))) > 0 Then
                Embalagem.Enabled = False
            Else
                Embalagem.Enabled = True
            End If
        
        'Restante das colunas
        Case NomeProdEmb.Name, UMEmbalagem.Name, Capacidade.Name, UMPeso.Name, PesoLiqTotal.Name, PesoBruto.Name
        
            If Len(Trim(GridEmbalagens.TextMatrix(iLinha, iGridEmbalagens_Embalagem_Col))) = 0 Then
                NomeProdEmb.Enabled = False
                UMEmbalagem.Enabled = False
                Capacidade.Enabled = False
                UMPeso.Enabled = False
                PesoLiqTotal.Enabled = False
                PesoBruto.Enabled = False
    
            Else
                NomeProdEmb.Enabled = True
                UMEmbalagem.Enabled = True
                Capacidade.Enabled = True
                UMPeso.Enabled = True
                PesoLiqTotal.Enabled = True
                PesoBruto.Enabled = True
                
            End If
    
    End Select
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165639)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Embalagem Then
            Call BotaoEmbalagens_Click
        End If
        
    End If

End Sub
