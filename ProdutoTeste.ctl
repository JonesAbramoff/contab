VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ProdutoTeste 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ScaleHeight     =   6405
   ScaleWidth      =   8850
   Begin VB.TextBox LabelObservacao 
      BackColor       =   &H8000000F&
      Height          =   795
      Left            =   4575
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   4950
      Width           =   4050
   End
   Begin VB.TextBox LabelEspecificacao 
      BackColor       =   &H8000000F&
      Height          =   795
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   4950
      Width           =   4050
   End
   Begin VB.CommandButton BotaoImprimirFicha 
      Caption         =   "Imprimir Ficha de Controle de Qualidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   180
      TabIndex        =   24
      Top             =   5865
      Width           =   4185
   End
   Begin VB.CommandButton BotaoTestes 
      Caption         =   "Testes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7215
      TabIndex        =   16
      Top             =   4365
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Testes associados ao Produto"
      Height          =   2700
      Left            =   165
      TabIndex        =   12
      Top             =   1545
      Width           =   8490
      Begin VB.TextBox Metodo 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1695
         Width           =   2475
      End
      Begin VB.CheckBox NoCertificado 
         Caption         =   "Check1"
         Height          =   240
         Left            =   3945
         TabIndex        =   20
         Top             =   870
         Width           =   975
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5040
         MaxLength       =   250
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1230
         Width           =   2475
      End
      Begin VB.TextBox Especificacao 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5025
         MaxLength       =   250
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   840
         Width           =   2475
      End
      Begin VB.TextBox Teste 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   375
         MaxLength       =   100
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   840
         Width           =   2500
      End
      Begin MSMask.MaskEdBox LimiteDe 
         Height          =   315
         Left            =   2115
         TabIndex        =   13
         Top             =   870
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LimiteAte 
         Height          =   315
         Left            =   3015
         TabIndex        =   14
         Top             =   855
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTestes 
         Height          =   2190
         Left            =   210
         TabIndex        =   15
         Top             =   300
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   3863
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ProdutoTeste.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ProdutoTeste.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ProdutoTeste.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ProdutoTeste.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
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
         TabIndex        =   6
         Top             =   885
         Width           =   930
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
         TabIndex        =   5
         Top             =   390
         Width           =   735
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
         Left            =   3435
         TabIndex        =   4
         Top             =   375
         Width           =   1230
      End
      Begin VB.Label LabelUMEstoque 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4755
         TabIndex        =   3
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   4410
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Especificação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   195
      TabIndex        =   22
      Top             =   4710
      Width           =   1305
   End
   Begin VB.Label Label10 
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
      Height          =   180
      Left            =   4575
      TabIndex        =   21
      Top             =   4710
      Width           =   1305
   End
End
Attribute VB_Name = "ProdutoTeste"
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

'Declaração utilizada para evento LabelProduto_Click
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Declaração utilizada para evento BotaoTestes_Click
Private WithEvents objEventoTestes As AdmEvento
Attribute objEventoTestes.VB_VarHelpID = -1

'Variáveis que indicam os índices de cada coluna do grid
Dim iGridTestes_Teste_Col As Integer
Dim iGridTestes_LimiteDe_Col As Integer
Dim iGridTestes_LimiteAte_Col As Integer
Dim iGridTestes_NoCertificado_Col As Integer
Dim iGridTestes_Especificacao_Col As Integer
Dim iGridTestes_Observacao_Col As Integer
Dim iGridTestes_Metodo_Col As Integer

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    '*** Fernando, Criar IDH
    'Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Produto x Testes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ProdutoTeste"

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

Private Sub BotaoImprimirFicha_Click()

Dim lErro As Long, sProduto As String, objRelatorio As New AdmRelatorio
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoImprimirFicha_Click

    'se estiver preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Formata o código do produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 95040
        
        'Move para o objProduto o código do produto no formato do BD
        sProduto = sProdutoFormatado
        
    End If
    
    Call objRelatorio.Rel_Menu_Executar("Ficha de Controle de Qualidade", sProduto)
        
    Exit Sub
     
Erro_BotaoImprimirFicha_Click:

    Select Case gErr
          
        Case 95040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165745)
     
    End Select
     
    Exit Sub

End Sub

Private Sub LimiteDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LimiteDe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub LimiteDe_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub LimiteDe_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub LimiteDe_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = LimiteDe
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub LimiteAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LimiteAte_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub LimiteAte_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub LimiteAte_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub LimiteAte_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = LimiteAte
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NoCertificado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NoCertificado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub NoCertificado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub NoCertificado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub NoCertificado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = NoCertificado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Especificacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Especificacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Especificacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Especificacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Especificacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Especificacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Observacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Observacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Metodo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Metodo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Metodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Metodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Metodo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Metodo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

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

Private Sub BotaoTestes_Click()

Dim objTeste As New ClassTestesQualidade
Dim colSelecao As New Collection
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoTestes_Click

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95121

    'Verifica se há uma linha selecionada
    If GridTestes.Row = 0 Then gError 95108
        
    'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
    faz uma nova leitura em busca do codigo
    objTeste.sNomeReduzido = GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col)
    lErro = CF("TesteQualidade_Le_NomeReduzido", objTeste)
    If lErro <> SUCESSO And lErro <> 130109 Then gError 95468
    
    'chama a tela de browser
    Call Chama_Tela("TestesQualidadeLista", colSelecao, objTeste, objEventoTestes)
    
    iAlterado = 1
    
    Exit Sub
    
Erro_BotaoTestes_Click:
    
    Select Case gErr
    
        Case 95468
        
        Case 95469
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTeste.sNomeReduzido)
            
            If vbMsgRes = vbYes Then
                'Chama a tela de Testes
                Call Chama_Tela("TestesQualidade", objTeste)
            End If
        
        Case 95108
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 95121
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165746)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoTestes_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTeste As ClassTestesQualidade
Dim iLinha As Integer
Dim sSiglaAux As String
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_objEventoTestes_evSelecao
           
    'Verifica se o teste da linha atual está preenchido
    If Len(Trim(GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col))) = 0 Then
            
        'Define o tipo de obj recebido (Tipo Teste)
        Set objTeste = obj1
        
        'Verifica se há algumo teste repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridTestes.Row Then
                
                If GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido Then
                    Teste.Text = ""
                    gError 95120
                End If
                    
            End If
                           
        Next
        
        If ActiveControl Is Teste Then
            Teste.Text = objTeste.sNomeReduzido
        Else
        
            'Preenche o grid
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Teste_Col) = objTeste.sNomeReduzido
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_LimiteDe_Col) = Format(objTeste.dLimiteDe, FORMATO_LIMITE_TESTE)
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_LimiteAte_Col) = Format(objTeste.dLimiteAte, FORMATO_LIMITE_TESTE)
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_NoCertificado_Col) = IIf(objTeste.iNoCertificado, MARCADO, DESMARCADO)
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Especificacao_Col) = objTeste.sEspecificacao
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Observacao_Col) = objTeste.sObservacao
            GridTestes.TextMatrix(GridTestes.Row, iGridTestes_Metodo_Col) = objTeste.sMetodoUsado
            
            'Cria mais uma linha no grid
            If GridTestes.Row - GridTestes.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
            
            Call Grid_Refresh_Checkbox(objGrid)
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTestes_evSelecao:

    Select Case gErr

        Case 95464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objTeste.iCodigo)
        
        Case 95120
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_REPETIDA", gErr, objTeste.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGrid)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165747)
              
    End Select
    
    Exit Sub
        
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    'Traz para a tela
    lErro = Traz_ProdutoTeste_Tela(objProduto)
    If lErro <> SUCESSO Then gError 95106
    
    Me.Show
    
    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case 95106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165748)
    
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
    
    '???ProdutosSubstLista
    'chama a tela de browser
    Call Chama_Tela("ProdutosSubstLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub
    
Erro_LabelProduto_Click:

    Select Case gErr
    
        Case 95012
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165749)
            
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
    
    'Se existir linha válida ...
    If objGrid.iLinhasExistentes > 0 Then
        '... Limpa o grid
        lErro = Grid_Limpa(objGrid)
        If lErro <> SUCESSO Then gError 95065
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
        
        'Exibe os dados na tela
        Produto.PromptInclude = False
        Produto.Text = objProduto.sCodigo
        Produto.PromptInclude = True
        Descricao.Caption = objProduto.sDescricao
        LabelUMEstoque.Caption = objProduto.sSiglaUMEstoque
        gsNomeRedProd = objProduto.sNomeReduzido
               
        'Le os testes já relacionadas com o produto
        lErro = CF("ProdutoTeste_Le_Produto", objProduto)
        If lErro <> SUCESSO And lErro <> 130296 Then gError 95008
        
        If lErro = SUCESSO Then
            
            'Carrega o grid com os dados do Obj
            lErro = Preenche_GridTestes(objProduto)
            If lErro <> SUCESSO Then gError 95009
            
        End If
        
        iAlterado = 0
        iProdutoAlterado = 0
        
    End If
        
    Exit Sub
        
Erro_Produto_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 95005, 95006, 95009, 95008, 95050
        
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
                Call Limpa_Tela_ProdutoTeste
                Produto.SetFocus
            End If
        
        Case 95007
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, objProduto.sCodigo, giFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165750)
            
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
    Call GridTestes_Inicializa(objGrid)
    
    'Inicializa o objEventoProduto
    Set objEventoProduto = New AdmEvento
    
    'Inicializa o objEventoTestes
    Set objEventoTestes = New AdmEvento

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
    
        Case 95038
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165751)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objProduto Is Nothing Then
        lErro = Traz_ProdutoTeste_Tela(objProduto)
        If lErro <> SUCESSO Then gError 95004
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 95004
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165752)
    End Select
    
    Exit Function

End Function

Private Sub Teste_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Teste_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Teste_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub
Private Sub Teste_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub
Private Sub Teste_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Teste
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
    Call Limpa_Tela_ProdutoTeste
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
        
        Case 95018
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165753)
            
    End Select
    
    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProdutoTeste As ClassProdutoTeste

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
       
    'Critica os dados da tela
    lErro = ProdutoTeste_Critica()
    If lErro <> SUCESSO Then gError 95089
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then gError 95021
    
    'Verifica se o registro já existe e confirma se deseja alterar
    lErro = Trata_Alteracao(objProdutoTeste, objProduto.sCodigo)
    If lErro <> SUCESSO Then gError 95023
    
    'Grava
    lErro = CF("ProdutoTeste_Grava", objProduto)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165754)
            
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
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PRODUTOTESTE", objProduto.sCodigo)

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
    lErro = CF("ProdutoTeste_Exclui", objProduto)
    If lErro <> SUCESSO Then gError 95036
    
    'limpa a tela
    lErro = Limpa_Tela_ProdutoTeste
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165755)
            
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
        lErro = Move_GridTestes_Memoria(objProduto)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165756)
            
    End Select

    Exit Function

End Function

Function Move_GridTestes_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutoTeste As ClassProdutoTeste
Dim objTeste As New ClassTestesQualidade
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Move_GridTestes_Memoria

    'Para cada linha do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        'Verifica se a Teste foi preenchida
        If Len(Trim(GridTestes.TextMatrix(iIndice, iGridTestes_Teste_Col))) <> 0 Then
                        
            'Inicializa o Obj
            Set objProdutoTeste = New ClassProdutoTeste
            
            'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
            faz uma nova leitura em busca do codigo
            Set objTeste = New ClassTestesQualidade
            objTeste.sNomeReduzido = GridTestes.TextMatrix(iIndice, iGridTestes_Teste_Col)
            lErro = CF("TesteQualidade_Le_NomeReduzido", objTeste)
            If lErro <> SUCESSO And lErro <> 130109 Then gError 95462
            
            'Se nao achou => ERRO
            If lErro = 95088 Then gError 95463
                                                                        
            'Preencher o obj
            With objProdutoTeste
                .sProduto = objProduto.sCodigo
                .iTesteCodigo = objTeste.iCodigo
                .iSeqGrid = iIndice
                .sTesteEspecificacao = GridTestes.TextMatrix(iIndice, iGridTestes_Especificacao_Col)
                .iTesteTipoResultado = objTeste.iTipoResultado '??? nao aparece na tela
                .dTesteLimiteDe = StrParaDbl(GridTestes.TextMatrix(iIndice, iGridTestes_LimiteDe_Col))
                .dTesteLimiteAte = StrParaDbl(GridTestes.TextMatrix(iIndice, iGridTestes_LimiteAte_Col))
                .sTesteMetodoUsado = GridTestes.TextMatrix(iIndice, iGridTestes_Metodo_Col)
                .sTesteObservacao = GridTestes.TextMatrix(iIndice, iGridTestes_Observacao_Col)
                .iTesteNoCertificado = StrParaInt(GridTestes.TextMatrix(iIndice, iGridTestes_NoCertificado_Col))
            End With
            
            objProduto.colProdutoTeste.Add objProdutoTeste

        End If

    Next
        
    Move_GridTestes_Memoria = SUCESSO

    Exit Function

Erro_Move_GridTestes_Memoria:

    Move_GridTestes_Memoria = gErr

    Select Case gErr
        
        Case 95462
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTEQUALIDADE", gErr, objTeste.iCodigo)
        
        Case 95643
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTeste.sNomeReduzido)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Testes
                Call Chama_Tela("TestesQualidade", objTeste)
            Else
                Teste.SetFocus
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165757)
        
    End Select
    
    Exit Function

End Function

Function Preenche_GridTestes(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoTeste As New ClassProdutoTeste
Dim iLinha As Integer
Dim objTeste As New ClassTestesQualidade
Dim sPeso As String
Dim dFator As Double

On Error GoTo Erro_Preenche_GridTestes
   
    'Limpa o grid
    lErro = Grid_Limpa(objGrid)
    If lErro <> SUCESSO Then gError 95144
            
    For Each objProdutoTeste In objProduto.colProdutoTeste
        
        'Incrementa o contador
        iLinha = iLinha + 1
        
        'Guarda em objTeste o código do teste que está em objProdutoTeste
        objTeste.iCodigo = objProdutoTeste.iTesteCodigo
        
        'Le do BD a descricao do teste
        lErro = CF("TestesQualidade_Le", objTeste)
        If lErro <> SUCESSO And lErro <> 130109 Then gError 95048
        
        'Se nao encontrou => erro
        If lErro = 82763 Then gError 95049
        
        'Exibe o teste
        GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = objTeste.sNomeReduzido
        
        'Exibir o restante
        GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col) = objProdutoTeste.sTesteEspecificacao
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteDe_Col) = Format(objProdutoTeste.dTesteLimiteDe, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_LimiteAte_Col) = Format(objProdutoTeste.dTesteLimiteAte, FORMATO_LIMITE_TESTE)
        GridTestes.TextMatrix(iLinha, iGridTestes_Metodo_Col) = objProdutoTeste.sTesteMetodoUsado
        GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col) = objProdutoTeste.sTesteObservacao
        GridTestes.TextMatrix(iLinha, iGridTestes_NoCertificado_Col) = objProdutoTeste.iTesteNoCertificado
    
    Next
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    objGrid.iLinhasExistentes = iLinha

    Preenche_GridTestes = SUCESSO

    Exit Function

Erro_Preenche_GridTestes:

    Preenche_GridTestes = gErr
    
    Select Case gErr
        
        Case 95048, 95144
        
        Case 95049
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTEQUALIDADE_NAO_ENCONTRADO", gErr, objTeste.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165758)
    
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
    Call Limpa_Tela_ProdutoTeste
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 95058
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165759)
    
    End Select
    
    Exit Sub
    
End Sub

Function Traz_ProdutoTeste_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoEnxuto As String

On Error GoTo Erro_Traz_ProdutoTeste_Tela
            
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 95107

    'Coloca o Codigo na tela
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Critica os dados
    Call Produto_Validate(bSGECancelDummy)
    
    Traz_ProdutoTeste_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ProdutoTeste_Tela:

    Traz_ProdutoTeste_Tela = gErr
    
    Select Case gErr
    
        Case 95107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165760)
    
    End Select
    
    Exit Function
    
End Function

Function Limpa_Tela_ProdutoTeste() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ProdutoTeste

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
    
    Call Teste_LimpaInfo
    
    iProdutoAlterado = 0
    iAlterado = 0
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 95066

    Limpa_Tela_ProdutoTeste = SUCESSO

    Exit Function
    
Erro_Limpa_Tela_ProdutoTeste:

    Limpa_Tela_ProdutoTeste = gErr
    
    Select Case gErr
    
        Case 95065, 95066
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165761)
    
    End Select
    
    Exit Function

End Function

Sub GridTestes_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If
    
End Sub

Sub GridTestes_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Sub GridTestes_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridTestes_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridTestes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridTestes_KeyDown

    'Guardo o item atual e o número de linhas existente
    iLinhasExistentesAnterior = objGrid.iLinhasExistentes
    iLinhaAnterior = GridTestes.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGrid)
    
    'se alguma linha foi excluída
    If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
        If objGrid.iLinhasExistentes = 0 Then
            Call Teste_LimpaInfo
        Else
            Call Teste_ExibeInfo(GridTestes.Row)
        End If
    End If
    
    Exit Sub

Erro_GridTestes_KeyDown:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165762)

    End Select

    Exit Sub

End Sub

Sub GridTestes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridTestes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Sub GridTestes_RowColChange()

    Call Grid_RowColChange(objGrid)

    If GridTestes.Row <> 0 Then
        Call Teste_ExibeInfo(GridTestes.Row)
    End If
    
End Sub

Sub GridTestes_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridTestes.Col

            Case iGridTestes_Teste_Col

                lErro = Saida_Celula_Teste(objGridInt)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_LimiteDe_Col
                
                lErro = Saida_Celula_Limite(objGridInt, LimiteDe)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_LimiteAte_Col
            
                lErro = Saida_Celula_Limite(objGridInt, LimiteAte)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_NoCertificado_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, NoCertificado)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Especificacao_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, Especificacao)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Observacao_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, Observacao)
                If lErro <> SUCESSO Then gError 95069

            Case iGridTestes_Metodo_Col
            
                lErro = Saida_Celula_Padrao(objGridInt, Metodo)
                If lErro <> SUCESSO Then gError 95069

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165763)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_Teste(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTeste As New ClassTestesQualidade
Dim iLinha As Integer
Dim sSiglaAux As Integer
Dim objTesteTextBox As Object
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Teste
    
    Set objGridInt.objControle = Teste

    'Verifica se o teste foi preenchida
    If Len(Trim(Teste.Text)) > 0 Then
        
        'Joga o controle Teste dentro de um obj
        Set objTesteTextBox = Teste
        
        'Le os dados do teste
        lErro = CF("TP_TesteQualidade_Le_Grid", objTesteTextBox, objTeste)
        If lErro <> SUCESSO Then gError 95077
        
        'Verifica se há algumo teste repetida no grid
        For iLinha = 1 To objGrid.iLinhasExistentes
            
            If iLinha <> GridTestes.Row Then
                                                    
                If GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col) = Teste.Text Then
                
                    Teste.Text = ""
                    gError 95081
                    
                End If
                    
            End If
                           
        Next
        
        'Preenche o grid
        Teste.Text = objTeste.sNomeReduzido
        '??? completar
        
        'Se necessário cria uma nova linha no Grid
        If GridTestes.Row - GridTestes.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
    
    End If
             
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Teste = SUCESSO

    Exit Function

Erro_Saida_Celula_Teste:

    Saida_Celula_Teste = gErr

    Select Case gErr

        Case 95465
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTEQUALIDADE", gErr, objTeste.iCodigo)
        
        Case 95466
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TESTEQUALIDADE1", objTeste.sNomeReduzido)
        
            If vbMsgRes = vbYes Then
                'Chama a tela de Testes
                Call Chama_Tela("TestesQualidade", objTeste)
            Else
                Teste.SetFocus
            End If
        
        Case 95081
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTEQUALIDADE_REPETIDO", gErr, objTeste.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 95077, 95082
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165764)

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

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing
    Set objEventoTestes = Nothing
    
    Set objGrid = Nothing
     
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Function ProdutoTeste_Critica() As Long

Dim iLinha As Integer

On Error GoTo Erro_ProdutoTeste_Critica

    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 95090

    'Verifica se o grid está vazio
    If objGrid.iLinhasExistentes = 0 Then gError 95091

    'Para cada linha do grid
    For iLinha = 1 To objGrid.iLinhasExistentes
            
        'Verifica se todas as colunas da linha iLinha do grid estao preenchidas
        If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) = 0 Then gError 95093
                
        '??? completar
        
    Next
    
    ProdutoTeste_Critica = SUCESSO
    
    Exit Function
    
Erro_ProdutoTeste_Critica:

    ProdutoTeste_Critica = gErr
    
    Select Case gErr
    
        Case 95090
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 95091
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165765)
            
    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProdutoTeste"

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165766)

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

    lErro = Traz_ProdutoTeste_Tela(objProduto)
    If lErro <> SUCESSO Then gError 95103

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165767)

    End Select

    Exit Sub

End Sub

Private Function GridTestes_Inicializa(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Seq.")
    objGridInt.colColuna.Add ("Teste")
    objGridInt.colColuna.Add ("Limite De")
    objGridInt.colColuna.Add ("Limite Até")
    objGridInt.colColuna.Add ("Certificado")
    objGridInt.colColuna.Add ("Especificação")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Método")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Teste.Name)
    objGridInt.colCampo.Add (LimiteDe.Name)
    objGridInt.colCampo.Add (LimiteAte.Name)
    objGridInt.colCampo.Add (NoCertificado.Name)
    objGridInt.colCampo.Add (Especificacao.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    objGridInt.colCampo.Add (Metodo.Name)

    iGridTestes_Teste_Col = 1
    iGridTestes_LimiteDe_Col = 2
    iGridTestes_LimiteAte_Col = 3
    iGridTestes_NoCertificado_Col = 4
    iGridTestes_Especificacao_Col = 5
    iGridTestes_Observacao_Col = 6
    iGridTestes_Metodo_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridTestes

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_TESTES_PRODUTO + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridTestes.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    'Habilita o teste
    Teste.Enabled = True
    
    GridTestes_Inicializa = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
        
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        
        'Teste
        Case Teste.Name
            If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) > 0 Then
                Teste.Enabled = False
            Else
                Teste.Enabled = True
            End If
        
        'Restante das colunas
        Case LimiteDe.Name, LimiteAte.Name, NoCertificado.Name, Especificacao.Name, Observacao.Name, Metodo.Name

            If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) = 0 Then
                LimiteDe.Enabled = False
                LimiteAte.Enabled = False
                NoCertificado.Enabled = False
                Especificacao.Enabled = False
                Observacao.Enabled = False
                Metodo.Enabled = False

            Else
                LimiteDe.Enabled = True
                LimiteAte.Enabled = True
                NoCertificado.Enabled = True
                Especificacao.Enabled = True
                Observacao.Enabled = True
                Metodo.Enabled = True

            End If
    
    End Select
    
    Call Teste_ExibeInfo(iLinha)
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165768)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        ElseIf Me.ActiveControl Is Teste Then
            Call BotaoTestes_Click
        End If
        
    End If

End Sub

Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Padrao
    
    Set objGridInt.objControle = objControle
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 95082
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165769)

    End Select

    Exit Function

End Function

Function Saida_Celula_Limite(objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Limite
    
    Set objGridInt.objControle = objControle
                   
    If Len(Trim(objControle.Text)) > 0 Then 'Inserido por Wagner
                   
        lErro = Valor_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 130314
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95082
    
    Saida_Celula_Limite = SUCESSO

    Exit Function

Erro_Saida_Celula_Limite:

    Saida_Celula_Limite = gErr

    Select Case gErr

        Case 95082, 130314
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165770)

    End Select

    Exit Function

End Function

Private Sub Teste_ExibeInfo(ByVal iLinha As Integer)

Dim lErro As Long

On Error GoTo Erro_Teste_ExibeInfo

    If Len(Trim(GridTestes.TextMatrix(iLinha, iGridTestes_Teste_Col))) > 0 Then
    
        LabelEspecificacao = GridTestes.TextMatrix(iLinha, iGridTestes_Especificacao_Col)
        LabelObservacao = GridTestes.TextMatrix(iLinha, iGridTestes_Observacao_Col)
    
    Else
        
        Call Teste_LimpaInfo
    
    End If
    
    Exit Sub
     
Erro_Teste_ExibeInfo:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165771)
     
    End Select
     
    Exit Sub

End Sub

Sub Teste_LimpaInfo()

    LabelEspecificacao = ""
    LabelObservacao = ""
    
End Sub


