VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelOpBarCodeOcx 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   6330
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   900
      Width           =   2910
   End
   Begin VB.TextBox Descricao 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2055
      MaxLength       =   250
      TabIndex        =   10
      Top             =   1710
      Width           =   2730
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   4830
      TabIndex        =   11
      Top             =   1725
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   225
      Left            =   885
      TabIndex        =   12
      Top             =   1710
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.CommandButton BotaoProduto 
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
      Left            =   180
      TabIndex        =   9
      Top             =   3825
      Width           =   1605
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBarCodeOcx.ctx":0000
      Left            =   900
      List            =   "RelOpBarCodeOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   255
      Width           =   2910
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
      Left            =   4245
      Picture         =   "RelOpBarCodeOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   795
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4065
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpBarCodeOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpBarCodeOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpBarCodeOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBarCodeOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   2175
      Left            =   180
      TabIndex        =   8
      Top             =   1530
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   21
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   375
      TabIndex        =   14
      Top             =   945
      Width           =   450
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
      Left            =   195
      TabIndex        =   7
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpBarCodeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giTipoPadrao As Integer

Dim iAlterado As Integer
Dim iLinhaAntiga As Integer

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quantidade_Col As Integer

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoProduto = New AdmEvento
    Set objGridItens = New AdmGrid
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 182781
    
    lErro = Carrega_Tipo(Tipo)
    If lErro <> SUCESSO Then gError 182782

    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 182783

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 182781 To 182783
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182761)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim iNumLinhas As Integer
Dim sProduto As String
Dim objProduto As ClassProduto

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 182784
    
    Call Grid_Limpa(objGridItens)
   
    'pega o número de linhas
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError 182785
    
    Tipo.ListIndex = -1
    For iIndice = 0 To Tipo.ListCount - 1
        If Codigo_Extrai(Tipo.List(iIndice)) = StrParaInt(sParam) Then
            Tipo.ListIndex = iIndice
            Exit For
        End If
    Next
   
    'pega o número de linhas
    lErro = objRelOpcoes.ObterParametro("NNUMLIN", sParam)
    If lErro <> SUCESSO Then gError 182785
    
    iNumLinhas = StrParaInt(sParam)
    
    For iIndice = 1 To iNumLinhas
       
        'pega Produto
        lErro = objRelOpcoes.ObterParametro("TPROD" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError 182786
        
        sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Mascara_RetornaProdutoEnxuto(sParam, sProduto)
        If lErro <> SUCESSO Then gError 182787
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = sParam
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 182788
    
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objProduto.sDescricao
    
        'pega quantidade
        lErro = objRelOpcoes.ObterParametro("NQUANT" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError 182789
        
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = sParam
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iNumLinhas
        
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 182784 To 182789

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182762)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 182790
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 182791
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 182790
        
        Case 182791
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182763)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(ByVal colRelBarCode As Collection) As Long

Dim lErro As Long
Dim sProd As String
Dim iIndice As Integer
Dim iProdPreenchido As Integer
Dim objRelOpBarCode As ClassRelBarCode

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Tipo.ListIndex = -1 Then gError 182809

    For iIndice = 1 To objGridItens.iLinhasExistentes

        'formata o Produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProd, iProdPreenchido)
        If lErro <> SUCESSO Then gError 182792
    
        If iProdPreenchido <> PRODUTO_PREENCHIDO Then sProd = ""
        
        Set objRelOpBarCode = New ClassRelBarCode

        objRelOpBarCode.iQuantidade = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objRelOpBarCode.sProduto = sProd
        objRelOpBarCode.sDescricao = GridItens.TextMatrix(iIndice, iGrid_Descricao_Col)

        colRelBarCode.Add objRelOpBarCode

    Next
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 182792
        
        Case 182809
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", gErr)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182764)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 182793
    
    Call Grid_Limpa(objGridItens)
    
    ComboOpcoes.Text = ""
    
    Tipo.ListIndex = giTipoPadrao
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 182793
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182765)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoProduto = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objGridItens = Nothing
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row <> 0 Then
    
        sProduto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 182794
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        If Not (Me.ActiveControl Is Produto) Then
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
        
            'verifica se precisa preencher o grid com uma nova linha
            If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
                objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
            End If
        
        End If
        
        Call ComandoSeta_Fechar(Me.Name)
        
        Me.Show
        
    End If

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 182794
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182766)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProduto_Click()
    
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 182795

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProduto_Click:

    Select Case gErr

        Case 182795
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182767)

    End Select

    Exit Sub
    
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long
Dim iIndice As Integer
Dim objRelBarCode As ClassRelBarCode
Dim colRelBarCode As New Collection

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(colRelBarCode)
    If lErro <> SUCESSO Then gError 182796
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 182797

    lErro = objRelOpcoes.IncluirParametro("NNUMLIN", CStr(colRelBarCode.Count))
    If lErro <> AD_BOOL_TRUE Then gError 182798

    lErro = objRelOpcoes.IncluirParametro("NTIPO", CStr(Codigo_Extrai(Tipo.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 182799

    iIndice = 0
    For Each objRelBarCode In colRelBarCode
    
        iIndice = iIndice + 1
         
        lErro = objRelOpcoes.IncluirParametro("TPROD" & CStr(iIndice), objRelBarCode.sProduto)
        If lErro <> AD_BOOL_TRUE Then gError 182800
        
        lErro = objRelOpcoes.IncluirParametro("NQUANT" & CStr(iIndice), CStr(objRelBarCode.iQuantidade))
        If lErro <> AD_BOOL_TRUE Then gError 182801
        
    Next
    
    If bExecutar Then
    
        lErro = CF("RelBarCode_Prepara", colRelBarCode, lNumIntRel)
        If lErro <> SUCESSO Then gError 182802
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 182803

    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 182796 To 182803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182768)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 182804

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 182805

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 182804
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 182805, 182806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182769)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 182807
    
    'Pega o Tsk de acordo com o tipo
    lErro = NomeTsk_Le()
    If lErro <> SUCESSO Then gError 182808

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 182807, 182808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182770)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 182810

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 182811

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 182812
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 182813
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 182810
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 182811 To 182813

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182771)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182772)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ANALISE_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Etiquetas para Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBarCode"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

Dim sTexto As String
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Produto Then
            Call BotaoProduto_Click
        End If
                
    ElseIf (KeyCode = vbKeyV And Shift = 2) Then
                
        If Me.ActiveControl Is Produto Then
            Produto.Text = Clipboard.GetText
        End If
        
        KeyCode = 0
                
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

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

Dim lErro As Long

On Error GoTo Erro_GridItens_RowColChange

    Call Grid_RowColChange(objGridItens)
    
    If (GridItens.Row <> iLinhaAntiga) Then

        'Guarda a Linha corrente
        iLinhaAntiga = GridItens.Row

    End If

    Exit Sub

Erro_GridItens_RowColChange:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182773)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub GridItens_LostFocus()

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridItens_KeyDown

    'Guarda iLinhasExistentes
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then
        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridItens.Row
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182774)
    
    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

Dim X As Clipboard

    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

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

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 182814

                Case iGrid_Quantidade_Col

                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 182815

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 182816

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 182814 To 182815

        Case 182816
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182775)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objProduto As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 182817
    
    If objControl.Name = "Produto" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
        
    ElseIf objControl.Name = "Descricao" Then

        objControl.Enabled = False

    ElseIf objControl.Name = "Quantidade" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If
    
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 182817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182776)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Copias")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (Quantidade.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_Quantidade_Col = 3

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_PRODUTOS_COTACAO + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se o campo foi preenchido
    If Len(Trim(Quantidade.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 182818
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182819

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr
        
        Case 182818, 182819
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182777)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodProduto As String
Dim iLinha As Integer
Dim objProduto As ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
                
    sCodProduto = Produto.Text
        
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 182820
    
    'Se o campo foi preenchido
    If Len(sProdutoFormatado) > 0 Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 182821
                
'        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
'        Produto.PromptInclude = True
                
        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridItens.Row Then
                                                    
                If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado Then
'                    Produto.PromptInclude = False
'                    Produto.Text = ""
'                    Produto.PromptInclude = True
                    gError 182824
                    
                End If
                    
            End If
                           
        Next
        
        Set objProduto = New ClassProduto

        objProduto.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 182822
        
        If lErro <> SUCESSO Then gError 182831
        
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 182832
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao

        'verifica se precisa preencher o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182823

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr
        
        Case 182820 To 182823
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 182824
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO", gErr, sProdutoMascarado, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 182831
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 182832
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182778)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Carrega_Tipo(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Tipo

    'carregar tipos de relatório de etiquetas
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPO_REL_ETIQUETA, objComboBox)
    If lErro <> SUCESSO Then gError 182825
    
    giTipoPadrao = objComboBox.ListIndex

    Carrega_Tipo = SUCESSO

    Exit Function

Erro_Carrega_Tipo:

    Carrega_Tipo = gErr

    Select Case gErr
    
        Case 182825

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182779)

    End Select

    Exit Function

End Function

Private Function NomeTsk_Le() As Long

Dim lErro As Long
Dim objCampoGenVal As New ClassCamposGenericosValores

On Error GoTo Erro_NomeTsk_Le

    objCampoGenVal.lCodValor = Codigo_Extrai(Tipo.Text)
    objCampoGenVal.lCodCampo = CAMPOSGENERICOS_TIPO_REL_ETIQUETA

    lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCampoGenVal)
    If lErro <> SUCESSO Then gError 182826
    
    If Len(Trim(objCampoGenVal.sComplemento1)) > 0 Then
        gobjRelatorio.sNomeTsk = objCampoGenVal.sComplemento1
    Else
        gobjRelatorio.sNomeTsk = "BarCode"
    End If

    NomeTsk_Le = SUCESSO

    Exit Function

Erro_NomeTsk_Le:

    NomeTsk_Le = gErr

    Select Case gErr
    
        Case 182826

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182780)

    End Select

    Exit Function

End Function

