VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ProdutoDesconto 
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
   KeyPreview      =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   5730
   Begin VB.Frame Frame1 
      Caption         =   "Desconto"
      Height          =   930
      Left            =   225
      TabIndex        =   10
      Top             =   1770
      Width           =   5160
      Begin MSMask.MaskEdBox Desconto 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   405
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescValor 
         Height          =   315
         Left            =   3045
         TabIndex        =   13
         Top             =   390
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   2445
         TabIndex        =   14
         Top             =   450
         Width           =   510
      End
      Begin VB.Label LabelDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Percentual:"
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
         Left            =   375
         TabIndex        =   12
         Top             =   435
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   1575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1635
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "ProdutoDesconto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "ProdutoDesconto.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ProdutoDesconto.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1725
      TabIndex        =   9
      Top             =   840
      Width           =   3225
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
      Left            =   855
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   8
      Top             =   390
      Width           =   735
   End
   Begin VB.Label LabelDescricao 
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
      Left            =   660
      TabIndex        =   7
      Top             =   900
      Width           =   930
   End
   Begin VB.Label LabelNomeReduz 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido :"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1410
      Width           =   1470
   End
   Begin VB.Label NomeReduzido 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1725
      TabIndex        =   5
      Top             =   1350
      Width           =   2625
   End
End
Attribute VB_Name = "ProdutoDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'Declaração utilizada para evento LabelProduto_Click
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    
    'Parent.HelpContextID = IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO
    Set Form_Load_Ocx = Me
    Caption = "Desconto em Item"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ProdutoDesconto"

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

Private Sub Desconto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
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
Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa a máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 113300
    
    'Inicializa o objEventoProduto
    Set objEventoProduto = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
    
        Case 113300
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165596)
            
    End Select
    
    Exit Sub
    
End Sub

'### Inicio dos Tratamentos do Browser's

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 113301

        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 113301

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165597)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProdFilial As New ClassProdutoFilial

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Função que lê no banco de Dados os Produtos para Uma determinada Filial com o Seu Respectivo Desconto
    lErro = CF("DescProdutosFilial_Le", objProduto, objProdFilial)
    If lErro <> SUCESSO And lErro <> 113313 Then gError 113318
    
    'Se não achou o Produto --> erro
    If lErro = 113313 Then gError 113304

    'Função que Preenche a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 113305
    
    'Prenche o Nome Reduzido
    NomeReduzido.Caption = objProduto.sNomeReduzido

    If objProdFilial.dDescontoItem > 0 Then

        'Preenche o Campo com o Desconto e Valida os Dados
        Desconto.Text = objProdFilial.dDescontoItem
    
    Else
    
        Desconto.Text = ""
        
    End If
    
    
    If objProdFilial.dDescontoValor > 0 Then
    
        DescValor.Text = objProdFilial.dDescontoValor

    Else
    
        DescValor.Text = ""
        
    End If

    
    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 113303

        Case 113304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case 113305, 113318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165598)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objProdFilial As New ClassProdutoFilial

On Error GoTo Erro_Produto_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", Produto.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113306

    'se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126819
        
        'se produto estiver preenchido
        If iProdPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial ==> erro
            If lErro = 25043 Then gError 126820
            
            'se nao está cadastrado
            If lErro <> SUCESSO Then gError 126821

        End If
        
    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 113306, 113307, 126819

        Case 126820
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, Produto.Text)

        Case 126821
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165599)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 113319

    'Limpa a Tela
    Call BotaoLimpar_Click

    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 113319

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165600)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProdFilial As New ClassProdutoFilial

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProdutoFilial"

    'Lê os atributos de objRastroLote que aparecem na Tela
    lErro = Move_Tela_Memoria(objProduto, objProdFilial)
    If lErro <> SUCESSO Then gError 113320

    'no BD no caso de STRING e Key igual ao nome do campo
    'Por Daniel em 17/02/03
    colCampoValor.Add "Produto", objProduto.sCodigo, STRING_PRODUTO, "Produto"
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    
    'Filtros para o Sistema de Setas e Para Browser
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 113320

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165601)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objProdutoDesc As New ClassProduto
Dim objProdFilial As New ClassProdutoFilial

On Error GoTo Erro_Tela_Preenche

   'Passa os dados da coleção para objRastroLote
    objProdutoDesc.sCodigo = colCampoValor.Item("Produto").vValor
        
    'Função que lê no banco de Dados os Produtos para Uma determinada Filial com o Seu Respectivo Desconto
    lErro = CF("DescProdutosFilial_Le", objProdutoDesc, objProdFilial)
    If lErro <> SUCESSO And lErro <> 113313 Then gError 113318
    
    'Se não achou o Produto --> erro
    If lErro = 113313 Then gError 113304
    
    'Traz dados do RastreamentoLote para a tela
    lErro = Traz_ProdutoDescTela_Tela(objProdutoDesc, objProdFilial)
    If lErro <> SUCESSO Then gError 113323

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 113323, 113318
         
        Case 113304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", gErr, objProdutoDesc.sCodigo, giFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165602)

    End Select

    Exit Sub

End Sub

Function Traz_ProdutoDescTela_Tela(objProdutoDesc As ClassProduto, objProdFilial As ClassProdutoFilial) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_ProdutoDescTela_Tela

    'Função que Preenche a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", objProdutoDesc.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 113324
    
    'Prenche o Nome Reduzido
    NomeReduzido.Caption = objProdutoDesc.sNomeReduzido

    If objProdFilial.dDescontoItem > 0 Then

        'Preenche o Campo com o Desconto e Valida os Dados
        Desconto.Text = objProdFilial.dDescontoItem
    
    Else
    
        Desconto.Text = ""
        
    End If
    
    
    If objProdFilial.dDescontoValor > 0 Then
    
        DescValor.Text = objProdFilial.dDescontoValor

    Else
    
        DescValor.Text = ""
        
    End If
    
    Traz_ProdutoDescTela_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_ProdutoDescTela_Tela:

    Traz_ProdutoDescTela_Tela = gErr
    
    Select Case gErr

        Case 113324

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165603)

    End Select

    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProdFilial As New ClassProdutoFilial

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios da Tela estão preenchidos
    If Len(Trim(Produto.ClipText)) = 0 Then gError 113325
        
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objProduto, objProdFilial)
    If lErro <> SUCESSO Then gError 113326

    'Grava o Produto
    lErro = CF("Produto_GravaDesc", objProduto, objProdFilial)
    If lErro <> SUCESSO Then gError 113328
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 113325
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 113326, 113328
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165604)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objProduto As ClassProduto, objProdFilial As ClassProdutoFilial) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Código foi preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 113336

        If iPreenchido = PRODUTO_VAZIO Then Error 113337

        objProduto.sCodigo = sProduto
        
        
    End If
    
    'Recolhe os demais dados
    objProduto.sDescricao = Descricao.Caption
    objProduto.sNomeReduzido = NomeReduzido.Caption
    objProdFilial.dDescontoItem = StrParaDbl(Desconto.Text)
    objProdFilial.dDescontoValor = StrParaDbl(DescValor.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 113336
        
        Case 113337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165605)

    End Select

    Exit Function

End Function

Private Sub Desconto_Validate(Cancel As Boolean)
'Função que Valida a Porcentagem do Desconto

Dim lErro As Long
 
On Error GoTo Erro_Desconto_Validate

    'Se o Desconto não estiver Preenchido então sai do Validate
    If Len(Trim(Desconto.Text)) <> 0 Then
    
        'Verefica se a Porcentagem é Valida
        lErro = Porcentagem_Critica2(Desconto.Text)
        If lErro <> SUCESSO Then gError 113312
        
        DescValor.Text = ""
    
    End If
    
    Exit Sub
    
Erro_Desconto_Validate:

    Cancel = True
    
    Select Case gErr
   
        Case 113312
        'Erro Tratado Dentro da Função
        
        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165606)

        End Select
        
   Exit Sub

End Sub

Private Sub DescValor_Validate(Cancel As Boolean)
'Função que Valida o Valor do Desconto

Dim lErro As Long
 
On Error GoTo Erro_DescValor_Validate

        
    If Len(Trim(DescValor.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(DescValor.Text)
        If lErro <> SUCESSO Then gError 126818
        
        Desconto.Text = ""
    
    End If
    
    Exit Sub
    
Erro_DescValor_Validate:

    Cancel = True
    
    Select Case gErr
   
        Case 126818
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165607)

        End Select
        
   Exit Sub

End Sub

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long


    Trata_Parametros = SUCESSO

    
End Function

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela(Me)

    Descricao.Caption = ""
    NomeReduzido.Caption = ""
    Desconto.Text = ""
    DescValor.Text = ""
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoProduto = Nothing
    
    Dim lErro As Long

    'Fecha o comando das setas se estiver aberto

    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Produto Then Call LabelProduto_Click
                
    End If

End Sub

