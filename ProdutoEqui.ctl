VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ProdutoEquiOcx 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   5385
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3000
      TabIndex        =   8
      Top             =   1620
      Width           =   1230
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   900
      TabIndex        =   7
      Top             =   1620
      Width           =   1230
   End
   Begin VB.Frame FrameProduto 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1530
      Left            =   210
      TabIndex        =   0
      Top             =   45
      Width           =   4935
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   75
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   " "
      End
      Begin VB.TextBox NomeReduzido 
         Height          =   312
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   6
         Top             =   855
         Width           =   1635
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   465
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   50
         PromptChar      =   " "
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
         Left            =   765
         TabIndex        =   1
         Top             =   105
         Width           =   660
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   510
         Width           =   930
      End
      Begin VB.Label LabelNomeReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
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
         Left            =   15
         TabIndex        =   5
         Top             =   915
         Width           =   1410
      End
   End
End
Attribute VB_Name = "ProdutoEquiOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gobjProduto As ClassProduto

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 131277

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 131278

    'Mostra os dados do Produto na tela
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError 131279

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 131277, 131279

        Case 131278
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165640)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

'    'Verifica se o produto foi preenchido
'    If Len(Codigo.ClipText) <> 0 Then
'
'        'Preenche o código de objProduto
'        lErro = CF("Produto_Formata", Codigo.Text, sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 131280
'
'        objProduto.sCodigo = sProdutoFormatado
'
'    End If
'
'    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case 131280

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165641)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

'    objProduto.sNomeReduzido = NomeReduzido.Text
'
'    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165642)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento
    
    iAlterado = 0
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Codigo)
    If lErro <> SUCESSO Then gError 131292
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 131292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165643)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then gError 131281

    End If
    
    Exit Sub

Erro_NomeReduzido_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 131281
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", gErr, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165644)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoEnxuto As String

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then
    
        Set gobjProduto = objProduto

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 131335
        
        Codigo.PromptInclude = False
        Codigo.Text = sProdutoEnxuto
        Codigo.PromptInclude = True
        
        NomeReduzido.Text = objProduto.sNomeReduzido
        
        Descricao.Text = objProduto.sDescricao
    
    Else
    
        Set gobjProduto = New ClassProduto
        
        Call Limpa_Tela_Produto
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 131335

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165645)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Codigo_Validate(Cancel As Boolean)
'se nao for produto do 1o nivel garantir que exista "pai" e este seja sintetico.
'Ex.Nao pode editar 1.1.2 se nao existir 1.1

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAntFormatado As String

On Error GoTo Erro_Codigo_Validate

    If Len(Codigo.ClipText) > 0 Then

        'critica o formato da Produto
        lErro = CF("Produto_Formata", Codigo.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 131282

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Verifica se o Produto tem um Produto "Pai" já está cadastrado
            lErro = CF("Produto_Critica_ProdutoPai", sProdutoFormatado, MODULO_ESTOQUE)
            If lErro <> SUCESSO Then gError 131283

        End If

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 131282 To 131283

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165646)

    End Select

    Exit Sub

End Sub

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Mostra os dados do Produto na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_Traz_Produto_Tela

    'Limpa a Tela
    Call Limpa_Tela_Produto

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 131285

    'Coloca o Codigo na tela
    Codigo.PromptInclude = False
    Codigo.Text = sProdutoEnxuto
    Codigo.PromptInclude = True
    
    'Coloca os demais dados do Produto na tela
    Descricao.Text = objProduto.sDescricao
    NomeReduzido.Text = objProduto.sNomeReduzido
   
    iAlterado = 0

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 131285
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165647)

    End Select

    Exit Function

End Function

Private Sub BotaoCancelar_Click()

    giRetornoTela = vbCancel

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    giRetornoTela = vbOK

    lErro = Move_Tela_Memoria(gobjProduto)
    If lErro <> SUCESSO Then gError 131290

    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 131290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165648)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_Produto()

Dim iIndice As Integer

    'Chama Limpa_Tela
    Call Limpa_Tela(Me)

End Sub

Private Function Move_Tela_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim colTabelas As New Collection
Dim iNivel As Integer
Dim sVerifica As String

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 131286

        If iPreenchido = PRODUTO_VAZIO Then gError 131287

        objProduto.sCodigo = sProduto
        
    End If
    
    'Recolhe os demais dados
    objProduto.sDescricao = Descricao.Text
    objProduto.sNomeReduzido = NomeReduzido.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 131286
        
        Case 131287
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165649)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Produtos"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then gError 131288

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProduto.sCodigo, STRING_PRODUTO, "Codigo"
    colCampoValor.Add "Descricao", objProduto.sDescricao, STRING_PRODUTO_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objProduto.sNomeReduzido, STRING_PRODUTO_NOME_REDUZIDO, "NomeReduzido"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 131288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165650)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iControleEstoque As Integer

On Error GoTo Erro_Tela_Preenche

    objProduto.sCodigo = colCampoValor.Item("Codigo").vValor

    If Len(Trim(objProduto.sCodigo)) <> 0 Then

        'Carrega objProduto com os dados passados em colCampoValor
        objProduto.sDescricao = colCampoValor.Item("Descricao").vValor
        objProduto.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError 131289
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 131289

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165651)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PRODUTO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Produto Equivalente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ProdutoEqui"
    
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
        
    If KeyCode = KEYCODE_BROWSER Then
    
                
    End If

End Sub

'Private Sub LabelDescricao_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(LabelDescricao, Source, X, Y)
'End Sub

'Private Sub LabelDescricao_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
'End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

