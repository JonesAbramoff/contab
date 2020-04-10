VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TiposDeMaodeObra 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   8280
   Begin VB.CommandButton BotaoCertificados 
      Caption         =   "Certificados Adquiridos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6300
      TabIndex        =   22
      ToolTipText     =   "Abre o Browse de Centros de Trabalhos que utilizam este Tipo"
      Top             =   3990
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton BotaoCursos 
      Caption         =   "Cursos Inscritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4410
      TabIndex        =   21
      ToolTipText     =   "Abre o Browse de Centros de Trabalhos que utilizam este Tipo"
      Top             =   4005
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton BotaoCT 
      Caption         =   "Centros de Trabalho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2250
      TabIndex        =   18
      ToolTipText     =   "Abre o Browse de Centros de Trabalhos que utilizam este Tipo"
      Top             =   4005
      Width           =   2100
   End
   Begin VB.CommandButton BotaoMaquinas 
      Caption         =   "Máquinas, Habilidades e Processos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   105
      TabIndex        =   17
      ToolTipText     =   "Abre o Browse de Máquinas que utilizam este Tipo"
      Top             =   4005
      Width           =   2100
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1935
      Picture         =   "TiposDeMaodeObra.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Numeração Automática"
      Top             =   585
      Width           =   300
   End
   Begin VB.ListBox Tipos 
      Height          =   2010
      ItemData        =   "TiposDeMaodeObra.ctx":00EA
      Left            =   5520
      List            =   "TiposDeMaodeObra.ctx":00EC
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1095
      Width           =   2625
   End
   Begin VB.TextBox Observacao 
      Height          =   930
      Left            =   1380
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2190
      Width           =   4005
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6075
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TiposDeMaodeObra.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TiposDeMaodeObra.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TiposDeMaodeObra.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TiposDeMaodeObra.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   570
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   1110
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CustoHora 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   1680
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Prod 
      Height          =   315
      Left            =   1365
      TabIndex        =   5
      Top             =   3345
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label DescProd 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3390
      TabIndex        =   20
      Top             =   3345
      Width           =   4395
   End
   Begin VB.Label LabelProduto 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   330
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   19
      Top             =   3375
      Width           =   870
   End
   Begin VB.Label LabelCustoHora 
      Caption         =   "Custo p/ hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   16
      Top             =   1695
      Width           =   1275
   End
   Begin VB.Label Label13 
      Caption         =   "Tipos de Mão-de-Obra"
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
      Left            =   5505
      TabIndex        =   15
      Top             =   825
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      Height          =   195
      Left            =   225
      TabIndex        =   14
      Top             =   2235
      Width           =   1095
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   420
      TabIndex        =   13
      Top             =   1155
      Width           =   930
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
      Left            =   705
      TabIndex        =   12
      Top             =   615
      Width           =   660
   End
End
Attribute VB_Name = "TiposDeMaodeObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de Mão-de-Obra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposDeMaodeObra"

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

Private Sub BotaoCT_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCT As ClassCentrodeTrabalho
Dim iCodigo As Integer
Dim sFiltro As String

On Error GoTo Erro_BotaoCT_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 139105
    
    iCodigo = StrParaInt(Codigo.Text)
    
    Set objCT = New ClassCentrodeTrabalho
    
    objCT.iFilialEmpresa = giFilialEmpresa

    sFiltro = "NumIntDoc IN (SELECT NumIntDocCT FROM CTOperadores WHERE TipoMaoDeObra = ? )"

    colSelecao.Add iCodigo

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCT, Nothing, sFiltro)

    Exit Sub

Erro_BotaoCT_Click:

    Select Case gErr
    
        Case 139105
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174959)

    End Select

    Exit Sub

End Sub


Private Sub BotaoMaquinas_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMaquinas As ClassMaquinas
Dim iCodigo As Integer
Dim sFiltro As String

On Error GoTo Erro_BotaoMaquinas_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 137931
    
    iCodigo = StrParaInt(Codigo.Text)
    
    Set objMaquinas = New ClassMaquinas
    
    objMaquinas.iFilialEmpresa = giFilialEmpresa

    sFiltro = "NumIntDoc IN (SELECT NumIntDocMaq FROM MaquinaOperadores WHERE TipoMaoDeObra = ? )"

    colSelecao.Add iCodigo

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, Nothing, sFiltro)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr
    
        Case 137931
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174960)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para um TiposDeMaodeObra
    lErro = CF("TiposDeMaodeObra_Automatico", iCodigo)
    If lErro <> SUCESSO Then gError 137556
    
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 137556
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174961)
    
    End Select

    Exit Sub

End Sub

Private Sub Tipos_DblClick()

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Tipos_DblClick

    'Guarda o valor do codigo do Tipo da Mão-de-Obra selecionado na ListBox Tipos
    objTiposDeMaodeObra.iCodigo = Tipos.ItemData(Tipos.ListIndex)

    'Mostra os dados do TiposDeMaodeObra na tela
    lErro = Traz_TiposDeMaodeObra_Tela(objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137557

    Me.Show
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Tipos_DblClick:

    Tipos.SetFocus

    Select Case gErr

    Case 137557
        'erro tratado na rotina chamada
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174962)

    End Select

    Exit Sub

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoProduto = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174963)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set colCodigoDescricao = New AdmColCodigoNome
    Set objEventoProduto = New AdmEvento

    'Lê o Código e a Descrição de cada Tipo de Mão-de-Obra
    lErro = CF("Cod_Nomes_Le", "TiposDeMaodeObra", "Codigo", "Descricao", STRING_TIPO_MO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 137558

    'preenche a ListBox Tipos com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Tipos.AddItem objCodigoDescricao.sNome
        Tipos.ItemData(Tipos.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Prod)
    If lErro <> SUCESSO Then gError 137558
    
    If gobjEST.iExibeMOCursos = MARCADO Then
        BotaoCursos.Visible = True
        BotaoCertificados.Visible = True
    End If
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137558
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174964)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTiposDeMaodeObra Is Nothing) Then

        lErro = Traz_TiposDeMaodeObra_Tela(objTiposDeMaodeObra)
        If lErro <> SUCESSO Then gError 137560

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 137560
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174965)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objTiposDeMaodeObra.iCodigo = StrParaInt(Codigo.Text)
    objTiposDeMaodeObra.sDescricao = Descricao.Text
    objTiposDeMaodeObra.sObservacao = Observacao.Text
    objTiposDeMaodeObra.dCustoHora = StrParaDbl(CustoHora.Text)

    lErro = CF("Produto_Formata", Prod.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134080

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objTiposDeMaodeObra.sProduto = sProdutoFormatado

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 134080

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174966)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDeMaodeObra"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137561

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTiposDeMaodeObra.iCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 137561
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174967)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Tela_Preenche

    objTiposDeMaodeObra.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTiposDeMaodeObra.iCodigo <> 0 Then
        lErro = Traz_TiposDeMaodeObra_Tela(objTiposDeMaodeObra)
        If lErro <> SUCESSO Then gError 137562
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 137562
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174968)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo esta vazio
    If Len(Trim(Codigo.Text)) = 0 Then gError 137563

    'Verifica se Descricao esta vazio
    If Len(Trim(Descricao.Text)) = 0 Then gError 137564
    
    'Preenche o objTiposDeMaodeObra
    lErro = Move_Tela_Memoria(objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137565
    
    lErro = Trata_Alteracao(objTiposDeMaodeObra, objTiposDeMaodeObra.iCodigo)
    If lErro <> SUCESSO Then gError 137684

    'Grava o/a TiposDeMaodeObra no Banco de Dados
    lErro = CF("TiposDeMaodeObra_Grava", objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137566

    'Remove o item da lista de Tipos
    Call Tipos_Exclui(objTiposDeMaodeObra.iCodigo)

    'Insere o item na lista de Tipos
    Call Tipos_Adiciona(objTiposDeMaodeObra)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137563
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 137564
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TIPOSMO_NAO_PREENCHIDO", gErr)
            Descricao.SetFocus

        Case 137565, 137566, 137684
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174969)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TiposDeMaodeObra() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TiposDeMaodeObra
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0
    
    DescProd.Caption = ""

    Limpa_Tela_TiposDeMaodeObra = SUCESSO

    Exit Function

Erro_Limpa_Tela_TiposDeMaodeObra:

    Limpa_Tela_TiposDeMaodeObra = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174970)

    End Select

    Exit Function

End Function

Function Traz_TiposDeMaodeObra_Tela(objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_TiposDeMaodeObra_Tela

    'Lê o TiposDeMaodeObra que está sendo Passado
    lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
    If lErro <> SUCESSO And lErro <> 137598 Then gError 137567

    If lErro = SUCESSO Then

        'Limpa a Tela
        Call Limpa_Tela_TiposDeMaodeObra
        
        If objTiposDeMaodeObra.iCodigo <> 0 Then Codigo.Text = CStr(objTiposDeMaodeObra.iCodigo)
        Descricao.Text = objTiposDeMaodeObra.sDescricao
        Observacao.Text = objTiposDeMaodeObra.sObservacao
        If objTiposDeMaodeObra.dCustoHora <> 0 Then CustoHora.Text = Format(objTiposDeMaodeObra.dCustoHora, "Standard")
    
        If Len(Trim(objTiposDeMaodeObra.sProduto)) > 0 Then
            lErro = Mascara_RetornaProdutoEnxuto(objTiposDeMaodeObra.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 137567
            
            Prod.PromptInclude = False
            Prod.Text = sProdutoMascarado
            Prod.PromptInclude = True
            Call Prod_Validate(bSGECancelDummy)
        End If
    End If

    iAlterado = 0

    Traz_TiposDeMaodeObra_Tela = SUCESSO

    Exit Function

Erro_Traz_TiposDeMaodeObra_Tela:

    Traz_TiposDeMaodeObra_Tela = gErr

    Select Case gErr

        Case 137567
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174971)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 137568

    'Limpa Tela
    Call Limpa_Tela_TiposDeMaodeObra

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 137568
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174972)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174973)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 137569

    Call Limpa_Tela_TiposDeMaodeObra

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 137569
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174974)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo esta vazio
    If Len(Trim(Codigo.Text)) = 0 Then gError 137570

    objTiposDeMaodeObra.iCodigo = StrParaInt(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPOSDEMAODEOBRA", objTiposDeMaodeObra.iCodigo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui o Tipo de Mão-de-Obra
    lErro = CF("TiposDeMaodeObra_Exclui", objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137571

    Call Tipos_Exclui(objTiposDeMaodeObra.iCodigo)

    'Limpa Tela
    Call Limpa_Tela_TiposDeMaodeObra

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137570
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 137571
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174975)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.ClipText)) <> 0 Then

        'Critica a Codigo
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 137572

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 137572, 137573
            'erros traados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174976)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoHora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoHora_Validate

    'Veifica se CustoHora está preenchida
    If Len(Trim(CustoHora.Text)) <> 0 Then

       'Critica a CustoHora
       lErro = Valor_Positivo_Critica(CustoHora.Text)
       If lErro <> SUCESSO Then gError 137574

    End If

    Exit Sub

Erro_CustoHora_Validate:

    Cancel = True

    Select Case gErr

        Case 137574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174977)

    End Select

    Exit Sub

End Sub

Private Sub CustoHora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CustoHora, iAlterado)
    
End Sub

Private Sub CustoHora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipos_Adiciona(objTiposDeMaodeObra As ClassTiposDeMaodeObra)

    Tipos.AddItem objTiposDeMaodeObra.sDescricao
    Tipos.ItemData(Tipos.NewIndex) = objTiposDeMaodeObra.iCodigo

End Sub

Private Sub Tipos_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Tipos.ListCount - 1

        If Tipos.ItemData(iIndice) = iCodigo Then

            Tipos.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub


Private Sub Prod_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Prod_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_Prod_Validate

    lErro = CF("Produto_Perde_Foco", Prod, DescProd)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 202440
    
    If lErro <> SUCESSO Then gError 202441

    Exit Sub

Erro_Prod_Validate:

    Cancel = True

    Select Case gErr

        Case 202440

        Case 202441
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202442)

    End Select

    Exit Sub
End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Prod.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134507

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    sFiltro = "Ativo = ? "
    
    colSelecao.Add PRODUTO_ATIVO
        
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sFiltro)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 134507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174528)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134508

    Prod.PromptInclude = False
    Prod.Text = sProdutoMascarado
    Prod.PromptInclude = True
    
    Call Prod_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 134508
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174529)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Prod Then Call LabelProduto_Click
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub

Private Sub BotaoCursos_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim iCodigo As Integer
Dim sFiltro As String

On Error GoTo Erro_BotaoCursos_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 137931
    
    iCodigo = StrParaInt(Codigo.Text)
    
    sFiltro = "CodMO = ?"

    colSelecao.Add iCodigo

    Call Chama_Tela("CursoMOLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub

Erro_BotaoCursos_Click:

    Select Case gErr
    
        Case 137931
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174960)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCertificados_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim iCodigo As Integer
Dim sFiltro As String

On Error GoTo Erro_BotaoCertificados_Click

    If Len(Trim(Codigo.ClipText)) = 0 Then gError 137931
    
    iCodigo = StrParaInt(Codigo.Text)
    
    sFiltro = "CodMO = ?"

    colSelecao.Add iCodigo

    Call Chama_Tela("CertificadoMOLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub

Erro_BotaoCertificados_Click:

    Select Case gErr
    
        Case 137931
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOSDEMAODEOBRA_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174960)

    End Select

    Exit Sub

End Sub
