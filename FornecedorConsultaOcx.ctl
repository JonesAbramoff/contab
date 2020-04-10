VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FornecedorConsultaOcx 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   8085
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   5730
      ScaleHeight     =   705
      ScaleWidth      =   1995
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   2055
      Begin VB.CommandButton BotaoConsulta 
         Height          =   585
         Left            =   90
         Picture         =   "FornecedorConsultaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1275
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   585
         Left            =   1425
         Picture         =   "FornecedorConsultaOcx.ctx":1DC2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.TextBox Descricao 
      Enabled         =   0   'False
      Height          =   795
      Left            =   4110
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4350
      Width           =   3675
   End
   Begin MSComctlLib.ImageList ImageListConsulta 
      Left            =   3390
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FornecedorConsultaOcx.ctx":1F40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TvwConsultas 
      Height          =   3075
      Left            =   4110
      TabIndex        =   2
      Top             =   1170
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   5424
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageListConsulta"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageListModulo 
      Left            =   3390
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FornecedorConsultaOcx.ctx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FornecedorConsultaOcx.ctx":2B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FornecedorConsultaOcx.ctx":2E5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListaModulo 
      Height          =   3855
      Left            =   300
      TabIndex        =   1
      Top             =   1170
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   6800
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageListModulo"
      SmallIcons      =   "ImageListModulo"
      ColHdrIcons     =   "ImageListModulo"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   330
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label FornecedorEtiqueta 
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
      Left            =   330
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   390
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
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
      Left            =   270
      TabIndex        =   8
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Consulta:"
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
      Left            =   4110
      TabIndex        =   9
      Top             =   960
      Width           =   810
   End
End
Attribute VB_Name = "FornecedorConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'CONSTANTE GLOBAL DA TELA
Private Const NOME_TELA_CONSULTA_FORNECEDOR = "FornecedorConsulta"

Private gcolConsultas As Collection

Private Sub BotaoConsulta_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoConsulta_Click
    
    If (TvwConsultas.SelectedItem Is Nothing) Then Error 60467
    
    Call TvwConsultas_DblClick
    
    Exit Sub
    
Erro_BotaoConsulta_Click:
    
    Select Case Err
        
        Case 60467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_ARVORE_CONSULTA_NAO_SELECIONADO", Err, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160564)

    End Select

    Exit Sub
        
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le2(Fornecedor, objFornecedor)
            If lErro <> SUCESSO Then Error 60468

        End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case Err

        Case 60468 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160565)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorEtiqueta_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Fornecedor com o Fornecedor da Tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub ListaModulo_Click()

Dim sSiglaModulo As String
Dim lErro As Long

On Error GoTo Erro_ListaModulo_Click

    If ListaModulo.ListItems.Count = 0 Then Exit Sub
    
    'Obtem a Sigla através do Nome
    sSiglaModulo = gcolModulo.Sigla(ListaModulo.SelectedItem.Text)

    Call Carrega_Consultas(sSiglaModulo)

    Exit Sub

Erro_ListaModulo_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160566)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1

    'Preenche o Fornecedor com o Fornecedor selecionado
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(Cancel)

    Me.Show

    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160567)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela(Me)

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160568)

     End Select

     Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolConsultas = Nothing
    Set objEventoFornecedor = Nothing
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gcolConsultas = New Collection
    Set objEventoFornecedor = New AdmEvento
    
    'Le para a Colecao global todos os Modulos e suas Consultas
    lErro = CF("Consultas_Le_Todos", NOME_TELA_CONSULTA_FORNECEDOR, gcolConsultas)
    If lErro <> SUCESSO Then Error 60469

    'Carrega a lista de Modulos
    Call Carrega_Modulos

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 60469 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160569)

    End Select

    Exit Sub

End Sub

Sub Carrega_Modulos()

Dim objConsulta As ClassConsultas
Dim sSiglaAnterior As String
Dim collCodigoNome As New AdmCollCodigoNome
Dim objUsuarioModulo As New ClassUsuarioModulo
Dim objUsuario As New ClassUsuarios
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Modulos

    objUsuario.sCodUsuario = gsUsuario
    
    lErro = CF("Usuarios_Le", objUsuario)
    If lErro <> SUCESSO Then Error 62297
      
   objUsuarioModulo.dtDataValidade = objUsuario.dtDataValidade
   objUsuarioModulo.iCodFilial = giFilialEmpresa
   objUsuarioModulo.lCodEmpresa = glEmpresa
   objUsuarioModulo.sCodUsuario = gsUsuario
   
    'Lê os Módulos
    lErro = CF("UsuarioModulos_Le", objUsuarioModulo, collCodigoNome)
    If lErro <> SUCESSO Then Error 43630
    
    For Each objConsulta In gcolConsultas

        If sSiglaAnterior <> objConsulta.sSigla Then
        
            For iIndice = 1 To collCodigoNome.Count
                
                If collCodigoNome(iIndice).sNome = gcolModulo.Nome(objConsulta.sSigla) Then
                    ListaModulo.ListItems.Add , , gcolModulo.Nome(objConsulta.sSigla), objConsulta.iIconeModulo, objConsulta.iIconeModulo
                    sSiglaAnterior = objConsulta.sSigla
                    ListaModulo.Refresh
                End If
            Next
        End If

    Next

    Exit Sub

Erro_Carrega_Modulos:

    Select Case Err
    
        Case 62297, 43630
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160570)
    
    End Select
    
    Exit Sub
    
End Sub

Function Carrega_Consultas(sSiglaModulo As String)

    'Limpa a Arvore
    TvwConsultas.Nodes.Clear
    Descricao.Text = ""
    
    'Carrega a Arvore Através da função recursiva abaixo
    Call Arvore_Empilha(1, 1, sSiglaModulo)

End Function

Sub Arvore_Empilha(iPosicaoPai As Integer, iPosicaoAtual As Integer, sSiglaModulo As String)

Dim objConsulta As ClassConsultas
Dim iPosicaoAnterior As Integer
Dim iNivelAnterior As Integer

    'Enquanto não atingir o total da Consulta
    Do While (iPosicaoAtual <= gcolConsultas.Count)

        'Pega uma consulta na coleção
        Set objConsulta = gcolConsultas.Item(iPosicaoAtual)

        'Se o NivelAnterior for "0" isto é se esta iniciando pega o Nivel
        If iNivelAnterior = 0 Then iNivelAnterior = objConsulta.iNivel

        'Se o Nivel diminui então --> Sai
        If objConsulta.iNivel < iNivelAnterior Then Exit Do

        'Se a Sigla for igual a passada
        If sSiglaModulo = objConsulta.sSigla Then

            'Se o Nivel aumentou ---> empilha
            If objConsulta.iNivel > iNivelAnterior Then

                'Empilha passado o Pai
                Call Arvore_Empilha(iPosicaoAnterior, iPosicaoAtual, sSiglaModulo)

            Else
                'Se o Nivel permanecer o mesmo

                'Se o Nivel for 1
                If objConsulta.iNivel = 1 Then
                    'Adiciona sem PAI
                    Call TvwConsultas.Nodes.Add(, tvwLast, , objConsulta.sConsulta, objConsulta.iIconeConsulta)
                Else
                    'Se não Adiciona pegando o Pai Passado
                    Call TvwConsultas.Nodes.Add(iPosicaoPai, tvwChild, , objConsulta.sConsulta, objConsulta.iIconeConsulta)
                End If

                'Adiciona a Posicao e Atualiza o Nivel e Posicao Anteriores
                iPosicaoAtual = iPosicaoAtual + 1
                iNivelAnterior = objConsulta.iNivel
                iPosicaoAnterior = objConsulta.iPosicao

            End If
        Else
            'Caso a Sigla não seja igual
            'Adiciona 1 para pegar a proxima consulta na coleção
            iPosicaoAtual = iPosicaoAtual + 1

        End If

    Loop

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160571)

    End Select

    Exit Function

End Function

Private Sub TvwConsultas_DblClick()

Dim lErro As Long
Dim objConsulta As ClassConsultas
Dim colSelecao As New Collection
Dim vParametro As Variant
Dim objEventoTemp As AdmEvento
Dim objTemp As Object

On Error GoTo Erro_TvwConsultas_DblClick
    
    'Se não tem filhos (Folha na árvore)
    If TvwConsultas.SelectedItem.Children = 0 Then
        
        'Se o Fornecedor não está Preenchido --> ERRO
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 60474
        
        'Procura a Consulta na coleção
        For Each objConsulta In gcolConsultas
            
            'Se encontrou
            If (objConsulta.sSigla = gcolModulo.Sigla(ListaModulo.SelectedItem.Text)) And (objConsulta.sConsulta = TvwConsultas.SelectedItem.Text) Then
                
                'Adiciona o Filtro
                colSelecao.Add (LCodigo_Extrai(Fornecedor.Text))
                
                'O Parametro é passado
                vParametro = "Fornecedor = ?"
                
                'Chama a Tela descrita pelo
                Call Chama_Tela(objConsulta.sTelaRelacionada, colSelecao, objTemp, objEventoTemp, vParametro)
            
                Exit For
            
            End If
        
        Next
    
    End If
    
    Exit Sub

Erro_TvwConsultas_DblClick:

    Select Case Err
        
        Case 60474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160572)

    End Select

    Exit Sub

End Sub

Private Sub TvwConsultas_NodeClick(ByVal Node As MSComctlLib.Node)

'Atualiza a Descrição das Consultas
Dim lErro As Long
Dim objConsulta As ClassConsultas

On Error GoTo Erro_TvwConsultas_NodeClick

    For Each objConsulta In gcolConsultas
        
        If (objConsulta.sSigla = gcolModulo.Sigla(ListaModulo.SelectedItem.Text)) And (objConsulta.sConsulta = Node.Text) Then
                        
            Descricao.Text = objConsulta.sDescricao
            Exit For
        End If
        
    Next
    
    Exit Sub

Erro_TvwConsultas_NodeClick:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160573)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FORNECEDOR_CONSULTA
    Set Form_Load_Ocx = Me
    Caption = "Consulta por Fornecedor"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ConsultaFornecedor"

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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub FornecedorEtiqueta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorEtiqueta, Source, X, Y)
End Sub

Private Sub FornecedorEtiqueta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorEtiqueta, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_Change()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Change
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134035

    Exit Sub

Erro_Fornecedor_Change:

    Select Case gErr

        Case 134035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160574)

    End Select
    
    Exit Sub

End Sub

