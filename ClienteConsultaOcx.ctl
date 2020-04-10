VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ClienteConsultaOcx 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   8085
   Begin VB.CommandButton BotaoVerCliente 
      Height          =   360
      Left            =   4920
      Picture         =   "ClienteConsultaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Abrir a tela de cadastro"
      Top             =   300
      Width           =   360
   End
   Begin VB.TextBox Descricao 
      Enabled         =   0   'False
      Height          =   795
      Left            =   4110
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4350
      Width           =   3675
   End
   Begin MSComctlLib.ImageList ImageListConsulta 
      Left            =   3360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":0BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":1046
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":11BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   5730
      ScaleHeight     =   705
      ScaleWidth      =   1995
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   2055
      Begin VB.CommandButton BotaoFechar 
         Height          =   585
         Left            =   1425
         Picture         =   "ClienteConsultaOcx.ctx":1616
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoConsulta 
         Height          =   585
         Left            =   90
         Picture         =   "ClienteConsultaOcx.ctx":1794
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1275
      End
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
      Left            =   3360
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
            Picture         =   "ClienteConsultaOcx.ctx":3556
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":3E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClienteConsultaOcx.ctx":415A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListaModulo 
      Height          =   3855
      Left            =   270
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
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   330
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
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
      TabIndex        =   7
      Top             =   960
      Width           =   810
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
      TabIndex        =   6
      Top             =   960
      Width           =   690
   End
   Begin VB.Label ClienteEtiqueta 
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
      Left            =   300
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   390
      Width           =   660
   End
End
Attribute VB_Name = "ClienteConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

'CONSTANTE GLOBAL DA TELA
Private Const NOME_TELA_CONSULTA_CLIENTE = "ClienteConsulta"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154190)

    End Select

    Exit Sub
        
End Sub

Private Sub BotaoVerCliente_Click()

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoVerCliente_Click

    objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    
    If objcliente.lCodigo <> 0 Then

        Call Chama_Tela("Clientes", objcliente)

    End If

    Exit Sub

Erro_BotaoVerCliente_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154191)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Critica e Preenche Cliente
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then Error 60468

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case Err

        Case 60468 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154191)

    End Select

    Exit Sub

End Sub

Private Sub ClienteEtiqueta_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    
    If objcliente.lCodigo <> 0 Then
    
        Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", "Código")
        
    Else
    
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text

        Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", "Nome Reduzido")
    
    End If

End Sub

Private Sub ListaModulo_Click()

Dim sSiglaModulo As String
Dim lErro As Long

On Error GoTo Erro_ListaModulo_Click

    'Obtem a Sigla através do Nome
    sSiglaModulo = gcolModulo.Sigla(ListaModulo.SelectedItem.Text)

    Call Carrega_Consultas(sSiglaModulo)

    Exit Sub

Erro_ListaModulo_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154192)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente, Cancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    Call Cliente_Validate(Cancel)

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154193)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154194)

     End Select

     Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolConsultas = Nothing

    Set objEventoCliente = Nothing

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    
    Set gcolConsultas = New Collection

    'Le para a Colecao global todos os Modulos e suas Consultas
    lErro = CF("Consultas_Le_Todos", NOME_TELA_CONSULTA_CLIENTE, gcolConsultas)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154195)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154196)
    
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

'Alterado por Luiz Nogueira em 13/01/04
Function Trata_Parametros(Optional objcliente As ClassCliente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    '*** Incluído por Luiz Nogueira em 13/01/04 - INÍCIO ***
    'Se foi passado um objCliente como parâmetro
    If Not (objcliente Is Nothing) Then
    
        Cliente.Text = objcliente.lCodigo
        Call Cliente_Validate(bSGECancelDummy)
    
    End If
    '*** Incluído por Luiz Nogueira em 13/01/04 - FIM ***

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154197)

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
    
    If TvwConsultas.SelectedItem.Children = 0 Then
        
        If Len(Trim(Cliente.Text)) = 0 Then Error 60474
        
        For Each objConsulta In gcolConsultas
        
            If (objConsulta.sSigla = gcolModulo.Sigla(ListaModulo.SelectedItem.Text)) And (objConsulta.sConsulta = TvwConsultas.SelectedItem.Text) Then
            
                colSelecao.Add (LCodigo_Extrai(Cliente.Text))
                                  
                vParametro = "Cliente = ?"
            
                Call Chama_Tela(objConsulta.sTelaRelacionada, colSelecao, objTemp, objEventoTemp, vParametro)
            
                Exit For
            
            End If
        
        Next
    
    End If
    
    Exit Sub

Erro_TvwConsultas_DblClick:

    Select Case Err
        
        Case 60474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154198)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154199)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_CLIENTE_CONSULTA
    Set Form_Load_Ocx = Me
    Caption = "Consulta por Cliente"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ConsultaCliente"

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

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Cliente Then Call ClienteEtiqueta_Click
    
    End If

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub ClienteEtiqueta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteEtiqueta, Source, X, Y)
End Sub

Private Sub ClienteEtiqueta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteEtiqueta, Button, Shift, X, Y)
End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134017

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134017

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154200)

    End Select
    
    Exit Sub

End Sub

