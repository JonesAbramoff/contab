VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPosCliOcx 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1215
   ScaleWidth      =   5820
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
      Left            =   3840
      Picture         =   "RelOpPosCliOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   165
      Width           =   1815
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      ItemData        =   "RelOpPosCliOcx.ctx":0102
      Left            =   885
      List            =   "RelOpPosCliOcx.ctx":0104
      TabIndex        =   1
      Top             =   735
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin VB.Label LabelCliente 
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
      Height          =   255
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   210
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   " Filial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   270
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   765
      Width           =   555
   End
End
Attribute VB_Name = "RelOpPosCliOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes
Dim iClienteAlterado As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFilial As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sFilial <> "" Then sExpressao = "Filial = " & Forprint_ConvInt(CInt(sFilial))
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171231)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sFilial As String

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23334

    'Verificar se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then Error 23335

    'Pegar parametros da tela
    lErro = objRelOpcoes.IncluirParametro("NCLIENTE", CStr(LCodigo_Extrai(Cliente.Text)))
    If lErro <> AD_BOOL_TRUE Then Error 23336
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTE", Cliente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54774
        
    lErro = objRelOpcoes.IncluirParametro("TFILIALCLI", Filial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54775
        
    If Filial.Text <> "" Then
        sFilial = CStr(LCodigo_Extrai(Filial.Text))
    Else
        sFilial = ""
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFilial)
    If lErro <> SUCESSO Then Error 23337

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23334, 23336, 23337, 54774, 54775

        Case 23335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err, Error$)
            Cliente.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171232)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29903
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 29903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171233)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23338

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23338

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171234)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iClienteAlterado = 1

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCria As Integer

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            iCria = 0 'Não deseja criar cliente caso não exista
            lErro = TP_cliente_Le2(Cliente, objCliente, iCria)
            If lErro <> SUCESSO Then Error 23340

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 23341

            'Preenche ComboBox de Filiais
            Filial.AddItem "<Empresa Toda>"
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True


    Select Case Err

        Case 23340, 23341
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171235)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se o cliente foi informado
    If Len(Trim(Cliente.Text)) = 0 Then Error 23343

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 23342

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 23345

        If lErro = 17660 Then Error 23346

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 23344

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 23342, 23345

        Case 23343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err, Error$)

        Case 23344, 23346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171236)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmCollCodigoNome
Dim objCodigoDescricao As AdmlCodigoNome

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCliente = New AdmEvento
    
'    'Preenche a listbox clientes
'    'Lê cada código e Nome Reduzido da tabela Clientes
'    lErro = CF("LCod_Nomes_Le", "clientes", "Codigo", "NomeReduzido", STRING_CLIENTE_NOME_REDUZIDO, colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 23339

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23339

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171237)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    If Len(Trim(Cliente.Text)) > 0 Then
        'Preenche com o Cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Posição do Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPosCli"
    
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

Public Sub Unload(objme As Object)
    
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
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub


Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

