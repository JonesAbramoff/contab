VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpHistRecOcx 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ScaleHeight     =   900
   ScaleWidth      =   6225
   Begin MSMask.MaskEdBox ClienteDesde 
      Height          =   285
      Left            =   945
      TabIndex        =   1
      Top             =   510
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
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
      Left            =   4965
      Picture         =   "RelOpHistRecOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   945
      TabIndex        =   0
      Top             =   120
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2265
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   510
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   " Desde:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   780
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
      Left            =   270
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "RelOpHistRecOcx"
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


Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDesde As String) As Long
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

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169380)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sDesde As String, dtDataAberto As Date

On Error GoTo Erro_PreencherRelOp

    lErro = ParcelaRec_Le_MenorData(dtDataAberto, LCodigo_Extrai(Cliente.Text))
    If lErro <> SUCESSO Then gError 87589

    If dtDataAberto <> DATA_NULA And StrParaDate(ClienteDesde.Text) > dtDataAberto Then gError 87595
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 84445 '23334

    'Verificar se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 84446 '23335

    'Pegar parametros da tela
    lErro = objRelOpcoes.IncluirParametro("NCLIENTE", CStr(LCodigo_Extrai(Cliente.Text))) '???William
    If lErro <> AD_BOOL_TRUE Then gError 84447 '23336

    lErro = objRelOpcoes.IncluirParametro("TCLIENTE", Cliente.Text)
    If lErro <> AD_BOOL_TRUE Then gError 84448 '54774

    lErro = objRelOpcoes.IncluirParametro("DDATA", StrParaDate(ClienteDesde.Text))
    If lErro <> AD_BOOL_TRUE Then gError 84449 '54775

    If ClienteDesde.Text <> "" Then
        sDesde = ClienteDesde.Text

    Else
      sDesde = ""

    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDesde)
    If lErro <> SUCESSO Then gError 84450 '23337

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 84445, 84447, 84448, 84449, 84450, 87589

        Case 84446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADESDE_NAO_PREENCHIDA", gErr, Error$)
            Cliente.SetFocus

        Case 87595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOSREC_ABERTO", gErr, dtDataAberto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169381)

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

    If Not (gobjRelatorio Is Nothing) Then gError 84451 '29903

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 84452 '29903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169382)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 84453


    If Len(Trim(ClienteDesde.ClipText)) = 0 Then gError 84454


    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 84455 '23338

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 84455

        Case 84453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 84454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)


        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169383)

    End Select

        If gErr = 84453 Then Cliente.SetFocus

        If gErr = 84454 Then ClienteDesde.SetFocus


    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCria As Integer

On Error GoTo Erro_Cliente_Validate

        If Len(Trim(Cliente.Text)) > 0 Then

            iCria = 0 'Não deseja criar cliente caso não exista
            lErro = TP_Cliente_Le2(Cliente, objCliente, iCria)
            If lErro <> SUCESSO Then gError 84456

        End If


    Exit Sub

Erro_Cliente_Validate:

    Cancel = True


    Select Case gErr

        Case 84456

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169384)

    End Select

    Exit Sub

End Sub

Private Sub ClienteDesde_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String

On Error GoTo Erro_ClienteDesde_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(ClienteDesde.ClipText)) = 0 Then Exit Sub

    'Verifica se é uma data válida
    lErro = Data_Critica(ClienteDesde.Text)
    If lErro <> SUCESSO Then gError 84458

    'Verifica se a data informada é maoir que a data atual
    If StrParaDate(ClienteDesde.Text) > gdtDataAtual Then gError 84457

    Exit Sub

Erro_ClienteDesde_Validate:

    Cancel = True


    Select Case gErr

        Case 84458

        Case 84457
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INFORMADA_MENOR_DATA_HOJE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169385)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

    Set objEventoCliente = New AdmEvento
    lErro_Chama_Tela = SUCESSO

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
    Caption = "Histórico de Recebimentos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpHistRec"

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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(ClienteDesde, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 84460

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 84460
            ClienteDesde.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169386)

    End Select

    Exit Sub


End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(ClienteDesde, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 84459

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 84459
            ClienteDesde.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169387)

    End Select

    Exit Sub

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


Public Function ParcelaRec_Le_MenorData(dtData As Date, lCodCliente As Long) As Long
'obtem a data de vencimento mais antiga de uma parcela a receber em aberto de um cliente

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ParcelaRec_Le_MenorData

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87590

    'Le Último titulo mais antigo em aberto
    lErro = Comando_Executar(lComando, "SELECT MIN(ParcelasRec.DataVencimentoReal) FROM ParcelasRec, TitulosRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.Status = ? AND TitulosRec.Cliente = ?" _
        , dtData, STATUS_ABERTO, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 87591

    'Busca o primeiro titulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87592

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ParcelaRec_Le_MenorData = SUCESSO

    Exit Function

Erro_ParcelaRec_Le_MenorData:

    ParcelaRec_Le_MenorData = gErr

    Select Case gErr

        Case 87590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 87591, 87592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169388)

    End Select

    'Fecha Comando --> saída por erro
    Call Comando_Fechar(lComando)

    Exit Function

End Function


