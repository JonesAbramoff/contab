VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTitRecTelOcx 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6390
   Begin VB.ListBox Clientes 
      Height          =   2535
      Left            =   225
      Style           =   1  'Checkbox
      TabIndex        =   17
      Top             =   1620
      Width           =   5910
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      Picture         =   "RelOpTitRecTelOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4245
      Width           =   1800
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2205
      Picture         =   "RelOpTitRecTelOcx.ctx":101A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4245
      Width           =   1800
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
      Left            =   4260
      Picture         =   "RelOpTitRecTelOcx.ctx":21FC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitRecTelOcx.ctx":22FE
      Left            =   960
      List            =   "RelOpTitRecTelOcx.ctx":2300
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitRecTelOcx.ctx":2302
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitRecTelOcx.ctx":2480
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitRecTelOcx.ctx":29B2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitRecTelOcx.ctx":2B3C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1230
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   930
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2295
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   930
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox NumDiasInicial 
      Height          =   300
      Left            =   795
      TabIndex        =   1
      Top             =   930
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumDiasFinal 
      Height          =   300
      Left            =   1845
      TabIndex        =   2
      Top             =   930
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "e"
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
      Left            =   1590
      TabIndex        =   14
      Top             =   990
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entre "
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
      Left            =   240
      TabIndex        =   13
      Top             =   990
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "dias de atraso"
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
      Left            =   2625
      TabIndex        =   12
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Op��o:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTitRecTelOcx"
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

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 57319

    'Pega parametros e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMINI", sParam)
    If lErro <> SUCESSO Then Error 57320
    
    NumDiasInicial.Text = CStr(sParam)
    
    lErro = objRelOpcoes.ObterParametro("NNUMFIM", sParam)
    If lErro <> SUCESSO Then Error 57321
    
    NumDiasFinal.Text = CStr(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 57319, 57320, 57321

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173533)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim sDevedores As String
'###########################
'Inserido por Wagner
Dim lNumIntRel As Long
Dim colCliente As New Collection
Dim dtDataIni As Date
Dim dtDataFim As Date
'###########################

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 57322
    
    'Verificar se N�mero de dias inicial foi preenchido
    If Len(Trim(NumDiasInicial.Text)) = 0 Then Error 57323
    
    'Verificar se N�mero de dias Final foi preenchido
    If Len(Trim(NumDiasFinal.Text)) = 0 Then Error 57324
    
    'Verificar se N�mero de dias finais � maior
    If CInt(NumDiasFinal.Text) < CInt(NumDiasInicial.Text) Then Error 57325
    
    'Pegar parametro data da tela
    lErro = objRelOpcoes.IncluirParametro("NNUMINI", NumDiasInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 57326

    lErro = objRelOpcoes.IncluirParametro("NNUMFIM", NumDiasFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 57327

    '##################################
    'Inserido por Wagner
    lErro = Move_Cliente_Memoria(colCliente)
    If lErro <> SUCESSO Then gError 132090
        
    dtDataIni = DateAdd("d", StrParaInt(NumDiasFinal.Text) * -1, gdtDataAtual)
    dtDataFim = DateAdd("d", StrParaInt(NumDiasInicial.Text) * -1, gdtDataAtual)

    lErro = CF("RelClienteEmAtraso_Prepara", giFilialEmpresa, lNumIntRel, colCliente)
    If lErro <> SUCESSO Then gError 132091

    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(dtDataFim))
    If lErro <> AD_BOOL_TRUE Then gError 132092
    
    lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(dtDataIni))
    If lErro <> AD_BOOL_TRUE Then gError 132093
    
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then gError 132094
    '##################################

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 57322
        
        Case 57323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO", Err, Error$)
            NumDiasInicial.SetFocus
            
        Case 57324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO", Err, Error$)
            NumDiasFinal.SetFocus
        
        Case 57326, 57327
            
        Case 57325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err, Error$)
            NumDiasInicial.SetFocus
            
        Case 132090 To 132094 'Inserido por wagner
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173534)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 57328
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche combo com as op��es de relat�rio
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 57329

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 57329
        
        Case 57328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173535)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 57330

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPTITRECMALA")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 57331

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as op��es da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 57330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 57331

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173536)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 57332

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 57332

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173537)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de op��o"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 57333

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 57334

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 57335

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 57698

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 57333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 57334, 57335, 57698

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173538)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

   ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim colCodigoDescricao As New AdmCollCodigoNome
Dim lErro As Long, iIndice As Integer
Dim objCodigoDescricao As AdmlCodigoNome

On Error GoTo Erro_OpcoesRel_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173539)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub NumDiasFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumDiasFinal)

End Sub

Private Sub NumDiasInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumDiasInicial)

End Sub

Private Sub UpDown1_DownClick()

Dim iIndice As Integer

    If Len(Trim(NumDiasInicial.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumDiasInicial.Text)
    
    If iIndice = 0 Then Exit Sub
    
    NumDiasInicial.PromptInclude = False
    NumDiasInicial.Text = CStr(iIndice - 1)
    NumDiasInicial.PromptInclude = True
    NumDiasInicial.SetFocus

End Sub

Private Sub UpDown2_UpClick()

Dim iIndice As Integer

    If Len(Trim(NumDiasFinal.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumDiasFinal.Text)
    
    NumDiasFinal.PromptInclude = False
    NumDiasFinal.Text = CStr(iIndice + 1)
    NumDiasFinal.PromptInclude = True
    NumDiasFinal.SetFocus


End Sub
Private Sub UpDown2_DownClick()

Dim iIndice As Integer

    If Len(Trim(NumDiasFinal.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumDiasFinal.Text)
    
    If iIndice = 0 Then Exit Sub
    
    NumDiasFinal.PromptInclude = False
    NumDiasFinal.Text = CStr(iIndice - 1)
    NumDiasFinal.PromptInclude = True
    NumDiasFinal.SetFocus

End Sub

Private Sub UpDown1_UpClick()

Dim iIndice As Integer

    If Len(Trim(NumDiasInicial.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumDiasInicial.Text)
    
    NumDiasInicial.PromptInclude = False
    NumDiasInicial.Text = CStr(iIndice + 1)
    NumDiasInicial.PromptInclude = True
    NumDiasInicial.SetFocus

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TIT_REC_TEL
    Set Form_Load_Ocx = Me
    Caption = "T�tulos para cobran�a por telefone"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitRecTel"
    
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



Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

'#########################################################
'Inserido por Wagner
Private Function Carrega_ComboClientesAtrasados(ByVal lNumDiasIni As Long, ByVal lNumDiasFim As Long) As Long

Dim lErro As Long
Dim objCliente As ClassCliente
Dim colCliente As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_ComboClientesAtrasados
    
    lErro = CF("ClientesAtrasados_Le", colCliente, lNumDiasIni, lNumDiasFim, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 132095

    Clientes.Clear

    For Each objCliente In colCliente
        Clientes.AddItem objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        Clientes.Selected(iIndice) = True
        iIndice = iIndice + 1
    Next

    Carrega_ComboClientesAtrasados = SUCESSO

    Exit Function

Erro_Carrega_ComboClientesAtrasados:

    Carrega_ComboClientesAtrasados = gErr

    Select Case gErr
    
        Case 132095

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173540)

    End Select

    Exit Function

End Function

Private Function Move_Cliente_Memoria(ByVal colCliente As Collection) As Long

Dim lErro As Long
Dim objCliente As ClassCliente
Dim iIndice As Integer

On Error GoTo Erro_Move_Cliente_Memoria

    'Verificar se teve algum item marcado
    For iIndice = 0 To Clientes.ListCount - 1
    
        If Clientes.Selected(iIndice) = True Then
        
            Set objCliente = New ClassCliente
            
            objCliente.lCodigo = Codigo_Extrai(Clientes.List(iIndice))
            
            colCliente.Add objCliente
            
        End If
    
    Next

    Move_Cliente_Memoria = SUCESSO

    Exit Function

Erro_Move_Cliente_Memoria:

    Move_Cliente_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173541)

    End Select

    Exit Function
    
End Function

Private Sub NumDiasInicial_Change()

    If IsNumeric(NumDiasInicial.Text) And IsNumeric(NumDiasFinal.Text) Then
        Call Carrega_ComboClientesAtrasados(StrParaLong(NumDiasInicial.Text), StrParaLong(NumDiasFinal.Text))
    End If

End Sub

Private Sub NumDiasFinal_Change()

    If IsNumeric(NumDiasInicial.Text) And IsNumeric(NumDiasFinal.Text) Then
        Call Carrega_ComboClientesAtrasados(StrParaLong(NumDiasInicial.Text), StrParaLong(NumDiasFinal.Text))
    End If

End Sub

Private Sub BotaoMarcarTodos_Click()

    Call MarcaDesmarca(True)
    
End Sub

Private Sub BotaoDesmarcarTodos_Click()

    Call MarcaDesmarca(False)

End Sub

Private Sub MarcaDesmarca(ByVal bFlag As Boolean)

Dim iIndice As Integer

    For iIndice = 0 To Clientes.ListCount - 1
    
        Clientes.Selected(iIndice) = bFlag
        
    Next

End Sub
'#########################################################

