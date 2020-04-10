VERSION 5.00
Begin VB.UserControl CustoMedioProducao 
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   LockControls    =   -1  'True
   ScaleHeight     =   2685
   ScaleWidth      =   6090
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3825
      ScaleHeight     =   690
      ScaleWidth      =   1980
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   2040
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1470
         Picture         =   "CustoMedioProducao.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   105
         Picture         =   "CustoMedioProducao.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "CustoMedioProducao.ctx":1A40
      Top             =   1140
      Width           =   5715
   End
   Begin VB.Label Label3 
      Caption         =   "Mês:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1620
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Mes 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Ano 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   750
      TabIndex        =   4
      Top             =   450
      Width           =   555
   End
End
Attribute VB_Name = "CustoMedioProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'Dim sMes As String
'Dim objEstoqueMes As New ClassEstoqueMes
'
'On Error GoTo Erro_CustoMedioProducao_Form_Load
'
'    'Tenta ler EstoqueMes com custo prod não apurado
'    lErro = CF("EstoqueMesNaoApurado_Le",objEstoqueMes)
'    If lErro = 25221 Then Error 25215  'não encontrou
'    If lErro <> SUCESSO Then Error 25216
'
'    'Converte mês para extenso
'    lErro = MesNome(objEstoqueMes.iMes, sMes)
'    If lErro <> SUCESSO Then Error 25217
'
'    'Coloca mês, ano na Tela
'    Ano.Caption = CStr(objEstoqueMes.iAno)
'    Mes.Caption = sMes
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_CustoMedioProducao_Form_Load:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case 25215
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_PRODUCAO_APURADO", Err)
'
'        Case 25216, 25217 'tratado na rotina chamada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158658)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Function Trata_Parametros() As Long
'
'    Trata_Parametros = SUCESSO
'
'End Function
'
'
'Private Sub BotaoApurar_Click()
'
'Dim lErro As Long
'Dim sNomeArqParam As String
'Dim iMes As Integer
'Dim iAno As Integer
'Dim objEstoqueMes As New ClassEstoqueMes
'
'On Error GoTo Erro_BotaoApurar_Click
'
'    'Pega ano e mês da Tela
'    iAno = CInt(Ano.Caption)
'    lErro = MesNumero(Mes.Caption, iMes)
'    If lErro <> SUCESSO Then Error 25213
'
'    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
'    objEstoqueMes.iAno = iAno
'    objEstoqueMes.iMes = iMes
'
'    'Lê EstoqueMes correspondente a este mês e ano
'    lErro = CF("EstoqueMes_Le",objEstoqueMes)
'    If lErro = 36513 Then Error 25225  'não encontrou
'    If lErro <> SUCESSO Then Error 25226
'
'    'Verifica se o mês está fechado
'    If objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_ABERTO Then Error 25224
'
'    'Prepara para chamar rotina batch
'    lErro = Sistema_Preparar_Batch(sNomeArqParam)
'    If lErro <> SUCESSO Then Error 25212
'
'    'Chama rotina batch que calcula custo médio de produção
'    'e valoriza movimentos de materiais produzidos
'    lErro = CF("Rotina_CustoMedioProducao_Calcula",sNomeArqParam, giFilialEmpresa, iAno, iMes)
'    If lErro <> SUCESSO Then Error 25214
'
'    Unload Me
'
'    Exit Sub
'
'Erro_BotaoApurar_Click:
'
'    Select Case Err
'
'        Case 25212, 25213, 25214, 25226
'
'        Case 25224
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_ABERTO", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
'
'        Case 25225
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
'
'         Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158659)
'
'    End Select
'
'   Exit Sub
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
''**** inicio do trecho a ser copiado *****
'
'Public Function Form_Load_Ocx() As Object
'
'    Parent.HelpContextID = IDH_CUSTO_MEDIO_PRODUCAO
'    Set Form_Load_Ocx = Me
'    Caption = "Custo Médio de Produção"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "CustoMedioProducao"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''**** fim do trecho a ser copiado *****
'
'
'
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
'Private Sub Mes_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Mes, Source, X, Y)
'End Sub
'
'Private Sub Mes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Mes, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
'Private Sub Ano_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Ano, Source, X, Y)
'End Sub
'
'Private Sub Ano_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Ano, Button, Shift, X, Y)
'End Sub
'
