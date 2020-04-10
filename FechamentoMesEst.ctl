VERSION 5.00
Begin VB.UserControl FechamentoMesEst 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   LockControls    =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   6195
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3960
      ScaleHeight     =   690
      ScaleWidth      =   1980
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2040
      Begin VB.CommandButton BotaoExecutar 
         Height          =   510
         Left            =   105
         Picture         =   "FechamentoMesEst.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1470
         Picture         =   "FechamentoMesEst.ctx":1C42
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   375
      Width           =   405
   End
   Begin VB.Label Ano 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   645
      TabIndex        =   5
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Mês que será aberto:"
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
      Left            =   165
      TabIndex        =   4
      Top             =   870
      Width           =   1995
   End
   Begin VB.Label NomeMes 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2175
      TabIndex        =   3
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "FechamentoMesEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Tem que transferir os custos(custo médio e custo standard - alterar o status do mes que está fechando para fechado)  de SldMesEst de um mes para o outro e quando ocorrer um final de
'exercicio, tem que criar os registros em SldMesEst e transferir tambem o saldo inicial.
'O saldo inicial do proximo ano vem do computo do saldo inicial do ano anterior somado com as entradas de cada mes e subtraido das saidas de cada mes
'O correlato para o valor inicial
'O CMPInicial vem do campo CustoMedioProducao12
'Os custos médio e standard vem do mes 12
'Atualizar a tabela EstoqueMes.

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim iMes As Integer
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_BotaoExecutar_Click
    
    'joga o nome do mes e recebe o numero do mes respectivo
    lErro = MesNumero(NomeMes.Caption, iMes)
    If lErro <> SUCESSO Then Error 40697
    
    'preenche o objeto que é passado como parametro
    objEstoqueMes.iMes = iMes
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    objEstoqueMes.iAno = Ano.Caption
    
    If objEstoqueMes.iMes = 1 Then
        objEstoqueMes.iMes = 12
        objEstoqueMes.iAno = objEstoqueMes.iAno - 1
    Else
        objEstoqueMes.iMes = objEstoqueMes.iMes - 1
    End If
    
    'chama a tela FechamentoMesEst1
    Call Chama_Tela_Modal("FechamentoMesEst1", objEstoqueMes)
    
    Unload Me
            
    Exit Sub
    
Erro_BotaoExecutar_Click:

    Select Case Err
        
        Case 40697
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160188)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Iniciar() As Long

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim sMes As String

On Error GoTo Erro_Iniciar
    
    'preenche o objeto
    
    If giFilialEmpresa = EMPRESA_TODA Then
        objEstoqueMes.iFilialEmpresa = 1
    Else
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa
    End If
    
    objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_ABERTO

    'Ler o mês e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
    lErro = CF("EstoqueMes_Le_Mes1", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 60861 Then Error 40667

    If lErro = 60861 Then Error 40668

    If objEstoqueMes.iMes = 12 Then
        objEstoqueMes.iMes = 1
        objEstoqueMes.iAno = objEstoqueMes.iAno + 1
    Else
        objEstoqueMes.iMes = objEstoqueMes.iMes + 1
    End If
    
    'formata o mes de numero para seu nome respectivo
    lErro = MesNome(objEstoqueMes.iMes, sMes)
    If lErro <> SUCESSO Then Error 40669

    'joga o ano e o mes que estiverem aberto
    Ano.Caption = CStr(objEstoqueMes.iAno)
    NomeMes.Caption = sMes

    Iniciar = SUCESSO
    
    Exit Function
    
Erro_Iniciar:

    Iniciar = Err

    Select Case Err
    
        Case 40667, 40669

        Case 40668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAOEXISTE_MES_ABERTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160189)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim sMes As String

On Error GoTo Erro_Form_Load
    
    lErro = Iniciar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160189)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FECHAMENTO_MES_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Abertura de Mês - Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FechamentoMesEst"
    
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

'**** fim do trecho a ser copiado *****

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Ano_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Ano, Source, X, Y)
End Sub

Private Sub Ano_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Ano, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub NomeMes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeMes, Source, X, Y)
End Sub

Private Sub NomeMes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeMes, Button, Shift, X, Y)
End Sub
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
