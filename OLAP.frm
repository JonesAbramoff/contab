VERSION 5.00
Begin VB.Form OLAP 
   Caption         =   "OLAP"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoProcessar 
      Caption         =   "Processa os Cubos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3360
      Picture         =   "OLAP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Processar o cubo"
      Top             =   1470
      Width           =   1095
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3360
      Picture         =   "OLAP.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2535
      Width           =   1095
   End
   Begin VB.TextBox Servidor 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox BD 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox Cubos 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   1485
      Width           =   2730
   End
   Begin VB.Label Label3 
      Caption         =   "Cubos:"
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
      TabIndex        =   3
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Base de Dados:"
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
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Servidor:"
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
      Left            =   1305
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "OLAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''''' API do WIndows
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'
'Private objdsoServer As DSO.Server
'Private objdsoDB As DSO.MDStore
'Private objdsoCube As DSO.MDStore
'
'Private Sub BotaoCancela_Click()
'
'    Unload Me
'
'End Sub
'
'Function Trata_Parametros()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Trata_Parametros
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163631)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
'Public Function Processa_DB_OLAP() As Long
'
'On Error GoTo Erro_Processa_DB_OLAP
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    objdsoDB.Process processFull
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Processa_DB_OLAP = SUCESSO
'
'    Exit Function
'
'Erro_Processa_DB_OLAP:
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    Processa_DB_OLAP = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163632)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoProcessar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoProcessar_Click
'
'    If Hour(Now) < PROC_OLAP_HORARIO_INVALIDO_FIM And Hour(Now) >= PROC_OLAP_HORARIO_INVALIDO_INICIO Then gError 189562
'
'    lErro = Processa_DB_OLAP
'    If lErro <> SUCESSO Then gError 140850
'
'    Call Rotina_Aviso(vbOKOnly, "AVISO_CUBOS_PROCESSADOS")
'
'    Exit Sub
'
'Erro_BotaoProcessar_Click:
'
'    Select Case gErr
'
'        Case 189562
'            Call Rotina_Erro(vbOKOnly, "PROC_OLAP_HORARIO_INVALIDO", gErr, PROC_OLAP_HORARIO_INVALIDO_INICIO, PROC_OLAP_HORARIO_INVALIDO_FIM)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163633)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Form_Load()
'
'Dim sBD As String
'Dim sServidor As String
'Dim iIndice As Integer
'Dim objdsoCubo As DSO.Cube
'
'On Error GoTo Erro_Form_Load
'
'    sBD = String(128, 0)
'    sServidor = String(128, 0)
'
'    Call GetPrivateProfileString("Forprint", "OLAP", "Demo", sBD, 128, "ADM100.INI")
'    Call GetPrivateProfileString("Forprint", "Servidor", "ASP16", sServidor, 128, "ADM100.INI")
'
'    sBD = Replace(sBD, Chr(0), "")
'    sServidor = Replace(sServidor, Chr(0), "")
'
'    BD.Text = sBD
'    Servidor.Text = sServidor
'
'    Set objdsoServer = New DSO.Server
'
'    objdsoServer.Connect (Servidor.Text)
'
'    Set objdsoDB = objdsoServer.MDStores(BD.Text)
'
'    Cubos.Clear
'
'    For iIndice = 1 To objdsoDB.MDStores.Count
'
'        Set objdsoCubo = objdsoDB.MDStores.Item(iIndice)
'
'        Cubos.AddItem objdsoCubo.Name
'
'    Next
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163634)
'
'    End Select
'
'    Exit Sub
'
'End Sub
