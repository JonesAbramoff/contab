VERSION 5.00
Begin VB.UserControl RelOpBordRetPagtoOcx 
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   2955
   ScaleWidth      =   5220
   Begin VB.CheckBox SoErros 
      Caption         =   "Exibir somente os erros"
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
      Left            =   135
      TabIndex        =   2
      Top             =   2220
      Width           =   2685
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
      Left            =   3180
      Picture         =   "RelOpBordRetPagtoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1995
      Width           =   1815
   End
   Begin VB.Label NumIntDoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Seq 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3900
      TabIndex        =   9
      Top             =   1380
      Width           =   1140
   End
   Begin VB.Label Label5 
      Caption         =   "Sequencial:"
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
      Left            =   2850
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   8
      Top             =   1410
      Width           =   1005
   End
   Begin VB.Label Data 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1380
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "Data:"
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
      Left            =   1140
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1410
      Width           =   510
   End
   Begin VB.Label Empresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa:"
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
      Left            =   840
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   870
      Width           =   840
   End
   Begin VB.Label NomeArquivo 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   300
      Width           =   3360
   End
   Begin VB.Label LabelArq 
      Caption         =   "Nome do Arquivo:"
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
      Left            =   105
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   345
      Width           =   1560
   End
End
Attribute VB_Name = "RelOpBordRetPagtoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoArq As AdmEvento
Attribute objEventoArq.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 205518
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 205518
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205519)

    End Select

    Exit Function

End Function

Private Sub Form_Load()
    
    Set objEventoArq = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 205520

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 205520

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205521)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long, sSelecao As String
Dim iSoErros As Integer

On Error GoTo Erro_PreencherRelOp
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 205522
    
    If Len(Trim(NomeArquivo.Caption)) = 0 Then gError 205523
                
    sSelecao = "NumIntDoc = " & Forprint_ConvLong(StrParaLong(NumIntDoc.Caption))
    
    If SoErros.Value = vbChecked Then
        iSoErros = MARCADO
        sSelecao = sSelecao & " E Erro = " & Forprint_ConvInt(iSoErros)
    Else
        iSoErros = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NSOERROS", CStr(iSoErros))
    If lErro <> AD_BOOL_TRUE Then gError 205524
      
    objRelOpcoes.sSelecao = sSelecao
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 205522, 205524
        
        Case 205523
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEARQ_NAO_PREENCHIDO1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205525)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoArq = Nothing

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_RELOP_RETCOBERR
    Set Form_Load_Ocx = Me
    Caption = "Retorno de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBordRetPagto"
    
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
Private Sub LabelArq_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelArq, Source, X, Y)
End Sub

Private Sub LabelArq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelArq, Button, Shift, X, Y)
End Sub

Private Sub objEventoArq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objArq As ClassRetPagto
Dim iIndice As Integer

On Error GoTo Erro_objEventoArq_evSelecao

    Set objArq = obj1

    NomeArquivo.Caption = objArq.sNomeArq
    Empresa.Caption = objArq.sNomeEmpresa
    Data.Caption = Format(objArq.dtDataGeracao, "dd/mm/yyyy")
    Seq.Caption = objArq.lSeqArquivo
    NumIntDoc.Caption = objArq.lNumIntDoc

    Me.Show

    Exit Sub

Erro_objEventoArq_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205526)

    End Select

    Exit Sub

End Sub

Public Sub LabelArq_Click()

Dim objArq As New ClassRetPagto
Dim colSelecao As Collection

On Error GoTo Erro_LabelArq_Click

    objArq.sNomeArq = NomeArquivo.Caption
    
    'Chama a Tela de browse
    Call Chama_Tela("RetPagtoLista", colSelecao, objArq, objEventoArq)

    Exit Sub

Erro_LabelArq_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205527)

    End Select

    Exit Sub
    
End Sub

