VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl RelOpBordRetCobErrOcx 
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   1620
   ScaleWidth      =   5220
   Begin VB.CheckBox CheckExibirCustasTarifas 
      Caption         =   "Exibir custas e tarifas"
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
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox CheckExibirBaixasLiq 
      Caption         =   "Exibir baixas por liquidação"
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
      TabIndex        =   4
      Top             =   810
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4650
      TabIndex        =   2
      Top             =   270
      Width           =   360
   End
   Begin VB.TextBox NomeArquivo 
      Height          =   315
      Left            =   1695
      TabIndex        =   1
      Top             =   285
      Width           =   2955
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
      Picture         =   "RelOpBordRetCobErrOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   900
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   3
      Top             =   345
      Width           =   1560
   End
End
Attribute VB_Name = "RelOpBordRetCobErrOcx"
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

Private Sub BotaoProcurar_Click()

    On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|Ret Files (*.ret)|*.ret"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen

    ' Display name of selected file

    NomeArquivo.Text = CommonDialog1.FileName
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 48587
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 48587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167389)

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
    If lErro <> SUCESSO Then Error 48602

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 48602

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167390)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long, sSelecao As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 48604
             
    'Preenche o nome do arquivo
    lErro = objRelOpcoes.IncluirParametro("TNOMEARQ", NomeArquivo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 48605
    
    If CheckExibirBaixasLiq.Value = False Then sSelecao = "ExcluirBaixasLiq"
        
    If CheckExibirCustasTarifas = False Then
        If sSelecao <> "" Then
            sSelecao = sSelecao & " E "
        End If
        sSelecao = sSelecao & "ExcluirCustasTarif"
    End If
    
    If sSelecao <> "" Then
        objRelOpcoes.sSelecao = sSelecao
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 48604, 48605
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167391)

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
    Caption = "Críticas de Arquivo de Retorno da Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBordRetCobErr"
    
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
Dim objArq As ClassBorderoRetCobr
Dim iIndice As Integer

On Error GoTo Erro_objEventoArq_evSelecao

    Set objArq = obj1

    NomeArquivo.Text = objArq.sNomeArq

    Me.Show

    Exit Sub

Erro_objEventoArq_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202129)

    End Select

    Exit Sub

End Sub

Public Sub LabelArq_Click()

Dim objArq As New ClassBorderoRetCobr
Dim colSelecao As Collection

On Error GoTo Erro_LabelArq_Click

    objArq.sNomeArq = NomeArquivo.Text
    
    'Chama a Tela de browse
    Call Chama_Tela("BordCobrRetArqsLista", colSelecao, objArq, objEventoArq)

    Exit Sub

Erro_LabelArq_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202131)

    End Select

    Exit Sub
    
End Sub

