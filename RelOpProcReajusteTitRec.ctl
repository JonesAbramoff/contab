VERSION 5.00
Begin VB.UserControl RelOpProcReajTitRecOcx 
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ScaleHeight     =   3765
   ScaleWidth      =   5175
   Begin VB.CommandButton BotaoProcessamentos 
      Caption         =   "Identificar Processamento..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   225
      TabIndex        =   17
      Top             =   150
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Relatório por:"
      Height          =   975
      Left            =   135
      TabIndex        =   9
      Top             =   2670
      Width           =   4890
      Begin VB.OptionButton OptIndices 
         Caption         =   "Índices"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2610
         TabIndex        =   13
         Top             =   615
         Width           =   2025
      End
      Begin VB.OptionButton OptParcela 
         Caption         =   "Parcela a Receber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   615
         Width           =   2025
      End
      Begin VB.OptionButton OptTitulo 
         Caption         =   "Título a Receber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2610
         TabIndex        =   11
         Top             =   225
         Width           =   2025
      End
      Begin VB.OptionButton OptCcl 
         Caption         =   "Centro de Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
      Height          =   1440
      Left            =   135
      TabIndex        =   2
      Top             =   1155
      Width           =   4890
      Begin VB.Label LabelAtualizado 
         Height          =   240
         Left            =   1605
         TabIndex        =   15
         Top             =   1095
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Atualizado até:"
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
         Left            =   255
         TabIndex        =   14
         Top             =   1095
         Width           =   1665
      End
      Begin VB.Label LabelHora 
         Height          =   240
         Left            =   3300
         TabIndex        =   8
         Top             =   330
         Width           =   975
      End
      Begin VB.Label LabelUsuario 
         Height          =   240
         Left            =   1590
         TabIndex        =   7
         Top             =   705
         Width           =   2820
      End
      Begin VB.Label LabelData 
         Height          =   240
         Left            =   1575
         TabIndex        =   6
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Left            =   2730
         TabIndex        =   5
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   795
         TabIndex        =   4
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   330
         Width           =   480
      End
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
      Height          =   615
      Left            =   3480
      Picture         =   "RelOpProcReajusteTitRec.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   1575
   End
   Begin VB.Label Sequencial 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1680
      TabIndex        =   16
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label LabelSequencial 
      AutoSize        =   -1  'True
      Caption         =   "Seqüencial:"
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
      Left            =   600
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   765
      Width           =   1020
   End
End
Attribute VB_Name = "RelOpProcReajTitRecOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoSequencial As AdmEvento
Attribute objEventoSequencial.VB_VarHelpID = -1

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

   If Not (gobjRelatorio Is Nothing) Then gError 138530
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 138530
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171551)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoSequencial = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171552)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
   
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 138531
    
    If OptCcl.Value = True Then
        gobjRelatorio.sNomeTsk = "PRREACCL"
    End If

    If OptTitulo.Value = True Then
        gobjRelatorio.sNomeTsk = "PRREATIT"
    End If

    If OptParcela.Value = True Then
        gobjRelatorio.sNomeTsk = "PRREAPAR"
    End If

    If OptIndices.Value = True Then
        gobjRelatorio.sNomeTsk = "PRREAIND"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 138531

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171553)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
    
    If Len(Trim(Sequencial.Caption)) = 0 Then gError 138532
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 138533
         
    lErro = objRelOpcoes.IncluirParametro("NNUMINTDOC", Sequencial.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 138534
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 138532
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)
            
        Case 138533, 138534
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171554)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoSequencial = Nothing
 
 End Sub

Private Sub BotaoProcessamentos_Click()
    Call LabelSequencial_Click
End Sub

Private Sub LabelSequencial_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProcReajTitRec As New ClassProcReajTitRec

On Error GoTo Erro_LabelSequencial_Click
    
    If Len(Trim(Sequencial.Caption)) > 0 Then
        objProcReajTitRec.lNumIntDoc = StrParaLong(Sequencial.Caption)
    End If
    
    'Chama Tela BorderoCobrancaLista
    Call Chama_Tela("ProcReajusteTitRecLista", colSelecao, objProcReajTitRec, objEventoSequencial)

    Exit Sub

Erro_LabelSequencial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171555)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoSequencial_evSelecao(obj1 As Object)

Dim objProcReajTitRec As ClassProcReajTitRec
Dim lErro As Long

On Error GoTo Erro_objEventoSequencial_evSelecao

    Set objProcReajTitRec = obj1
    
    Sequencial.Caption = CStr(objProcReajTitRec.lNumIntDoc)
    LabelData.Caption = Format(objProcReajTitRec.dtDataProc, "dd/mm/yyyy")
    LabelAtualizado.Caption = Format(objProcReajTitRec.dtAtualizadoAte, "mm/yyyy")
    LabelUsuario.Caption = objProcReajTitRec.sUsuario
    LabelHora.Caption = Format(objProcReajTitRec.dHoraProc, "HH:MM:SS")

    Me.Show

    Exit Sub

Erro_objEventoSequencial_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171556)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BORDERO_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Informações sobre Reajuste de Títulos a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProcReajusteTitRec"
    
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
            
    End If

End Sub

Private Sub Labeldata_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub Labeldata_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub LabelHora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHora, Source, X, Y)
End Sub

Private Sub LabelHora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHora, Button, Shift, X, Y)
End Sub

Private Sub LabelAtualizado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtualizado, Source, X, Y)
End Sub

Private Sub LabelAtualizado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtualizado, Button, Shift, X, Y)
End Sub

Private Sub LabelUsuario_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelUsuario, Source, X, Y)
End Sub

Private Sub LabelUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelUsuario, Button, Shift, X, Y)
End Sub
