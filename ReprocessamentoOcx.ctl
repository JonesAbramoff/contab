VERSION 5.00
Begin VB.UserControl ReprocessamentoOcx 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   LockControls    =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   6120
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   3135
      ScaleHeight     =   765
      ScaleWidth      =   2670
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   195
      Width           =   2730
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   2190
         Picture         =   "ReprocessamentoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoReprocessar 
         Height          =   600
         Left            =   90
         Picture         =   "ReprocessamentoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   1995
      End
   End
   Begin VB.ComboBox Exercicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1590
   End
   Begin VB.ComboBox Periodo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      Top             =   225
      Width           =   855
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exercicio:"
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   270
      Top             =   900
      Width           =   735
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Periodo:"
   End
End
Attribute VB_Name = "ReprocessamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel: Mario
'Revisado em 20/8/98

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoReprocessar_Click()
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim lErro As Long
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoReprocessar_Click

    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 20362

    lErro = CF("Rotina_Reprocessamento",sNomeArqParam, giFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 11781
    
    Unload Me
    
    Exit Sub

Erro_BotaoReprocessar_Click:

    Select Case Err
    
        Case 11781, 20362
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173761)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_Exercicio_Click

    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
    lErro = Preenche_ComboPeriodo(iExercicio)
    If lErro <> SUCESSO Then Error 11782
    
    Periodo.ListIndex = 0
    
    Exit Sub
    
Erro_Exercicio_Click:

    Select Case Err
    
        Case 11782
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173762)
        
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_Form_Load

    lErro = Preenche_ComboExercicio
    If lErro <> SUCESSO Then Error 11783
    
    If Exercicio.ListCount = 0 Then Error 11784
    
    Exercicio.ListIndex = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case Err
    
        Case 11783
        
        Case 11784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TODOS_EXERCICIOS_FECHADOS", Err, Error$)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173763)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros() As Long
 
    Trata_Parametros = SUCESSO
 
    Exit Function
 
End Function

Function Preenche_ComboExercicio() As Long
'preenche Combo de Exercicios com os Exercicio que nao estao fechados

Dim colExercicios As New Collection
Dim lErro As Long
Dim iConta As Integer
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Preenche_ComboExercicio

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos",colExercicios)
    If lErro <> SUCESSO Then Error 11785

    'preenche ComboBox com NomeExterno e ItemData com Exercicio
    For iConta = 1 To colExercicios.Count

        Set objExercicio = colExercicios.Item(iConta)
        If objExercicio.iStatus <> EXERCICIO_FECHADO Then
            Exercicio.AddItem objExercicio.sNomeExterno
            Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
        End If
    Next

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = Err

    Select Case Err

        Case 11785

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173764)

    End Select

    Exit Function

End Function
Private Function Preenche_ComboPeriodo(iExercicio As Integer)

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_Preenche_ComboPeriodo

    lErro = CF("Periodo_Le_Todos_Exercicio",giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 11786

    Periodo.Clear

    For Each objPeriodo In colPeriodos

        Periodo.AddItem objPeriodo.sNomeExterno
        Periodo.ItemData(Periodo.NewIndex) = objPeriodo.iPeriodo

    Next

    Exit Function

Erro_Preenche_ComboPeriodo:

    Select Case Err

        Case 11786

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173765)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REPROCESSAMENTO
    Set Form_Load_Ocx = Me
    Caption = "Reprocessamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Reprocessamento"
    
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
   ' Parent.UnloadDoFilho
    
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


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

