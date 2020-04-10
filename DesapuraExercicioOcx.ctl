VERSION 5.00
Begin VB.UserControl DesapuraExercicioOcx 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ScaleHeight     =   1170
   ScaleWidth      =   5760
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "DesapuraExercicioOcx.ctx":0000
      Left            =   1005
      List            =   "DesapuraExercicioOcx.ctx":0002
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3600
      ScaleHeight     =   690
      ScaleWidth      =   1995
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   2055
      Begin VB.CommandButton BotaoDesapurar 
         Caption         =   "Desapurar"
         Height          =   510
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1500
         Picture         =   "DesapuraExercicioOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Left            =   330
      TabIndex        =   6
      Top             =   750
      Width           =   615
   End
   Begin VB.Label Status 
      Height          =   255
      Left            =   1035
      TabIndex        =   5
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Exercício:"
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
      Left            =   90
      TabIndex        =   4
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "DesapuraExercicioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim iExercicio As Integer
Dim iPosInicial As Integer
Dim iLote As Integer
Dim objCTBConfig As New ClassCTBConfig
Dim sContaEnxuta As String

On Error GoTo Erro_ApuraExercicio_Form_Load

    lErro = Preenche_ComboExercicio(iPosInicial)
    If lErro <> SUCESSO Then gError 188389

    If Exercicio.ListCount = 0 Then gError 188390

    Exercicio.ListIndex = 0

    iExercicio = CInt(Exercicio.ItemData(0))

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_ApuraExercicio_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 188389
        
        Case 188390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOS_FECHADOS", gErr, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188391)

    End Select

    Exit Sub

End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iLote As Integer
Dim objExerciciosFilial As New ClassExerciciosFilial

On Error GoTo Erro_Exercicio_Click

    objExerciciosFilial.iFilialEmpresa = giFilialEmpresa
    objExerciciosFilial.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
    lErro = CF("ExerciciosFilial_Le", objExerciciosFilial)
    If lErro <> SUCESSO And lErro <> 20389 Then gError 188400

    If lErro = 20389 Then Error 188401

    If objExerciciosFilial.iStatus = EXERCICIO_ABERTO Then
        Status.Caption = EXERCICIO_DESC_ABERTO
    ElseIf objExerciciosFilial.iStatus = EXERCICIO_APURADO Then
        Status.Caption = EXERCICIO_DESC_APURADO
    ElseIf objExerciciosFilial.iStatus = EXERCICIO_FECHADO Then
        Status.Caption = EXERCICIO_DESC_FECHADO
    End If
    
    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 188400

        Case 188401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOSFILIAL_INEXISTENTE", Err, objExerciciosFilial.iExercicio, objExerciciosFilial.iFilialEmpresa)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143116)

    End Select

    Exit Sub
    
End Sub

Function Preenche_ComboExercicio(iPosInicial As Integer) As Long
'preenche Combo de Exercicios

Dim colExercicios As New Collection
Dim lErro As Long
Dim iConta As Integer
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Preenche_ComboExercicio

    iPosInicial = 1

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then gError 188398

    'preenche ComboBox com NomeExterno e ItemData com Exercicio
    For iConta = 1 To colExercicios.Count

        Set objExercicio = colExercicios.Item(iConta)
        If objExercicio.iStatus <> EXERCICIO_FECHADO Then
            Exercicio.AddItem objExercicio.sNomeExterno
            Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
            If objExercicio.iExercicio = giExercicioAtual Then
                iPosInicial = Exercicio.NewIndex
            End If
        End If
    Next

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = gErr

    Select Case gErr

        Case 188398

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188399)

    End Select

    Exit Function

End Function

Private Sub BotaoDesapurar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoDesapurar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Exercicio.Text) = 0 Then gError 188395

    lErro = Desapura_Exercicio()
    If lErro <> SUCESSO Then gError 188396

    GL_objMDIForm.MousePointer = vbDefault
    
    Unload Me

    Exit Sub

Erro_BotaoDesapurar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 188395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", gErr)

        Case 188396

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188397)

    End Select

    Exit Sub

End Sub

Private Function Desapura_Exercicio() As Long

Dim sContaResultado As String
Dim sHistorico As String
Dim sConta As String
Dim iExercicio As Integer
Dim iLote As Integer
Dim lErro As Long
Dim colPlanoConta As New Collection
Dim colContaCategoria As New Collection
Dim objPlanoConta As New ClassPlanoConta
Dim objContaCategoria As ClassContaCategoria
Dim colContasApuracao As New Collection
Dim objCTBConfig As New ClassCTBConfig
Dim vbMsgRes As VbMsgBoxResult
Dim sContaResultadoNivel1 As String
Dim sNomeArqParam As String
Dim lExercicio As Long

On Error GoTo Erro_Desapura_Exercicio
    
    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 188392
    
    lErro = CF("Rotina_Desapura_Exercicio", sNomeArqParam, giFilialEmpresa, iExercicio, iLote, sHistorico, sContaResultado, colContasApuracao)
    If lErro <> SUCESSO Then gError 188393

    Desapura_Exercicio = SUCESSO

    Exit Function

Erro_Desapura_Exercicio:

    Desapura_Exercicio = gErr

    Select Case gErr

        Case 188392, 188393
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188394)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Desapuração de Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DesapuraExercicio"
    
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



