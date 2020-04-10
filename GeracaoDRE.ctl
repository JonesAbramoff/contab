VERSION 5.00
Begin VB.UserControl GeracaoDREOcx 
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   6615
   Begin VB.ComboBox ComboGrupoEmp 
      Height          =   315
      Left            =   1905
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   5970
      Width           =   2415
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
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
      Left            =   5460
      Picture         =   "GeracaoDRE.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4830
      Width           =   1005
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todos"
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
      Left            =   5460
      Picture         =   "GeracaoDRE.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1005
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      Left            =   5265
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2970
      Width           =   1215
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "GeracaoDRE.ctx":21FC
      Left            =   1095
      List            =   "GeracaoDRE.ctx":21FE
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2970
      Width           =   1860
   End
   Begin VB.ListBox Modelos 
      Columns         =   2
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   3705
      Width           =   5280
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   1095
      TabIndex        =   2
      Top             =   2505
      Width           =   5385
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   4800
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   45
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "GeracaoDRE.ctx":2200
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "GeracaoDRE.ctx":237E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   105
         Picture         =   "GeracaoDRE.ctx":28B0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gera o arquivo"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   105
      TabIndex        =   1
      Top             =   690
      Width           =   6390
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   2820
   End
   Begin VB.Label Label7 
      Caption         =   "Grupo de Empresas:"
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
      Left            =   135
      TabIndex        =   17
      Top             =   6015
      Width           =   1755
   End
   Begin VB.Label Label5 
      Caption         =   "Período:"
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
      Left            =   4485
      TabIndex        =   15
      Top             =   3015
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Exercicio:"
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
      Left            =   195
      TabIndex        =   14
      Top             =   3015
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Modelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   13
      Top             =   3405
      Width           =   1755
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Diretório:"
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
      Left            =   255
      TabIndex        =   12
      Top             =   2550
      Width           =   795
   End
End
Attribute VB_Name = "GeracaoDREOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iListIndexDefault As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim sDiretorio As String
Dim colModelos As New Collection
Dim sNomeArqParam As String, iGrupoEmp As Integer

On Error GoTo Erro_BotaoGerar_Click
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 187139
    
    'Exercício não pode ser vazio
    If ComboExercicio.Text = "" Then gError 187140
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\"
    End If
    
    For iIndice = 0 To Modelos.ListCount - 1
        If Modelos.Selected(iIndice) Then
            colModelos.Add Modelos.List(iIndice)
        End If
    Next
    
    If colModelos.Count = 0 Then gError 187141

    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 187142

    If ComboGrupoEmp.ListIndex = -1 Then
        iGrupoEmp = 0
    Else
        iGrupoEmp = ComboGrupoEmp.ItemData(ComboGrupoEmp.ListIndex)
    End If

    lErro = CF("Rotina_Gerar_DRE_DRP", sNomeArqParam, colModelos, ComboExercicio.ItemData(ComboExercicio.ListIndex), ComboPeriodo.ItemData(ComboPeriodo.ListIndex), giFilialEmpresa, sDiretorio, iGrupoEmp)
    If lErro <> SUCESSO Then gError 187143

    Exit Sub
    
Erro_BotaoGerar_Click:

    Select Case gErr
    
        Case 187139
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeDiretorio.SetFocus
            
        Case 187140
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", gErr)
            ComboExercicio.SetFocus
            
        Case 187141
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
            
        Case 187142, 187143
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187144)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    NomeDiretorio.Text = CurDir
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187145)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colModelos As New Collection
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load
    
    iListIndexDefault = Drive1.ListIndex
    
    If Len(Trim(CurDir)) > 0 Then
        Dir1.Path = CurDir
        Drive1.Drive = left(CurDir, 2)
    End If
    
    NomeDiretorio.Text = Dir1.Path
    
    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then gError 187146
    
    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next
    
    'Le os Modelos na tabela RelDRE
    lErro = CF("RelDRE_Le_Modelos_Distintos", RELDRE, colModelos)
    If lErro <> SUCESSO And lErro <> 47101 Then gError 187147
    
    'se  encontrou alguém
    If lErro = SUCESSO Then
    
        'preenche a combo Modelos
        For iIndice = 1 To colModelos.Count
            Modelos.AddItem colModelos.Item(iIndice)
        Next
        
    End If
    
    lErro = CF("GrupoEmp_CarregaCombo", ComboGrupoEmp)
    If lErro <> SUCESSO Then gError 187146
    
    ComboGrupoEmp.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 187146, 187147
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187148)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187149)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187150)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_ARQICMS
    Set Form_Load_Ocx = Me
    Caption = "Geração de DRE\DRP em excel"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoDRE"

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

Function Trata_Parametros(Optional obj1 As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    If Not (obj1 Is Nothing) Then
    
             
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187151)

    End Select

    Exit Function

End Function

Private Sub Dir1_Change()

     NomeDiretorio.Text = Dir1.Path

End Sub

Private Sub Dir1_Click()

On Error GoTo Erro_Dir1_Click

    Exit Sub
    
Erro_Dir1_Click:

    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187152)
    
    Exit Sub

End Sub

Private Sub Drive1_Change()

On Error GoTo Erro_Drive1_Change

    Dir1.Path = Drive1.Drive
       
    Exit Sub

Erro_Drive1_Change:

    Select Case gErr
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187153)

    End Select

    Drive1.ListIndex = iListIndexDefault
    
    Exit Sub
    
End Sub

Private Sub Drive1_GotFocus()
    
    iListIndexDefault = Drive1.ListIndex

End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 187154

    Drive1.Drive = Mid(NomeDiretorio.Text, 1, 2)

    Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 187154, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187155)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click
    
    'se esta vazia
    If ComboExercicio.ListIndex = -1 Then Exit Sub
          
    'preenche a combo com periodo 1
    lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 0)
    If lErro <> SUCESSO Then gError 187156
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case gErr

        Case 187156

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187157)

    End Select

    Exit Sub

End Sub

Function PreencheComboPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long, iConta As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodo

    ComboPeriodo.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then gError 187158

    ComboPeriodo.AddItem ""
    ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = 0

    For iConta = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iConta)
        ComboPeriodo.AddItem objPeriodo.sNomeExterno
        ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = objPeriodo.iPeriodo
    Next

    'mostra o período
    For iConta = 0 To ComboPeriodo.ListCount - 1
        If ComboPeriodo.ItemData(iConta) = iPeriodo Then
            ComboPeriodo.ListIndex = iConta
            Exit For
        End If
    Next

    PreencheComboPeriodo = SUCESSO

    Exit Function

Erro_PreencheComboPeriodo:

    PreencheComboPeriodo = gErr

    Select Case gErr

        Case 187158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187159)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To Modelos.ListCount - 1
        Modelos.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To Modelos.ListCount - 1
        Modelos.Selected(iIndice) = True
    Next

End Sub
