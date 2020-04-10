VERSION 5.00
Begin VB.UserControl RelOpFlCxTotalOcx 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   ScaleHeight     =   2610
   ScaleWidth      =   6060
   Begin VB.Frame Frame1 
      Caption         =   "Fluxo de Caixa"
      Height          =   810
      Left            =   480
      TabIndex        =   12
      Top             =   1665
      Width           =   5430
      Begin VB.OptionButton Sintetico 
         Caption         =   "Sintético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         TabIndex        =   14
         Top             =   345
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton Analitico 
         Caption         =   "Analítico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   585
         TabIndex        =   13
         Top             =   345
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4800
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   225
      Width           =   1140
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpFlCxTotalOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpFlCxTotalOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
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
      Height          =   600
      Left            =   4800
      Picture         =   "RelOpFlCxTotalOcx.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   915
      Width           =   1140
   End
   Begin VB.ComboBox Identificacao 
      Height          =   315
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   2790
   End
   Begin VB.Label Descricao 
      Height          =   240
      Left            =   1650
      TabIndex        =   7
      Top             =   870
      Width           =   2820
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
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
      Left            =   2565
      TabIndex        =   6
      Top             =   1305
      Width           =   945
   End
   Begin VB.Label DataFinal 
      Height          =   240
      Left            =   3615
      TabIndex        =   5
      Top             =   1290
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   555
      TabIndex        =   4
      Top             =   870
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Identificação:"
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
      Left            =   300
      TabIndex        =   3
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Base:"
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
      Left            =   525
      TabIndex        =   2
      Top             =   1305
      Width           =   960
   End
   Begin VB.Label DataInicial 
      Height          =   240
      Left            =   1575
      TabIndex        =   1
      Top             =   1290
      Width           =   930
   End
End
Attribute VB_Name = "RelOpFlCxTotalOcx"
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

Private Sub Identificacao_Click()

Dim lErro As Long
Dim iIndice As Integer, sFluxo As String

On Error GoTo Erro_Identificacao_Click

    If Identificacao.ListIndex = -1 Then Exit Sub

    'Pega o nome do fluxo atual
    sFluxo = Identificacao.Text

    'Exibe na tela os dados do fluxo
    lErro = Traz_Fluxo_Tela(sFluxo)
    If lErro <> SUCESSO Then gError 193301

    Exit Sub

Erro_Identificacao_Click:

    Select Case gErr

        Case 193301

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193302)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim objFluxo As New ClassFluxo
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    'Faz Critica se data inicial é maior que data Final
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 193303

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 193304
    
    lErro = objRelOpcoes.IncluirParametro("TFLUXO", Identificacao.Text)
    If lErro <> AD_BOOL_TRUE Then gError 193305

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 193303 To 193305

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193306)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 193307

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 193307
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193308)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 193309

    If Analitico.Value = True Then
        gobjRelatorio.sNomeTsk = "FlCxAnTl"
    Else
        gobjRelatorio.sNomeTsk = "FlCxSiTl"
    End If


    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 193309

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193310)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

    Identificacao.ListIndex = -1
    Descricao.Caption = ""
    DataInicial.Caption = ""
    DataFinal.Caption = ""

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colFluxo As New Collection
Dim objFluxo As ClassFluxo

On Error GoTo Erro_Form_Load

    lErro = CF("Fluxo_Le_Todos", colFluxo)
    If lErro <> SUCESSO Then gError 193311

    For Each objFluxo In colFluxo

        Identificacao.AddItem objFluxo.sFluxo
        Identificacao.ItemData(Identificacao.NewIndex) = objFluxo.lFluxoId

    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 193311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193312)

    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa Analítico Total"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFlCxAnTotal"

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Verificar se o cliente foi preenchido
    If Len(Trim(Identificacao.Text)) = 0 Then gError 193313

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 193313
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_FLUXO_VAZIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193314)

    End Select

    Exit Function

End Function

Private Function Traz_Fluxo_Tela(sFluxo As String) As Long
'Coloca na Tela os dados do Fluxo passado como parametro

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Traz_Fluxo_Tela

    objFluxo.sFluxo = sFluxo
    objFluxo.iFilialEmpresa = giFilialEmpresa

    'Le o fluxo passado como parametro
    lErro = CF("Fluxo_Le", objFluxo)
    If lErro <> SUCESSO And lErro <> 20104 Then gError 193315

    If lErro = 20104 Then gError 193316

    'passa os dados para a Tela
    Descricao.Caption = objFluxo.sDescricao
    DataInicial.Caption = Format(objFluxo.dtDataBase, "dd/mm/yyyy")
    DataFinal.Caption = Format(objFluxo.dtDataFinal, "dd/mm/yyyy")

    Traz_Fluxo_Tela = SUCESSO

    Exit Function

Erro_Traz_Fluxo_Tela:

    Traz_Fluxo_Tela = gErr

    Select Case gErr

        Case 193315

        Case 193316
            Call Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_CADASTRADO", gErr, objFluxo.sFluxo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193317)

    End Select

    Exit Function

End Function

