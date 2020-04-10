VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPrazoEnt 
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   ScaleHeight     =   1905
   ScaleWidth      =   5160
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   885
      Left            =   150
      TabIndex        =   3
      Top             =   165
      Width           =   4815
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "RelOpPrazoEntOcx.ctx":0000
         Left            =   690
         List            =   "RelOpPrazoEntOcx.ctx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   1365
      End
      Begin MSMask.MaskEdBox Ano 
         Height          =   300
         Left            =   3255
         TabIndex        =   6
         Top             =   330
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "0"
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelMes 
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   225
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
      Begin VB.Label LabelAno 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2820
         TabIndex        =   4
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   600
      Left            =   3375
      Picture         =   "RelOpPrazoEntOcx.ctx":0094
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Fechar"
      Top             =   1185
      Width           =   1575
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
      Left            =   180
      Picture         =   "RelOpPrazoEntOcx.ctx":0212
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1170
      Width           =   1575
   End
End
Attribute VB_Name = "RelOpPrazoEnt"
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


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 91655
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 91655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim sMes As String

On Error GoTo Erro_PreencherRelOp
        
    If Mes.ListIndex = -1 Then gError 91657
    If Len(Trim((Ano.ClipText))) = 0 Then gError 91656
   
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 91658
            
    lErro = Prepara_Param_RelOp(dtDataIni, dtDataFim)
    If lErro <> SUCESSO Then gError 91664
            
    lErro = objRelOpcoes.IncluirParametro("DINI", CStr(dtDataIni))
    If lErro <> AD_BOOL_TRUE Then gError 91659

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(dtDataFim))
    If lErro <> AD_BOOL_TRUE Then gError 91660
            
    sMes = Mes.List(Mes.ListIndex)
            
    lErro = objRelOpcoes.IncluirParametro("TMES", sMes)
    If lErro <> AD_BOOL_TRUE Then gError 91665

    lErro = objRelOpcoes.IncluirParametro("NANO", Ano.Text)
    If lErro <> AD_BOOL_TRUE Then gError 91666
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr
        
        Case 91656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
                    
        Case 91657
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
                   
        Case 91658 To 91660
                   
        Case 91664, 91665, 91666
                          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Ano_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Ano)

End Sub

Private Sub Ano_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Ano_Validate

    If Len(Ano.ClipText) > 0 Then

        lErro = Valor_Double_Critica(Ano.Text)
        If lErro <> SUCESSO Then gError 91661

    End If

    Exit Sub

Erro_Ano_Validate:

    Cancel = True

    Select Case gErr

        Case 91661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 91662

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 91662

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Relação de Prazos de Entrega"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrazoEnt"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'    If KeyCode = KEYCODE_BROWSER Then
'        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
'    End If
        
End Sub

Private Sub LabelAno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAno, Source, X, Y)
End Sub

Private Sub LabelAno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAno, Button, Shift, X, Y)
End Sub

Private Sub LabelMes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMes, Source, X, Y)
End Sub

Private Sub LabelMes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMes, Button, Shift, X, Y)
End Sub

Function Prepara_Param_RelOp(dtDataIni As Date, dtDataFim As Date) As Long
'Extrai Faixa de Datas para serem passadas com Parâmtero do Relatório

Dim lErro As Long

On Error GoTo Erro_Prepara_Param_RelOp
    
    dtDataIni = StrParaDate("01/" & Mes.ItemData(Mes.ListIndex) & "/" & _
    Ano.Text)

    dtDataFim = DateAdd("m", 1, dtDataIni) - 1

    Prepara_Param_RelOp = SUCESSO
    
    Exit Function
    
Erro_Prepara_Param_RelOp:

    Prepara_Param_RelOp = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function

End Function
