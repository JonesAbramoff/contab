VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFlCxAn1Ocx 
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1515
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3525
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpFlCxAn1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpFlCxAn1Ocx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
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
      Left            =   3540
      Picture         =   "RelOpFlCxAn1Ocx.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   750
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Títulos"
      Height          =   1290
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3270
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Top             =   795
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   2115
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   315
         Left            =   945
         TabIndex        =   10
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   510
         TabIndex        =   4
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   540
         TabIndex        =   3
         Top             =   375
         Width           =   315
      End
   End
End
Attribute VB_Name = "RelOpFlCxAn1Ocx"
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

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    GL_objMDIForm.MousePointer = vbHourglass

    'Faz Critica se data inicial é maior que data Final
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 195439

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 195440
    
    lErro = objRelOpcoes.IncluirParametro("DINIC", EmissaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 195441

    lErro = objRelOpcoes.IncluirParametro("DFIM", EmissaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 195442

    lErro = CF("RelFlCxAn_Prepara", giFilialEmpresa, StrParaDate(EmissaoDe.Text), StrParaDate(EmissaoAte.Text), lNumIntRel)
    If lErro <> SUCESSO Then gError 195443
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then gError 195444

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 195439 To 195444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195445)

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

    If Not (gobjRelatorio Is Nothing) Then gError 195446

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 195446
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195447)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 195448

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 195448

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195449)

    End Select

    Exit Sub

End Sub


Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela(Me)

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(EmissaoDe.ClipText)) = 0 Then Exit Sub

    'Verifica se é uma data válida
    lErro = Data_Critica(EmissaoDe.Text)
    If lErro <> SUCESSO Then gError 195450

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 195450

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195451)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(EmissaoAte.ClipText)) = 0 Then Exit Sub

    'Verifica se é uma data válida
    lErro = Data_Critica(EmissaoAte.Text)
    If lErro <> SUCESSO Then gError 195451

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 195451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195452)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

    EmissaoDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    EmissaoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    lErro_Chama_Tela = SUCESSO

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa Analítico 1"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFlCXAn1"

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

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 195453
    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 195453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195454)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 195455

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 195455
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195456)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 195457
    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 195457

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195458)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 195459

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 195459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195460)

    End Select

    Exit Sub

End Sub

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

    If Len(Trim(EmissaoDe.ClipText)) = 0 Then gError 195434

    If Len(Trim(EmissaoAte.ClipText)) = 0 Then gError 195435

    'data inicial nao pode ser menor que a data atual
    If CDate(EmissaoDe.Text) < gdtDataAtual Then gError 195436

    'data inicial não pode ser maior que a data final
    If CDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then gError 195437

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 195434
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)

        Case 195435
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)

        Case 195436
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MENOR_DATAATUAL", gErr)

        Case 195437
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195438)

    End Select
 
    Exit Function

End Function


