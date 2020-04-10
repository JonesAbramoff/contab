VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ExportarNFiscalOcx 
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   ScaleHeight     =   2595
   ScaleWidth      =   8370
   Begin VB.TextBox NomeDiretorio 
      Enabled         =   0   'False
      Height          =   330
      Left            =   135
      TabIndex        =   17
      Top             =   2025
      Width           =   4260
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4635
      TabIndex        =   16
      Top             =   165
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   4620
      TabIndex        =   15
      Top             =   690
      Width           =   2190
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   7005
      ScaleHeight     =   495
      ScaleWidth      =   1125
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   150
      Width           =   1185
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   615
         Picture         =   "ExportarNFiscalOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExecutar 
         Height          =   345
         Left            =   75
         Picture         =   "ExportarNFiscalOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Executa a rotina"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   750
      Left            =   165
      TabIndex        =   7
      Top             =   60
      Width           =   4245
      Begin MSComCtl2.UpDown UpDownPeriodoDe 
         Height          =   330
         Left            =   1665
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoDe 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownPeriodoAte 
         Height          =   330
         Left            =   3750
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoAte 
         Height          =   330
         Left            =   2775
         TabIndex        =   1
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoAte 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2355
         TabIndex        =   11
         Top             =   300
         Width           =   450
      End
      Begin VB.Label LabelPeriodoDe 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   270
         TabIndex        =   10
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Frame FrameNFiscal 
      Caption         =   "Nota Fiscal"
      Height          =   750
      Left            =   165
      TabIndex        =   6
      Top             =   885
      Width           =   4230
      Begin MSMask.MaskEdBox NFiscalDe 
         Height          =   315
         Left            =   675
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalAte 
         Height          =   315
         Left            =   2775
         TabIndex        =   3
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNFiscalAte 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2340
         TabIndex        =   13
         Top             =   330
         Width           =   450
      End
      Begin VB.Label LabelNFicalDe 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   285
         TabIndex        =   12
         Top             =   330
         Width           =   390
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Localização do Arquivo"
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
      Height          =   285
      Left            =   150
      TabIndex        =   18
      Top             =   1740
      Width           =   2145
   End
End
Attribute VB_Name = "ExportarNFiscalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iListIndexDefault As Integer

'***** FUNÇÕES DE INICIALIZAÇÃO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long
    
On Error GoTo Erro_Form_Load
    
    iListIndexDefault = Drive1.ListIndex
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159825)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159826)

    End Select

    Exit Function

End Function

'*** EVENTOS CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
   
    'Verifica se os campos obrigatorios estao preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 128102
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 128103
    
    'data inicial não pode ser maior que a data final
    If Len(Trim(PeriodoDe.ClipText)) <> 0 And Len(Trim(PeriodoAte.ClipText)) <> 0 Then

         If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 128104

    End If
    
    'Nota Fiscal inicial não pode ser maior que a Nota Fiscal final
    If Len(Trim(NFiscalDe.ClipText)) <> 0 And Len(Trim(NFiscalAte.ClipText)) <> 0 Then

         If StrParaLong(NFiscalDe.Text) > StrParaLong(NFiscalAte.Text) Then gError 128105

    End If
    
    lErro = CF("NFiscais_Exportar", StrParaDate(PeriodoDe.Text), StrParaDate(PeriodoAte.Text), StrParaLong(NFiscalDe.Text), StrParaLong(NFiscalAte.Text), NomeDiretorio.Text)
    If lErro <> SUCESSO Then gError 130006
    
    Unload Me
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
   
    Exit Sub
    
Erro_BotaoExecutar_Click:

    Select Case gErr
     
        Case 130006
        
        Case 128102
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA1", gErr)
        
        Case 128103
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 128104
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 128105
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_INICIAL_MAIOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159827)
            
    End Select
    
    'Transforma a ampulheta em ponteiro
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Call Unload(Me)

End Sub

Private Sub UpDownPeriodoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 128106

    End If

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case 128106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159828)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 128107

    End If

    Exit Sub

Erro_UpDownPeriodoDe_UpClick:

    Select Case gErr

        Case 128107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159829)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 128108

    End If

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case 128108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159830)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 128109

    End If

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case 128109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159831)

    End Select

    Exit Sub

End Sub

'*** EVENTOS VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError 128110

    Exit Sub

Erro_PeriodoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 128110
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159832)
            
    End Select
    
    Exit Sub

End Sub

Private Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError 128111

    Exit Sub

Erro_PeriodoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 128111
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159833)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTOS VALIDATE DOS CONTROLES - FIM ***

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Exportação de Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExportarNFiscal"
    
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
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'**** fim do trecho a ser copiado *****

Private Sub Dir1_Change()

    NomeDiretorio = Dir1.Path

End Sub

Private Sub Dir1_Click()

On Error GoTo Erro_Dir1_Click

    Exit Sub
    
Erro_Dir1_Click:

    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159834)
    
    Exit Sub

End Sub

Private Sub Drive1_Change()

On Error GoTo Erro_Drive1_Change

    Dir1.Path = Drive1.Drive
       
    Exit Sub

Erro_Drive1_Change:

    Select Case Err
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159835)

    End Select

    Drive1.ListIndex = iListIndexDefault
    
    Exit Sub
    
End Sub

Private Sub Drive1_GotFocus()
    
    iListIndexDefault = Drive1.ListIndex

End Sub


