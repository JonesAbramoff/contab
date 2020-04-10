VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ArqComissoes 
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   ScaleHeight     =   1830
   ScaleWidth      =   5295
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ArqComissoesInpal.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ArqComissoesInpal.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ArqComissoesInpal.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FramePeriodo 
      Caption         =   "Período"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   367
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   315
         Left            =   4520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Top             =   367
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   315
         Left            =   2115
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelFim 
         AutoSize        =   -1  'True
         Caption         =   "Fim:"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelInicio 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   570
      End
   End
End
Attribute VB_Name = "ArqComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Public Function Form_Load_Ocx() As Object
    '??? criar IDH Parent.HelpContextID = IDH_ARQCOMISSOES
    Set Form_Load_Ocx = Me
    Caption = "Arquivo para o sistema de Comissões"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ArqComissoes"
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
    'm_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******


Private Sub LabelInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelInicio, Source, X, Y)
End Sub

Private Sub LabelFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFim, Button, Shift, X, Y)
End Sub

Private Sub Form_Load()
    lErro_Chama_Tela = SUCESSO
End Sub

Function Trata_Parametros() As Long
    Trata_Parametros = SUCESSO
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
Dim iAlterado As Integer
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Altera o ponteiro do mouse para ampulheta
    MousePointer = vbHourglass
    
    'Se a data início não foi preenchida => erro
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 94845
    
    'Se a data fim não foi preenchida => erro
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 94846
    
    'Se a data início é maior do que a data fim => erro
    If DataInicial.Text > DataFinal.Text Then gError 94847
    
    'Gera o arquivo
    lErro = CF("Gera_ArqComissoes", DataInicial.Text, DataFinal.Text)
    If lErro <> SUCESSO Then gError 94844
    
    'Retorna o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 94844
        
        Case 94845
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            
        Case 94846
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            
        Case 94847
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    'Retorna o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
    Limpa_Tela (Me)
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub DataInicial_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataInicial, iAlterado)
End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) > 0 Then
    
        'verifica se a data final é válida
        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 94851
    
    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 94851

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)
End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(Trim(DataFinal.ClipText)) > 0 Then
    
        'verifica se a data final é válida
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 94852
    
    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 94852

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_DownClick

    DataInicial.SetFocus

    'verifica se a data foi preenchida
    If Len(Trim(DataInicial.ClipText)) > 0 Then

        sData = DataInicial.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 94848

        DataInicial.PromptInclude = False
        DataInicial.Text = sData
        DataInicial.PromptInclude = True
        
    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 94848

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_UpClick

    DataInicial.SetFocus

    'Verifica de a data foi preenchida
    If Len(Trim(DataInicial.ClipText)) > 0 Then

        sData = DataInicial.Text

        'aumenta a data de um dia
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 94849

        DataInicial.PromptInclude = False
        DataInicial.Text = sData
        DataInicial.PromptInclude = True

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 94849

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    'verifica se a data foi preenchida
    If Len(Trim(DataFinal.ClipText)) > 0 Then

        sData = DataFinal.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 94848

        DataFinal.PromptInclude = False
        DataFinal.Text = sData
        DataFinal.PromptInclude = True
        
    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 94848

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    'Verifica de a data foi preenchida
    If Len(Trim(DataFinal.ClipText)) > 0 Then

        sData = DataFinal.Text

        'aumenta a data de um dia
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 94849

        DataFinal.PromptInclude = False
        DataFinal.Text = sData
        DataFinal.PromptInclude = True

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 94849

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub
