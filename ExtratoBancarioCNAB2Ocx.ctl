VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ExtratoBancarioCNAB2Ocx 
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ScaleHeight     =   2670
   ScaleWidth      =   4050
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   1065
      ScaleHeight     =   495
      ScaleWidth      =   1710
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1935
      Width           =   1770
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1215
         Picture         =   "ExtratoBancarioCNAB2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   75
         Width           =   1050
      End
   End
   Begin VB.CommandButton BotaoIntAtualiza 
      Caption         =   "Interromper Atualização"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   405
      TabIndex        =   3
      Top             =   1395
      Width           =   3105
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   1995
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   255
      TabIndex        =   2
      Top             =   915
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Height          =   210
      Left            =   165
      Top             =   360
      Width           =   690
      ForeColor       =   0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Arquivo:"
   End
   Begin VB.Label LabelNomeArq 
      Height          =   210
      Left            =   990
      Top             =   360
      Width           =   2790
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ExtratoBancarioCNAB2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private bAtualizacaoInterrompida As Boolean
Dim giBanco As Integer
Dim gsNomeArq As String
Dim objExtrCNAB As ClassExtrCNAB

Function Trata_Parametros(iBanco As Integer, sNomeArq As String) As Long
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giBanco = iBanco
    gsNomeArq = sNomeArq

    LabelNomeArq.Caption = sNomeArq
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159893)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    bAtualizacaoInterrompida = False
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err
        
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159894)

    End Select

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objExtrCNAB = Nothing
    
End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoAtualizar_Click

    BotaoAtualizar.Enabled = False

    Set objExtrCNAB = New ClassExtrCNAB
    lErro = objExtrCNAB.Extrato_Abrir(giBanco, gsNomeArq)
    If lErro <> SUCESSO Then Error 7166

    BotaoIntAtualiza.Enabled = True
    Timer1.Enabled = True
    
    Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case Err
        
        Case 7166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159895)

    End Select

    Set objExtrCNAB = Nothing
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    If (objExtrCNAB Is Nothing) Then
    
        Unload Me
    
    End If
    
End Sub

Private Sub BotaoIntAtualiza_Click()

    bAtualizacaoInterrompida = True
    BotaoIntAtualiza.Enabled = False
    
End Sub

Private Sub Timer1_Timer()
Dim lErro As Long
Dim iPerc As Integer
Dim iAcabou As Integer
On Error GoTo Erro_Timer1_Timer

    If Not (objExtrCNAB Is Nothing) Then
    
        If (bAtualizacaoInterrompida = False) Then
        
            lErro = objExtrCNAB.Extrato_LerReg(iPerc, iAcabou)
            If lErro <> SUCESSO Then Error 7167
            
            ProgressBar1.Value = iPerc
            
            'se já acabou de ler os registros
            If iAcabou <> 0 Then
                
                BotaoIntAtualiza.Enabled = False
                Set objExtrCNAB = Nothing
                
            End If
                
        Else
        
            Set objExtrCNAB = Nothing
        
        End If
        
    End If
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case Err
        
        Case 7167
        Case 7168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159896)

    End Select

    Timer1.Enabled = False
    Set objExtrCNAB = Nothing

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXTRATO_BANCARIO_CNAB2
    Set Form_Load_Ocx = Me
    Caption = "Recepção de Extrato Bancário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExtratoBancarioCNAB2"
    
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

'***** fim do trecho a ser copiado ******



Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeArq_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeArq, Source, X, Y)
End Sub

Private Sub LabelNomeArq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeArq, Button, Shift, X, Y)
End Sub

