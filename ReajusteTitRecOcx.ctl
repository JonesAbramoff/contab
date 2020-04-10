VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ReajusteTitRecOcx 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ScaleHeight     =   1050
   ScaleWidth      =   5655
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   3780
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   1665
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   120
         Picture         =   "ReajusteTitRecOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Executa a rotina"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "ReajusteTitRecOcx.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "ReajusteTitRecOcx.ctx":05C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox ReajustarAte 
      Height          =   300
      Left            =   2580
      TabIndex        =   0
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      Caption         =   "Reajustar Até (MM/AAAA):"
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
      Left            =   210
      TabIndex        =   5
      Top             =   405
      Width           =   2325
   End
End
Attribute VB_Name = "ReajusteTitRecOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'pegar de algum config um default para proximo mes de reajuste
'o que fazer se um titulo estiver com reajuste "atrasado", por exemplo, for de reajuste mensal
'e o ultimo reajuste tiver sido feito há 3 meses ?
'ver quantas casas decimais para arredondar/truncar indice de reajuste
    
Private Sub BotaoGerar_Click()

Dim lErro As Long, dtData As Date, objProcReajTitRec As New ClassProcReajTitRec

On Error GoTo Erro_BotaoGerar_Click

    If Len(Trim(ReajustarAte.ClipText)) = 0 Then gError 130153
    
    dtData = StrParaDate("01/" & ReajustarAte.Text)
    
    With objProcReajTitRec
    
        .iFilialEmpresa = giFilialEmpresa
        .dtAtualizadoAte = dtData
        .dHoraProc = Time
        .dtDataProc = gdtDataHoje
        .sUsuario = gsUsuario
        
    End With
    
    lErro = CF("ReajusteTitRecs_Processa", objProcReajTitRec)
    If lErro <> SUCESSO Then gError 130154
    
    Call BotaoFechar_Click
    
    Exit Sub
     
Erro_BotaoGerar_Click:

    Select Case gErr
          
        Case 130153
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_ANO_NAO_PREENCHIDO", gErr)
        
        Case 130154
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166218)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    '??? colocar default

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166219)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_FECHAMENTO_MES_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Reajuste de Títulos a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ReajusteTitRec"
    
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

Private Sub ReajustarAte_Validate(Cancel As Boolean)
    
Dim lErro As Long, sData As String

On Error GoTo Erro_ReajustarAte_Validate

    If Len(Trim(ReajustarAte.ClipText)) <> 0 Then
        
        sData = "01/" & ReajustarAte.Text
        lErro = Data_Critica(sData)
        If lErro <> SUCESSO Then gError 130201
        
'        ReajustarAte.PromptInclude = False
'        ReajustarAte.Text = Format(sData, "mm/yyyy")
'        ReajustarAte.PromptInclude = True
        
    End If
    
    Exit Sub
     
Erro_ReajustarAte_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 130201
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166220)
     
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

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

