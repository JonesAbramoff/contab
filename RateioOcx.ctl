VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl RateioOcx 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   3255
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   420
      Picture         =   "RateioOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancel 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1860
      Picture         =   "RateioOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1335
      Width           =   975
   End
   Begin MSMask.MaskEdBox Rateio 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   285
      Left            =   1815
      TabIndex        =   1
      Top             =   735
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código do Rateio : "
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
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   150
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor :"
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
      Left            =   1005
      TabIndex        =   5
      Top             =   795
      Width           =   570
   End
End
Attribute VB_Name = "RateioOcx"
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

Dim objRateioOn1 As ClassRateioOn
Dim objConfirmaTela1 As AdmConfirmaTela
Private WithEvents objEventoRateio As AdmEvento
Attribute objEventoRateio.VB_VarHelpID = -1

Private Sub BotaoCancel_Click()
    
    Set objRateioOn1 = Nothing
    
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoRateio = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case Err
        
        Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166079)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoOK_Click()

Dim iCodigo As Integer
Dim dValor As Double
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    'Verifica se o Rateio esta preenchido
    If Len(Rateio.ClipText) = 0 Then Error 11360
    
    'Verifica se algum valor foi digitado
    If Len(Valor.ClipText) = 0 Then Error 11361
    
    objRateioOn1.iCodigo = CInt(Rateio.Text)
    objRateioOn1.dPercentual = CDbl(Valor.Text)
    
    'Verifica se o valor é positivo
    If objRateioOn1.dPercentual <= 0 Then Error 11362
    
    objConfirmaTela1.iTelaOK = OK
            
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case Err
    
        Case 11360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_RATEIO_NAO_PREENCHIDO", Err)
        
        Case 11361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO", Err)
        
        Case 11362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err)
            
        Case 11368
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166080)
            
    End Select
    
    Exit Sub

End Sub


Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoRateio = Nothing
    
    Set objRateioOn1 = Nothing
    Set objConfirmaTela1 = Nothing
        
End Sub

Private Sub Label1_Click()

Dim objRateioOn As New ClassRateioOn
Dim colSelecao As Collection

    'Verificica se ja existe um codigo de rateio digitado
    If Len(Rateio.Text) = 0 Then
        objRateioOn.iCodigo = 0
    Else
        objRateioOn.iCodigo = CInt(Rateio.ClipText)
    End If

    objRateioOn.iSeq = 0

    Call Chama_Tela_Modal("RateioOnLista", colSelecao, objRateioOn, objEventoRateio)

End Sub

Private Sub objEventoRateio_evSelecao(obj1 As Object)

Dim objRateioOn As ClassRateioOn
    
    Set objRateioOn = obj1
    
    'Coloca na tela o codigo do rateio escolhido
    Rateio.Text = CStr(objRateioOn.iCodigo)
    
    Exit Sub
    
End Sub

Function Trata_Parametros(objRateioOn As ClassRateioOn, objConfirmaTela As AdmConfirmaTela) As Long
    
    Set objRateioOn1 = objRateioOn
    Set objConfirmaTela1 = objConfirmaTela
    
    objConfirmaTela1.iTelaOK = CANCELA
    
    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RATEIO
    Set Form_Load_Ocx = Me
    Caption = "Rateio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Rateio"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Rateio Then
            Call Label1_Click
        End If
    
    End If

End Sub



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

