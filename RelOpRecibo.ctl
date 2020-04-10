VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpReciboOcx 
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LockControls    =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   5850
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
      Left            =   2040
      Picture         =   "RelOpRecibo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1590
      TabIndex        =   0
      Top             =   285
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorRecebido 
      Height          =   300
      Left            =   1590
      TabIndex        =   1
      Top             =   750
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Faturas 
      Height          =   300
      Left            =   1590
      TabIndex        =   2
      Top             =   1215
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Faturas Pagas:"
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
      Left            =   210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1275
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Recebido:"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   810
      Width           =   1575
   End
   Begin VB.Label LabelCliente 
      Caption         =   "Cliente:"
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
      Left            =   855
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   300
      Width           =   645
   End
End
Attribute VB_Name = "RelOpReciboOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros


    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182869)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCria As Integer

On Error GoTo Erro_BotaoExecutar_Click

    If Len(Trim(Cliente.Text)) > 0 Then

        iCria = 0 'Não deseja criar cliente caso não exista
        lErro = TP_cliente_Le2(Cliente, objCliente)
        If lErro <> SUCESSO Then gError 182870

    End If

    lErro = CF("RelRecibo_Prepara", 0, objCliente.lCodigo, objCliente.sRazaoSocial, StrParaDbl(ValorRecebido.Text), Faturas.Text)
    If lErro <> SUCESSO Then gError 182871

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 182870, 182871

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182872)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCria As Integer

On Error GoTo Erro_Cliente_Validate

    If Len(Trim(Cliente.Text)) > 0 Then

        iCria = 0 'Não deseja criar cliente caso não exista
        lErro = TP_cliente_Le2(Cliente, objCliente, iCria)
        If lErro <> SUCESSO Then gError 182873
        
        Cliente.Text = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 182873
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182874)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCliente = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182875)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    If Len(Trim(Cliente.Text)) > 0 Then
        'Preenche com o Cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI_L
    Set Form_Load_Ocx = Me
    Caption = "Recibo de Pagamentos do Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRecibo"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub ValorRecebido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorRecebido_Validate

    'Verifica se ValorRecebido está preenchida
    If Len(Trim(ValorRecebido.Text)) <> 0 Then

        'Critica a ValorRecebido
        lErro = Valor_Positivo_Critica(ValorRecebido.Text)
        If lErro <> SUCESSO Then gError 182876
        
        ValorRecebido.Text = Format(ValorRecebido.Text, "STANDARD")

    End If

    Exit Sub

Erro_ValorRecebido_Validate:

    Cancel = True

    Select Case gErr

        Case 182876

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182877)

    End Select

    Exit Sub

End Sub
