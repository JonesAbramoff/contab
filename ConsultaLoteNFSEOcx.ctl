VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ConsultaLoteNFSEOcx 
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ScaleHeight     =   1305
   ScaleWidth      =   3750
   Begin VB.CommandButton BotaoConsulta 
      Caption         =   "Consultar"
      Height          =   735
      Left            =   2010
      Picture         =   "ConsultaLoteNFSEOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Consultar"
      Top             =   300
      Width           =   825
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   735
      Left            =   2985
      Picture         =   "ConsultaLoteNFSEOcx.ctx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fechar"
      Top             =   300
      Width           =   480
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   495
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label LoteLbl 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   540
      Width           =   450
   End
End
Attribute VB_Name = "ConsultaLoteNFSEOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoLotes As AdmEvento
Attribute objEventoLotes.VB_VarHelpID = -1


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)


End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()
    
On Error GoTo Erro_Form_Load
    
    Set objEventoLotes = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207048)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoConsulta_Click()

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sNFSEEXE As String

On Error GoTo Erro_BotaoConsulta_Click

    'verifica se o codigo foi preenchido
    If Len(Lote.Text) = 0 Then gError 207049

    lErro = CF("NFSE_Obter_EXE", giFilialEmpresa, sNFSEEXE)
    If lErro <> SUCESSO Then gError 201196
                
    lErro = WinExec(sNFSEEXE & " Consulta " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & Lote.Text, SW_NORMAL)

    Call Rotina_Aviso(vbOK, "AVISO_INICIO_CONSULTA_LOTE_NFE", Lote.Text)
    
    Lote.Text = ""

    Exit Sub
    
Erro_BotaoConsulta_Click:

    Select Case gErr

        Case 207049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PREENCHIDO", gErr)

        Case 201196

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207050)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta de Lote de Envio de NFSE"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConsultaLoteNFSE"
    
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


Private Sub LoteLbl_Click()

Dim objRPSWEBLoteView As New ClassRPSWEBLoteView
Dim colSelecao As Collection

    If Len(Trim(Lote.Text)) > 0 Then
        objRPSWEBLoteView.lLote = Lote.Text
    End If

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBLoteViewLista", colSelecao, objRPSWEBLoteView, objEventoLotes)

End Sub

Private Sub objEventoLotes_evSelecao(obj1 As Object)

Dim objRPSWEBLoteView As ClassRPSWEBLoteView
Dim bCancel As Boolean

    Set objRPSWEBLoteView = obj1

    'Preenche o Cliente com o Cliente selecionado
    Lote.Text = objRPSWEBLoteView.lLote

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
   ' Parent.UnloadDoFilho
    Set objEventoLotes = Nothing
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call LoteLbl_Click
        End If
          
    End If

End Sub



