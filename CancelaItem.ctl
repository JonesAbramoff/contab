VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl CancelaItem 
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   3900
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1485
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1965
      TabIndex        =   2
      Top             =   720
      Width           =   1485
   End
   Begin MSMask.MaskEdBox Item 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   165
      Width           =   435
   End
End
Attribute VB_Name = "CancelaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As New ClassVenda

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
    Set gobjVenda = objVenda
    
    If gobjVenda.objCupomFiscal.iItem > 0 Then Item.Text = gobjVenda.objCupomFiscal.iItem
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Private Sub BotaoCancelar_Click()

    giRetornoTela = vbCancel
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()
    
Dim lErro As Long
Dim objItem As ClassItemCupomFiscal
Dim bAchou As Boolean

On Error GoTo Erro_BotaoOk_Click

    If Item.Text = "" Then gError 112424
    
    gobjVenda.objCupomFiscal.iItem = StrParaInt(Item.Text)
    
    bAchou = False
    For Each objItem In gobjVenda.objCupomFiscal.colItens
        If objItem.iItem = gobjVenda.objCupomFiscal.iItem Then
            If objItem.iStatus = STATUS_CANCELADO Then gError 112425
            bAchou = True
        End If
    Next
    
    If Not (bAchou) Then gError 112426
    
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 112424
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_NAO_PREENCHIDO1, gErr)
        
        Case 112425
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_CANCELADO, gErr)
        
        Case 112426
            Call Rotina_ErroECF(vbOKOnly, ERRO_ITEM_NAO_EXISTENTE1, gErr)
        
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144151)

    End Select

    Exit Sub

End Sub


Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Item_Validate
    
    If Len(Trim(Item.Text)) > 0 Then
    
        lErro = Valor_Inteiro_Critica(Item.Text)
        If lErro <> SUCESSO Then gError 115001
        
    End If
        
    Exit Sub
    
Erro_Item_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 115001
            Call Rotina_Erro(vbOKOnly, ERRO_NUMERO_NAO_INTEIRO1, gErr, Item.Text)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144152)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Clique em f5
    If KeyCode = vbKeyF5 Then
        If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
        Call BotaoOk_Click
    End If

    'Clique em esc
    If KeyCode = vbKeyEscape Then
        If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
        Call BotaoCancelar_Click
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Item a ser cancelado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CancelaItem"
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'***** fim do trecho a ser copiado ******



