VERSION 5.00
Begin VB.UserControl ParcelamentoTEF 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   ScaleHeight     =   1290
   ScaleWidth      =   4395
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
      Left            =   2445
      TabIndex        =   3
      Top             =   840
      Width           =   1485
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
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
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1485
   End
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      ItemData        =   "ParcelamentoTEF.ctx":0000
      Left            =   1395
      List            =   "ParcelamentoTEF.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   285
      Width           =   1230
   End
End
Attribute VB_Name = "ParcelamentoTEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjMovCaixa As New ClassMovimentoCaixa

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Function Trata_Parametros(objMovCaixa As ClassMovimentoCaixa) As Long
    
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
    
    Set gobjMovCaixa = objMovCaixa
    
    For iIndice = 1 To gcolAdmMeioPagto.Count
        Set objAdmMeioPagto = gcolAdmMeioPagto.Item(iIndice)
        If objAdmMeioPagto.iCodigo = objMovCaixa.iAdmMeioPagto Then
            'Adiciona na combo de Parcelamento
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                Parcelamento.AddItem objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
                Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
            Next
        End If
    Next
    
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

On Error GoTo Erro_BotaoOk_Click

    'se Adm não selecionado --> Erro.
    If Parcelamento.ListIndex = -1 Then gError 99675
    
    gobjMovCaixa.iParcelamento = Codigo_Extrai(Parcelamento.Text)
    
    giRetornoTela = vbOK
    
    Unload Me
    
     Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 99675
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_SELECIONADO, gErr)
        
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164309)

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
    Caption = "Seleção de Parcelamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ParcelamentoTEF"
    
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

'***** fim do trecho a ser copiado ******


