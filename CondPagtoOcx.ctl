VERSION 5.00
Begin VB.UserControl CondPagtoOcx 
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   2325
   ScaleWidth      =   5220
   Begin VB.ComboBox ComboMoeda 
      Height          =   288
      Left            =   1608
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1230
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton BotaoCancela 
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
      Height          =   345
      Left            =   2700
      TabIndex        =   4
      Top             =   1695
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condição de Pagamento para geração de Pedidos de Compra"
      Height          =   870
      Left            =   240
      TabIndex        =   1
      Top             =   195
      Width           =   4740
      Begin VB.OptionButton CondPagto 
         Caption         =   "A Vista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   750
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton CondPagto 
         Caption         =   "A Prazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   2
         Top             =   420
         Width           =   1005
      End
   End
   Begin VB.CommandButton BotaoOK 
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
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   1695
      Width           =   1845
   End
   Begin VB.Label LabelMoeda 
      AutoSize        =   -1  'True
      Caption         =   "Moeda:"
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
      Height          =   192
      Left            =   936
      TabIndex        =   6
      Top             =   1296
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "CondPagtoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjPedidoCompras As ClassPedidoCompras

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

'Só abre se os 2 preços (a vista e a prazo estiverem preenchidos)

Private Sub BotaoOK_Click()

    If CondPagto(1).Value = True Then
        gobjPedidoCompras.iCondicaoPagto = CONDPAGTO_VISTA
    Else
        gobjPedidoCompras.iCondicaoPagto = CONDPAGTO_PRAZO
    End If
    
    'Guarda a Moeda
    If ComboMoeda.Visible = True Then gobjPedidoCompras.imoeda = Codigo_Extrai(ComboMoeda.List(ComboMoeda.ListIndex))
    
    Unload Me

End Sub

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO

End Sub

Public Function Trata_Parametros(objPedidoCompras As ClassPedidoCompras, Optional ColMoedasUsadas As Collection) As Long

Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros
    
    Set gobjPedidoCompras = objPedidoCompras
    
    gobjPedidoCompras.iCondicaoPagto = 0
    
    'Se a coleção de moedas foi passada
    If Not IsMissing(ColMoedasUsadas) Then
    
        If ColMoedasUsadas.Count > 1 Then
            
            'Torna a combo (e o sue label) visivel
            LabelMoeda.Visible = True
            ComboMoeda.Visible = True
            
            'carrega a combo com as moedas passadas
            For iIndice = 1 To ColMoedasUsadas.Count
                ComboMoeda.AddItem ColMoedasUsadas.Item(iIndice)
            Next
            
            ComboMoeda.ListIndex = 0
            
        Else
        
            gobjPedidoCompras.imoeda = Codigo_Extrai(ColMoedasUsadas.Item(1))
            
        End If
        
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154673)
            
    End Select
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Condição de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CondPagto"
    
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

