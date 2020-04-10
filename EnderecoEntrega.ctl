VERSION 5.00
Begin VB.UserControl EnderecoEntrega 
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ScaleHeight     =   3990
   ScaleWidth      =   7710
   Begin VB.TextBox Email 
      Height          =   300
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   15
      Top             =   3570
      Width           =   4950
   End
   Begin VB.ComboBox ComboCidade 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3015
      Width           =   4935
   End
   Begin VB.ComboBox UF 
      Height          =   315
      ItemData        =   "EnderecoEntrega.ctx":0000
      Left            =   1230
      List            =   "EnderecoEntrega.ctx":0055
      TabIndex        =   13
      Text            =   "UF"
      Top             =   2475
      Width           =   900
   End
   Begin VB.TextBox Bairro 
      Height          =   300
      Left            =   1260
      MaxLength       =   60
      TabIndex        =   10
      Top             =   1920
      Width           =   4905
   End
   Begin VB.TextBox Complemento 
      Height          =   300
      Left            =   1275
      MaxLength       =   60
      TabIndex        =   8
      Top             =   1380
      Width           =   4905
   End
   Begin VB.TextBox Numero 
      Height          =   330
      Left            =   1290
      MaxLength       =   60
      TabIndex        =   6
      Top             =   825
      Width           =   1035
   End
   Begin VB.TextBox Logradouro 
      Height          =   300
      Left            =   1290
      MaxLength       =   60
      TabIndex        =   4
      Top             =   270
      Width           =   4950
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6465
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "EnderecoEntrega.ctx":00C5
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "EnderecoEntrega.ctx":0243
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   16
      Top             =   3585
      Width           =   1110
   End
   Begin VB.Label Label6 
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   870
      TabIndex        =   12
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label5 
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   405
      TabIndex        =   11
      Top             =   3090
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   615
      TabIndex        =   9
      Top             =   1950
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Complemento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   30
      TabIndex        =   7
      Top             =   1410
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   465
      TabIndex        =   5
      Top             =   870
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Logradouro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   3
      Top             =   300
      Width           =   1110
   End
End
Attribute VB_Name = "EnderecoEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Private sUFCidades As String

Public Sub Form_Unload(Cancel As Integer)

    Set gobjVenda = Nothing

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    With gobjVenda.objCupomFiscal
        .sEndEntLogradouro = Trim(Logradouro.Text)
        .sEndEntNúmero = Trim(Numero.Text)
        .sEndEntComplemento = Trim(Complemento.Text)
        .sEndEntBairro = Trim(Bairro.Text)
        .sEndEntCidade = Trim(ComboCidade.Text)
        .sEndEntUF = Trim(UF.Text)
        .lEndEntIBGECidade = ComboCidade.ItemData(ComboCidade.ListIndex)
        .sEndEntEmail = Trim(Email.Text)
    End With
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144091)
    
    End Select

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144093)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144094)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(ByVal objVenda As ClassVenda) As Long

Dim bCancela As Boolean

On Error GoTo Erro_Trata_Parametros

    With objVenda.objCupomFiscal
        Logradouro.Text = .sEndEntLogradouro
        Numero.Text = .sEndEntNúmero
        Complemento.Text = .sEndEntComplemento
        Bairro.Text = .sEndEntBairro
        
        UF.Text = IIf(Len(Trim(.sEndEntUF)) = 0, gsUF, .sEndEntUF)
        Call UF_Validate(bCancela)
        Call List_Item_Igual(ComboCidade, IIf(Len(Trim(.sEndEntCidade)) = 0, gsCidade, .sEndEntCidade))
        Email.Text = .sEndEntEmail
    End With

    Set gobjVenda = objVenda

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144095)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Endereço de Entrega"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "EnderecoEntrega"

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

Private Sub Email_Validate(Cancel As Boolean)
    
Dim sEmail As String

On Error GoTo Erro_Email_Validate

    sEmail = Trim(Email.Text)

    If Len(sEmail) <> 0 Then
        If Not ValidEmail(sEmail) Then gError 201581
    End If
    
    Exit Sub

Erro_Email_Validate:

    Cancel = True

    Select Case gErr

        Case 201581
            Call Rotina_ErroECF(vbOKOnly, ERRO_EMAIL_INVALIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 201580)

    End Select

    Exit Sub

End Sub

Private Sub UF_Validate(Cancel As Boolean)
    If UF.Text <> sUFCidades Then
    
        ComboCidade.Clear
        
        Call CF_ECF("CidadesUF_CarregaCombo", ComboCidade, UF.Text)
    
        ComboCidade.ListIndex = -1
    
        sUFCidades = UF.Text
        
    End If
    
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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****

Public Function objParent() As Object

    Set objParent = Parent
    
End Function


