VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl NFEntProduto 
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   4530
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   2616
      Picture         =   "NFEntProduto.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   912
      Picture         =   "NFEntProduto.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox CodigoBarras 
      Height          =   288
      Left            =   1764
      TabIndex        =   4
      Top             =   1536
      Width           =   2616
   End
   Begin VB.TextBox Referencia 
      Height          =   312
      Left            =   1764
      TabIndex        =   2
      Top             =   900
      Width           =   1635
   End
   Begin VB.ComboBox Fabricante 
      Height          =   288
      Left            =   1764
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Preco 
      Height          =   312
      Left            =   1764
      TabIndex        =   6
      Top             =   2124
      Width           =   1692
      _ExtentX        =   2963
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Preço:"
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
      Left            =   1140
      TabIndex        =   7
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código de Barras:"
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
      Left            =   168
      TabIndex        =   5
      Top             =   1596
      Width           =   1512
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Referência:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   948
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante:"
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
      Left            =   732
      TabIndex        =   1
      Top             =   408
      Width           =   948
   End
End
Attribute VB_Name = "NFEntProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjProduto As ClassProduto
Dim glErro As Long

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)


End Sub

Private Sub BotaoCancela_Click()
    gobjProduto.sCodigoBarras = "Cancel"
    Unload Me
End Sub

Public Sub Form_Load()
    
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim colItensCategoria As New Collection
Dim lErro As Long
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

    
On Error GoTo Erro_Form_Load
    
    
    objCategoriaProduto.sCategoria = "FabricanteRelogio"
    
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 22541 Then gError 199438
    
    For Each objCategoriaProdutoItem In colItensCategoria
    
        Fabricante.AddItem objCategoriaProdutoItem.sItem
    
    Next
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 199438
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 199439)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(objProduto As ClassProduto) As Long

    Set gobjProduto = objProduto

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(Fabricante.Text)) = 0 Then gError 199440
    
    If Len(Trim(Referencia.Text)) = 0 Then gError 199441
    
    If Len(Trim(CodigoBarras.Text)) = 0 Then gError 199442
    
    'Verifica preenchimento de valor
    If Len(Trim(Preco.Text)) = 0 Then gError 199443

    gobjProduto.sCodigo = Fabricante.Text
    gobjProduto.sReferencia = Referencia.Text
    gobjProduto.dPrecoLoja = StrParaDbl(Preco.ClipText)
    gobjProduto.sCodigoBarras = CodigoBarras.Text

    gobjProduto.lErro = SUCESSO
    
    Unload Me

    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 199440
            Call Rotina_Erro(vbOKOnly, "ERRO_FABRICANTE_NAO_PREENCHIDO", gErr)

        Case 199441
            Call Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_NAO_PREENCHIDA", gErr)

        Case 199442
            Call Rotina_Erro(vbOKOnly, "ERRO_CODBARRA_NAO_PREENCHIDO", gErr)

        Case 199443
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199444)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID =
    Set Form_Load_Ocx = Me
    Caption = "Relógios"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFEntProduto"
    
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

Public Sub Preco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Preco_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(Preco.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Preco.Text)
    If lErro <> SUCESSO Then gError 199445

    Preco.Text = Format(Preco.Text, "Fixed")

    Exit Sub

Erro_Preco_Validate:

    Cancel = True

    Select Case gErr

        Case 199445

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199446)

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



