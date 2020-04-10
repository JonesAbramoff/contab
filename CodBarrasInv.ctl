VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CodBarrasInv 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4185
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   4185
   Begin VB.CheckBox Fixar 
      Caption         =   "Fixar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   3108
      TabIndex        =   6
      Top             =   768
      Width           =   912
   End
   Begin VB.TextBox Codigo 
      Height          =   288
      Left            =   1260
      TabIndex        =   0
      Top             =   204
      Width           =   2616
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   2040
      Picture         =   "CodBarrasInv.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1272
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   864
      Picture         =   "CodBarrasInv.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1272
      Width           =   855
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   288
      Left            =   1248
      TabIndex        =   1
      Top             =   684
      Width           =   1548
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   132
      TabIndex        =   5
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cod. Barras:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   132
      TabIndex        =   4
      Top             =   240
      Width           =   1044
   End
End
Attribute VB_Name = "CodBarrasInv"
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

Private Sub BotaoCancela_Click()
    gobjProduto.sCodigoBarras = "Cancel"
    Unload Me
End Sub

Public Sub Form_Load()
    
On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 199522)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(objProduto As ClassProduto) As Long

    Set gobjProduto = objProduto

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoOK_Click()

Dim objProduto As New ClassProduto
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    If Len(Codigo.Text) = 12 Then
    
        Codigo.Text = "4" & Right(Codigo.Text, 11)
    
        lErro = CF("ProdutoCodBarras_Le12", Codigo.Text, gobjProduto)
        If lErro <> SUCESSO And lErro <> 199873 Then gError 199525
        
        If lErro = SUCESSO Then
            Codigo.Text = gobjProduto.sCodigoBarras
        End If
    Else

        lErro = CF("ProdutoCodBarras_Le", Codigo.Text, gobjProduto)
        If lErro <> SUCESSO And lErro <> 193965 Then gError 199525
    
    End If
    
    gobjProduto.sCodigoBarras = Codigo.Text
    gobjProduto.dPesoEspecifico = StrParaDbl(Quantidade.Text)
    
    'se o codigo de barras nao existir
    If lErro <> SUCESSO Then gError 199526
    
    gobjProduto.lErro = SUCESSO
    
    Unload Me

    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 199525

        Case 199526
            If gobjProduto.lErro = 0 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_BARRAS_NAO_ENCONTRADO", gErr, Codigo.Text)
            Else
                gobjProduto.lErro = gErr
                Unload Me
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199527)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANAIS_VENDA
    Set Form_Load_Ocx = Me
    Caption = "Código de Barras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CodBarrasInv"
    
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

Private Sub Quantidade_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Quantidade_validate

    If Len(Trim(Quantidade.ClipText)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 199523

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If
    
    Exit Sub

Erro_Quantidade_validate:

    Select Case gErr

        Case 199523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199524)

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



