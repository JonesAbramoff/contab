VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl Preco 
   Appearance      =   0  'Flat
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   5745
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4305
      ScaleHeight     =   495
      ScaleWidth      =   1230
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1290
      Begin VB.CommandButton BotaoFechar 
         Cancel          =   -1  'True
         Height          =   360
         Left            =   690
         Picture         =   "Preco.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   195
         Picture         =   "Preco.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProdutos 
      Caption         =   "(F3) Produtos"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   1500
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   360
      Left            =   825
      TabIndex        =   1
      Top             =   210
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Image Figura 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   150
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   45
      TabIndex        =   0
      Top             =   285
      Width           =   735
   End
   Begin VB.Label UMVenda 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2610
      TabIndex        =   3
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Preco 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label4 
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
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   930
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "UM:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   930
      Width           =   390
   End
End
Attribute VB_Name = "Preco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******** IMPORTANTE ************************************
'Ao incluir algum campo novo nesta tela,  verificar a necessidade de alterar a rotina UserControl_KeyDown.

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjProduto As ClassProduto

Dim iAlterado As Integer

Dim giProdutoAlterado As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

    'Se o obj estiver carregado
    If Not (objProduto Is Nothing) Then

        'Coloca a variável global como igual ao obj recebido
        Set gobjProduto = objProduto
        
        'Se a descrição do produto estiver preenchida
        If Len(Trim(objProduto.sReferencia)) <> 0 Then
            Produto.Text = objProduto.sReferencia
        Else
            Produto.Text = objProduto.sCodigoBarras
        End If
        
        Call Produto_Validate(False)
            
    Else
        
        'Se o obj não estiver carregado, então apenas inicializa o obj global
        Set gobjProduto = New ClassProduto
    
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function
     
End Function

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
'Chama a função que limpa a tela

    Call Limpa_Tela_Precos

End Sub

Sub Limpa_Tela_Precos()

    'Chama Limpa_Tela
    Call Limpa_Tela(Me)
    
    'Chama a função que irá limpar o restante dos campos da tela
    Call Limpa_Tela_Precos2
    
End Sub

Sub Limpa_Tela_Precos2()
    
    'Limpa os Campos do tipo label presentes na tela
    Preco.Caption = ""
    UMVenda.Caption = ""
    
End Sub

Public Sub BotaoProdutos_Click()
'Chama o browser do ProdutoLojaLista
'So traz produtos onde codigo de barras ou referencia está preenchida

Dim objProduto As New ClassProduto
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click
    
'    objProduto.sNomeReduzido = Produto.Text
'
'    Call Chama_TelaECF_Modal("ProdutosLista", objProduto)
'
'    If giRetornoTela = vbOK Then
'        If objProduto.sReferencia <> "" Then
'            Produto.Text = objProduto.sReferencia
'        Else
'            Produto.Text = objProduto.sCodigoBarras
'        End If
'        Call Produto_Validate(False)
'    End If

    objProduto.sNomeReduzido = Produto.Text

    Call Chama_TelaECF_Modal("ProdutosLista", colSelecao, objProduto, objEventoProduto)

    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 165100)

    End Select

    Exit Sub

End Sub


Private Sub Produto_Change()

    giProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim objProduto As ClassProduto
Dim sRet As String
Dim sProduto As String
Dim objProdNomeRed As TextBox

On Error GoTo Erro_Produto_Validate

    'Verifica se o campo código foi preenchido
    If Len(Trim(Produto.Text)) = 0 Then
         Preco.Caption = ""
         UMVenda.Caption = ""
    Else
    
        sProduto = Produto.Text
    
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
        If Not (objProduto Is Nothing) Then
            Produto.Text = objProduto.sNomeReduzido
            Preco.Caption = Format(objProduto.dPrecoLoja, "standard")
            UMVenda.Caption = objProduto.sSiglaUMVenda
        
            'verifica se a figura foi preenchida
            If objProduto.sFigura <> "" Then
                'verifica se o arquivo é do tipo imagem
                sRet = Dir(objProduto.sFigura, vbNormal)
                If sRet <> "" Then
                    If GetAttr(objProduto.sFigura) = vbArchive Or GetAttr(objProduto.sFigura) = vbArchive + vbReadOnly Then
                        'coloca a figura na tela
                        Figura.Picture = LoadPicture(objProduto.sFigura)
                    End If
                Else
                    gError 99607
                End If
            Else
                Figura.Picture = LoadPicture
            End If
        
        End If
    End If
    
    Exit Sub
    
Erro_Produto_Validate:
    
    Cancel = True

    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 165101)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

    Set objEventoProduto = New AdmEvento

    lErro_Chama_Tela = SUCESSO
    
    giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjProduto = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Preço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Preco"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
    
On Error GoTo Erro_UserControl_KeyDown
    
    Select Case KeyCode
    
        Case vbKeyF3
            Call BotaoProdutos_Click
    
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 165102)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(objProduto.sReferencia)) > 0 Then
        Produto.Text = objProduto.sReferencia
    Else
        Produto.Text = objProduto.sCodigoBarras
    End If
    Call Produto_Validate(False)

'    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 214935)

    End Select

    Exit Sub

End Sub

