VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl RegPAFECFParcial 
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   8655
   Begin VB.Frame FrameDatas 
      Caption         =   "Intervalo de Datas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   1035
      TabIndex        =   11
      Top             =   3690
      Width           =   4380
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   435
         Left            =   1935
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   420
         Left            =   585
         TabIndex        =   13
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   435
         Left            =   4080
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   180
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   420
         Left            =   2790
         TabIndex        =   15
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2265
         TabIndex        =   17
         Top             =   540
         Width           =   510
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   525
         Width           =   435
      End
   End
   Begin VB.CommandButton BotaoIncluir 
      Caption         =   "(F4) Adicionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6975
      TabIndex        =   9
      Top             =   1860
      Width           =   1560
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6885
      ScaleHeight     =   495
      ScaleWidth      =   1650
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   225
      Width           =   1710
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "RegPAFECFParcial.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "F5 - Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   630
         Picture         =   "RegPAFECFParcial.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "RegPAFECFParcial.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
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
      Left            =   5040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   315
      Width           =   1500
   End
   Begin VB.CommandButton BotaoExcluir 
      Caption         =   "(F5) Remover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6975
      TabIndex        =   4
      Top             =   2310
      Width           =   1560
   End
   Begin VB.ListBox ListaProduto 
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   6675
   End
   Begin MSMask.MaskEdBox ProdutoNomeRed 
      Height          =   375
      Left            =   1110
      TabIndex        =   0
      Top             =   300
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Produtos Selecionados"
      Height          =   210
      Left            =   165
      TabIndex        =   3
      Top             =   1020
      Width           =   5115
   End
   Begin VB.Label LabelProduto 
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
      Left            =   165
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "RegPAFECFParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim giProdutoAlterado As Integer

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
    Exit Function
     
End Function

Private Sub BotaoExcluir_Click()

On Error GoTo Erro_BotaoExcluir_Click

    If ListaProduto.ListIndex = -1 Then gError 210179
    
    ListaProduto.RemoveItem ListaProduto.ListIndex
    ProdutoNomeRed.Text = ""
    
    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 210179
            Call Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_SELECIONADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 210180)

    End Select

    Exit Sub
    
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim sProduto As String
Dim colProdutos As New Collection
Dim iPos As Integer
Dim lErro As Long
Dim iIndice As Integer
Dim sNomeRedProd As String
Dim objProduto As ClassProduto
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_BotaoGravar_Click

    For iIndice = 0 To ListaProduto.ListCount - 1
        
        sNomeRedProd = ListaProduto.List(iIndice)
    
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sNomeRedProd, objProduto)
    
        colProdutos.Add objProduto.sCodigo
    
    Next

    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 214079
    
    If Len(Trim(DataDe.ClipText)) > 0 Then dtDataDe = CDate(DataDe.Text)
    If Len(Trim(DataAte.ClipText)) > 0 Then dtDataAte = CDate(DataAte.Text)

    If dtDataDe > dtDataAte Then gError 214080

    lErro = CF_ECF("Relatorio_Registros_PAFECF", dtDataDe, dtDataAte, colProdutos)
    If lErro <> SUCESSO Then gError 214081





'    lErro = CF_ECF("EstProdParcial_Grava", colProdutos)
'    If lErro <> SUCESSO Then gError 210177
    
    Call Limpa_Tela(Me)
    ListaProduto.Clear
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr

'        Case 210177

        Case 214079
            Call Rotina_ErroECF(vbOKOnly, ERRO_INTERVALO_NAO_PREENCHIDO, gErr)

        Case 214080
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 214081
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 210178)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoIncluir_Click()

Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_BotaoIncluir_Click

    If Len(Trim(ProdutoNomeRed.Text)) > 0 Then

        Call ProdutoNomeRed_Validate(bCancel)

        If Not bCancel Then

            For iIndice = 0 To ListaProduto.ListCount - 1
                
                If ListaProduto.List(iIndice) = ProdutoNomeRed.Text Then gError 210801
            
            Next
        
        
            ListaProduto.AddItem ProdutoNomeRed.Text
            ProdutoNomeRed.Text = ""
    
        End If
    
    End If
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 210801
            Call Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_JA_SELECIONADO_LOJA, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 210800)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Chama a função que limpa a tela

    'Chama Limpa_Tela
    Call Limpa_Tela(Me)

End Sub

Public Sub BotaoProdutos_Click()
'Chama o browser do ProdutoLojaLista
'So traz produtos onde codigo de barras ou referencia está preenchida

Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoProdutos_Click
    
    objProduto.sNomeReduzido = ProdutoNomeRed.Text
    
    Call Chama_TelaECF_Modal("ProdutosLista", objProduto)
        
    If giRetornoTela = vbOK Then
        ProdutoNomeRed.Text = objProduto.sNomeReduzido
    End If
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 210183)

    End Select

    Exit Sub

End Sub


Private Sub Produto_Change()

    giProdutoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO
    
    giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)

'    Set gobjProduto = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Registros de PAF_ECF Estoque Parcial"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RegPAFECFParcial"

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

Private Sub ProdutoNomeRed_Validate(Cancel As Boolean)
    
Dim sProduto1 As String
Dim objProduto As ClassProduto
    
On Error GoTo Erro_ProdutoNomeRed_Validate
    
    If Len(Trim(ProdutoNomeRed.Text)) > 0 Then
    
        sProduto1 = ProdutoNomeRed.Text
    
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto1, objProduto)
            
        'caso o produto não seja encontrado
        If objProduto Is Nothing Then gError 210798

        ProdutoNomeRed.Text = objProduto.sNomeReduzido

    End If

    Exit Sub

Erro_ProdutoNomeRed_Validate:

    Cancel = True

    Parent.MousePointer = vbDefault

    Select Case gErr
                
        Case 210798
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210799)

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
    
On Error GoTo Erro_UserControl_KeyDown
    
    Select Case KeyCode
    
        Case vbKeyF3
            If Not TrocaFoco(Me, BotaoProdutos) Then Exit Sub
            Call BotaoProdutos_Click
    
        Case vbKeyF4
            If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
            Call BotaoIncluir_Click
            
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210181)

    End Select

    Exit Sub
    
End Sub


Private Sub DataDe_GotFocus()
    'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataDe_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDe)

    Exit Sub

Erro_DataDe_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204584)

    End Select

    Exit Sub


End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDe_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 204585

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 204585
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204586)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataAte_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAte)

    Exit Sub

Erro_DataAte_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204587)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAte_Validate

    'Verifica se Data Até esta Preenchida se não sai do Validate
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 204588

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 204588
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204589)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204776
    
    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 204776

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204777)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204778

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 204778

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204779)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204780
    
    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 204780

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204781)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204782

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 204782

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204783)

    End Select

    Exit Sub

End Sub

