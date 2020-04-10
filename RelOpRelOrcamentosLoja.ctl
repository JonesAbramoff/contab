VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelOrcamentosLoja 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   ScaleHeight     =   4545
   ScaleWidth      =   6690
   Begin VB.CheckBox ExibirItens 
      Caption         =   "Exibir Itens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame FrameVendedor 
      Caption         =   "Vendedor"
      Height          =   735
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   4215
      Begin MSMask.MaskEdBox VendedorDe 
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorAte 
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelVendedorDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameOrcamento 
      Caption         =   "Orçamento"
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox OrcamentoDe 
         Height          =   315
         Left            =   810
         TabIndex        =   3
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OrcamentoAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   4
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelOrcamentoDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   360
         Width           =   315
      End
      Begin VB.Label LabelOrcamentoAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2280
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   4215
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   2560
         TabIndex        =   6
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelCaixaAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelCaixaDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Height          =   195
         Left            =   2265
         TabIndex        =   21
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   345
         Width           =   315
      End
   End
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
      Left            =   4733
      Picture         =   "RelOpRelOrcamentosLoja.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelOrcamentosLoja.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelOrcamentosLoja.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelOrcamentosLoja.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpRelOrcamentosLoja.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelOrcamentosLoja.ctx":0A96
      Left            =   1080
      List            =   "RelOpRelOrcamentosLoja.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2670
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelOrcamentosLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type typeCampo
    sVend_De As String
    sVend_Ate As String
    sCaixa_De As String
    sCaixa_Ate As String
    sOrcamento_De As String
    sOrcamento_Ate As String
    sData_De As String
    sData_Ate As String
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giVendedorDe As Integer
Dim giOrcamentoDe As Integer
Dim giCaixaDe As Integer

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF '???
    Set Form_Load_Ocx = Me
    Caption = "Relação de Orçamentos emitidos em ECF"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelOrcamentosLoja"
    
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

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelOrcamentoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrcamentoDe, Source, X, Y)
End Sub

Private Sub LabelOrcamentoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrcamentoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

Private Sub labelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub
'
'Private Sub LabelCaixaAte_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
'End Sub

Private Sub LabelCaixaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaAte, Button, Shift, X, Y)
End Sub
'
'Private Sub LabelCaixaDe_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
'End Sub

Private Sub LabelCaixaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoCaixa = Nothing
    Set objEventoOrcamento = Nothing
    Set objEventoVendedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
End Sub

Private Sub LabelVendedorAte_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

On Error GoTo Erro_LabelVendedorAte_Click
    
    giVendedorDe = 0
    
    'Verifica se Vendedor está preenchido
    If Len(Trim(VendedorAte.ClipText)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorAte.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)
    
    Exit Sub

Erro_LabelVendedorAte_Click:
  
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172605)

    End Select

    Exit Sub

End Sub

Private Sub LabelVendedorDe_Click()
'Chama browser Vendedor

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

On Error GoTo Erro_LabelVendedorDe_Click
    
    'Determina o label que está sendo clicado
    giVendedorDe = 1
    
    If Len(Trim(VendedorDe.ClipText)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorDe.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)
    
    Exit Sub

Erro_LabelVendedorDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172606)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaDe_Click()
'Chama browser de Caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaDe_Click
    
    'Determina a label que está sendo clicado
    giCaixaDe = 1

    'Verifica se o caixa está preenchido
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com o Caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If

    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    End If
    
    Exit Sub

Erro_LabelCaixaDe_Click:

    LabelCaixaDe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172607)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCaixaAte_Click()
'Chama browser de Caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaAte_Click
    
    'Determina qual label está sendo clicada
    giCaixaDe = 0
    
    'Verifica se o Caixa final está preenchido
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com o Caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If

    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)

    End If
    
    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172608)

    End Select

    Exit Sub

End Sub

Private Sub LabelOrcamentoDe_Click()
'Chama Browser Orcamento

Dim objOrcamento As New ClassCupomFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LabelOrcamentoDe_Click

    'Determina qual label está sendo clicada
    giOrcamentoDe = 1
    
    'Verifica se Orcamento inicial está preenchido
    If Len(Trim(OrcamentoDe.ClipText)) > 0 Then
        'Preenche com o Orcamento da tela
        objOrcamento.lNumOrcamento = Codigo_Extrai(OrcamentoDe.Text)
    End If

    'Chama Tela OrcamentoLista
    Call Chama_Tela("OrcamentoLojaLista", colSelecao, objOrcamento, objEventoOrcamento)

    Exit Sub

Erro_LabelOrcamentoDe_Click:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172609)

    End Select

    Exit Sub

End Sub

Private Sub LabelOrcamentoAte_Click()

Dim objOrcamento As New ClassCupomFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LabelOrcamentoAte_click
    
    'Determina qual label está sendo clicada
    giOrcamentoDe = 0

    'Se o campo OrcamentoAte estiver preenchido...
    If Len(Trim(OrcamentoAte.ClipText)) > 0 Then
        'Preenche com o Orcamento da tela
        objOrcamento.lNumOrcamento = OrcamentoAte.Text
    End If

    'Chama Tela OrcamentoLista
    Call Chama_Tela("OrcamentoLojaLista", colSelecao, objOrcamento, objEventoOrcamento)
  
    Exit Sub

Erro_LabelOrcamentoAte_click:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172610)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao
    
    'Seta o objeto
    Set objCaixa = obj1

    'Preenche campo Caixa
    If giCaixaDe = 1 Then
        CaixaDe.PromptInclude = False
        CaixaDe.Text = CStr(objCaixa.iCodigo)
        CaixaDe.PromptInclude = True
        CaixaDe_Validate (bSGECancelDummy)
    Else
        CaixaAte.PromptInclude = False
        CaixaAte.Text = CStr(objCaixa.iCodigo)
        CaixaAte.PromptInclude = True
        CaixaAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172611)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    Set objEventoVendedor = New AdmEvento
    Set objEventoOrcamento = New AdmEvento
    Set objEventoCaixa = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172612)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 117001
        
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 117000
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 117000
        
        Case 117001
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172613)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()
'Executa relatorio

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 117002
        
    'Se ExibirItens estiver selecionado então.....
    If ExibirItens.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "REORITLJ"
    
    'Se não...
    Else
        gobjRelatorio.sNomeTsk = "RELORLJ"
    End If
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 117002

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172614)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()
    
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
' valida e critica a data final

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Se Data Final estiver preenchida...
    If Len(DataAte.ClipText) > 0 Then
         
        'Critica se é uma data válida....
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 117003

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 117003

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172615)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'valida e critica a data Inicial
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a data inicial estiver preenchida...
    If Len(DataDe.ClipText) > 0 Then
                
        'Critica se é uma data válida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 117004

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 117004

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172616)

    End Select

    Exit Sub

End Sub

Private Sub VendedorAte_LostFocus()
    
    Call VendedorAte_Validate(bSGECancelDummy)

End Sub

Private Sub VendedorDe_LostFocus()
 
    Call VendedorDe_Validate(bSGECancelDummy)

End Sub

Private Sub VendedorDe_Validate(Cancel As Boolean)
'Verifica se existe o vendedor no bd

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorDe_Validate

    'Se o vendedor inicial estiver preenchido
    If Len(Trim(VendedorDe.ClipText)) > 0 Then
        'Tenta ler o codigo do vendedor
        lErro = TP_Vendedor_Le2(VendedorDe, objVendedor, 0)
        If lErro <> SUCESSO Then gError 117005

    End If
    
    giVendedorDe = 1
    
    Exit Sub

Erro_VendedorDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 117005
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172617)

    End Select

End Sub

Private Sub VendedorAte_Validate(Cancel As Boolean)
'Verifica se existe o vendedor no bd

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorAte_Validate

    If Len(Trim(VendedorAte.ClipText)) > 0 Then
        'Tenta ler o Código do Vendedor
        lErro = TP_Vendedor_Le2(VendedorAte, objVendedor, 0)
        If lErro <> SUCESSO Then gError 117224

    End If
    
    giVendedorDe = 0
 
    Exit Sub

Erro_VendedorAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 117224
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172618)

    End Select

End Sub

Private Sub CaixaAte_LostFocus()
    Call CaixaAte_Validate(bSGECancelDummy)
End Sub

Private Sub CaixaDe_LostFocus()
    Call CaixaDe_Validate(bSGECancelDummy)
End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaDe_Validate

    'Se o Caixa Inicial estiver preenchido...
    If Len(Trim(CaixaDe.Text)) > 0 Then

        'Preenche o objeto com o código da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler o codigo do vendedor
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 Then gError 117212

    End If

    giCaixaDe = 1

    Exit Sub

Erro_CaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 117212

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172619)

    End Select

End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'Valida se o Caixa existe no bd

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    If Len(Trim(CaixaAte.Text)) > 0 Then
    
        'Preenche o objeto com os dados da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa

        'Tenta ler o Código do Vendedor
        lErro = CF("TP_Caixa_Le1", CaixaAte, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 Then gError 117006

    End If

    giCaixaDe = 0

    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 117006
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172620)

    End Select

End Sub

Private Sub OrcamentoAte_LostFocus()
    Call OrcamentoAte_Validate(bSGECancelDummy)
End Sub

Private Sub OrcamentoDe_LostFocus()
    Call OrcamentoDe_Validate(bSGECancelDummy)
End Sub

Private Sub OrcamentoDe_Validate(Cancel As Boolean)
'Valida o campo de OrcamentoDe

Dim lErro As Long

On Error GoTo Erro_OrcamentoDe_Validate
     
    'Se o orçamento inicial estiver preenchido
    If Len(Trim(OrcamentoDe.ClipText)) > 0 Then
        'Critica se é um código válido
        lErro = Long_Critica(OrcamentoDe.Text)
        If lErro <> SUCESSO Then Error 117212
        
    End If
        
    Exit Sub

Erro_OrcamentoDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 117212
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172621)
            
    End Select
    
    Exit Sub

End Sub

Private Sub OrcamentoAte_Validate(Cancel As Boolean)
'Valida o campo OrcamentoAte

Dim lErro As Long

On Error GoTo Erro_OrcamentoAte_Validate
     
    'Se o orçamento final estiver preenchido
    If Len(Trim(OrcamentoAte.ClipText)) > 0 Then

        'Critica se o código de orcamento é um número válido
        lErro = Long_Critica(OrcamentoAte.Text)
        If lErro <> SUCESSO Then Error 117213
        
    End If
    
    Exit Sub

Erro_OrcamentoAte_Validate:

    Cancel = True

    Select Case Err
    
        Case 117213
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172622)
            
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 117007

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 117007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172623)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aumente a data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 117008

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 117008

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172624)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 117009

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 117009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172625)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aumenta a data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 117010

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 117010

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172626)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao
    
    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorDe = 1 Then
        VendedorDe.Text = CStr(objVendedor.iCodigo)
        VendedorDe_Validate (bSGECancelDummy)
    Else
        VendedorAte.Text = CStr(objVendedor.iCodigo)
        VendedorAte_Validate (bSGECancelDummy)
    End If

    Me.Show
 
 Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 117010

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172627)

    End Select

    Exit Sub

    

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim tCampo As typeCampo

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(tCampo)
    If lErro <> SUCESSO Then gError 117011
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 117012
    
    lErro = objRelOpcoes.IncluirParametro("NCAIXADE", tCampo.sCaixa_De)
    If lErro <> AD_BOOL_TRUE Then gError 117013
    
    lErro = objRelOpcoes.IncluirParametro("TCAIXADE", Trim(CaixaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117212

    lErro = objRelOpcoes.IncluirParametro("NCAIXAATE", tCampo.sCaixa_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117014
    
    lErro = objRelOpcoes.IncluirParametro("TCAIXAATE", Trim(CaixaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117213
    
    lErro = objRelOpcoes.IncluirParametro("NORCAMENTODE", tCampo.sOrcamento_De)
    If lErro <> AD_BOOL_TRUE Then gError 117015
    
    lErro = objRelOpcoes.IncluirParametro("NORCAMENTOATE", Trim(OrcamentoAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117215
         
    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", tCampo.sVend_De)
    If lErro <> AD_BOOL_TRUE Then gError 117017
    
    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", Trim(VendedorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117018

    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", tCampo.sVend_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117019
    
    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", Trim(VendedorAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117020
           
    lErro = objRelOpcoes.IncluirParametro("DINIC", tCampo.sData_De)
    If lErro <> AD_BOOL_TRUE Then gError 117021
    
    lErro = objRelOpcoes.IncluirParametro("DFIM", tCampo.sData_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117022
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, tCampo)
    If lErro <> SUCESSO Then gError 117023
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 117011 To 117023
        
        Case 117213 To 117215
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172628)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(tCampo As typeCampo) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    '*** critica vendedor Inicial e Final ***
    If VendedorDe.ClipText <> "" Then
        tCampo.sVend_De = CStr(Codigo_Extrai(VendedorDe.Text))
    Else
        tCampo.sVend_De = ""
    End If
    
    If VendedorAte.ClipText <> "" Then
        tCampo.sVend_Ate = CStr(Codigo_Extrai(VendedorAte.Text))
    Else
        tCampo.sVend_Ate = ""
    End If
            
    If tCampo.sVend_De <> "" And tCampo.sVend_Ate <> "" Then
        
        'Se o Vendedor Inicial for maior que o final --> erro
        If CInt(tCampo.sVend_De) > CInt(tCampo.sVend_Ate) Then gError 117024
        
    End If
    '****************************************
     
    '*** critica Caixa Inicial e Final ***
    If CaixaDe.ClipText <> "" Then
        tCampo.sCaixa_De = CStr(Codigo_Extrai(CaixaDe.Text))
    Else
        tCampo.sCaixa_De = ""
    End If
    
    If CaixaAte.ClipText <> "" Then
        tCampo.sCaixa_Ate = CStr(Codigo_Extrai(CaixaAte.Text))
    Else
        tCampo.sCaixa_Ate = ""
    End If
            
    If tCampo.sCaixa_De <> "" And tCampo.sCaixa_Ate <> "" Then
        
        'Se o Caixa Inicial for maior que o final --> erro
        If CInt(tCampo.sCaixa_De) > CInt(tCampo.sCaixa_Ate) Then gError 117025
        
    End If
    '*****************************************
    
    '*** critica Orcamento Inicial e Final ***
    If OrcamentoDe.ClipText <> "" Then
        tCampo.sOrcamento_De = CStr(OrcamentoDe.ClipText)
    Else
        tCampo.sOrcamento_De = ""
    End If
    
    If OrcamentoAte.ClipText <> "" Then
        tCampo.sOrcamento_Ate = CStr(OrcamentoAte.ClipText)
    Else
        tCampo.sOrcamento_Ate = ""
    End If
            
    If tCampo.sOrcamento_De <> "" And tCampo.sOrcamento_Ate <> "" Then
        
        'Se o Orçamento Inicial for maior que o final --> erro
        If CInt(tCampo.sOrcamento_De) > CInt(tCampo.sOrcamento_Ate) Then gError 117026
        
    End If
    '**********************************************
    
    '*** Data ***
    If Trim(DataDe.ClipText) <> "" Then
        tCampo.sData_De = CStr(DataDe.Text)
    Else
        tCampo.sData_De = DATA_NULA
    End If
    
    If Trim(DataAte.ClipText) <> "" Then
        tCampo.sData_Ate = CStr(DataAte.Text)
    Else
        tCampo.sData_Ate = DATA_NULA
    End If
   
    'data inicial não pode ser maior que a data final
    If Trim(tCampo.sData_De) <> DATA_NULA And Trim(tCampo.sData_Ate) <> DATA_NULA Then
    
         If CDate(tCampo.sData_De) > CDate(tCampo.sData_Ate) Then gError 117035
    
    End If
    '***************************************************
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 117035
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
       
        Case 117024
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
        
        Case 117026
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_INICIAL_MAIOR", gErr)
                                
        Case 117025
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INICIAL_MAIOR", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172629)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, tCampo As typeCampo) As Long
'Monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   'Se o campo vendedorDe for diferente de vazio.....
   If tCampo.sVend_De <> "" Then sExpressao = "Vendedor >= " & Forprint_ConvInt(CInt(tCampo.sVend_De))

   'Se o campo vendedorAte for diferente de vazio.....
   If tCampo.sVend_Ate <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(tCampo.sVend_Ate))

    End If
      
   'Se o campo OrcamentoDe for diferente de vazio.....
   If tCampo.sOrcamento_De <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Orcamento >= " & Forprint_ConvInt(CInt(tCampo.sOrcamento_De))
        
   End If

   'Se o campo OrcamentoAte for diferente de vazio.....
   If tCampo.sOrcamento_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Orcamento <= " & Forprint_ConvInt(CInt(tCampo.sOrcamento_Ate))

    End If
    
   'Se o campo CaixaDe for diferente de vazio.....
   If tCampo.sCaixa_De <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(CInt(tCampo.sCaixa_De))
        
   End If

   'Se o campo CaixaAte for diferente de vazio.....
   If tCampo.sCaixa_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(CInt(tCampo.sCaixa_Ate))

    End If
      
    'Se o campo DataDe for diferente de data nula.....
    If tCampo.sData_De <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If
    
    'Se o campo DataAte for diferente de data nula.....
    If tCampo.sData_Ate <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    
    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172630)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 117029

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 117028

        'Retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
           
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 117029
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)

        Case 117028

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172631)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 117030
    
    'Limpa  a combo opções
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 117030
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172632)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is VendedorDe Then
            Call LabelVendedorDe_Click
            
        ElseIf Me.ActiveControl Is VendedorAte Then
            Call LabelVendedorAte_Click
            
        ElseIf Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
            
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
            
        End If
    
    End If

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Nome da opção de Relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 117031

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 117032

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 117033

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 117034
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 117031
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)

        Case 117032, 117033, 117034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172633)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 117036
    
    'Exibe vendedor inicial
    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
    If lErro Then gError 117037

    'O campo da tela recebe o parametro
    VendedorDe.Text = sParam
    Call VendedorDe_Validate(bSGECancelDummy)
    
    'Exibe vendedor final
    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
    If lErro Then gError 117038
    
    'O campo da tela recebe o parametro
    VendedorAte.Text = sParam
    Call VendedorAte_Validate(bSGECancelDummy)
    
    'Exibe Caixa inicial
    lErro = objRelOpcoes.ObterParametro("NCAIXADE", sParam)
    If lErro Then gError 117039
    
    CaixaDe.PromptInclude = False
    CaixaDe.Text = sParam
    CaixaDe.PromptInclude = True
    Call CaixaDe_Validate(bSGECancelDummy)
    
    'Exibe Caixa final
    lErro = objRelOpcoes.ObterParametro("NCAIXAATE", sParam)
    If lErro Then gError 117040
    
    CaixaAte.PromptInclude = False
    CaixaAte.Text = sParam
    CaixaAte.PromptInclude = True
    Call CaixaAte_Validate(bSGECancelDummy)
    
    'Exibe Orcamento inicial
    lErro = objRelOpcoes.ObterParametro("NORCAMENTODE", sParam)
    If lErro Then gError 117041
    
    OrcamentoDe.PromptInclude = False
    OrcamentoDe.Text = sParam
    OrcamentoDe.PromptInclude = True
    
    'Exibe Orcamento final
    lErro = objRelOpcoes.ObterParametro("NORCAMENTOATE", sParam)
    If lErro Then gError 117042
    
    OrcamentoAte.PromptInclude = False
    OrcamentoAte.Text = sParam
    OrcamentoAte.PromptInclude = True
    
    'Exibe data inicial
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 117043

    Call DateParaMasked(DataDe, CDate(sParam))
    
    'Exibe data final
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 117044

    Call DateParaMasked(DataAte, CDate(sParam))
           
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 117036 To 117044

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172634)

    End Select

    Exit Function

End Function

'Public Function TP_Caixa_Le(objCaixaMaskEdBox As Object, objCaixa As ClassCaixa) As Long
''Lê a Caixa com Código ou NomeRed em objCaixaMaskEdBox.Text
''Devolve em objCaixa. Coloca código-NomeReduzido no .Text
'
'Dim sCaixa As String
'Dim iCodigo As Integer
'Dim Caixa As Object
'Dim lErro As Long
'
'On Error GoTo Erro_TP_Caixa_Le
'
'    Set Caixa = objCaixaMaskEdBox
'    sCaixa = Trim(Caixa.Text)
'
'    'Tenta extrair código de sCaixa
'    iCodigo = Codigo_Extrai(sCaixa)
'
'    'Se é do tipo código
'    If iCodigo > 0 Then
'
'        objCaixa.iCodigo = iCodigo
'
'        'verifica se o codigo existe
'        lErro = CF("Caixas_Le", objCaixa)
'        If lErro <> SUCESSO And lErro <> 79405 Then gError 116174
'
'        'sem dados
'        If lErro = 79405 Then gError 116175
'
'        Caixa.Text = objCaixa.iCodigo & SEPARADOR & objCaixa.sNomeReduzido
'
'    Else  'Se é do tipo String
'
'         objCaixa.sNomeReduzido = sCaixa
'
'         'verifica se o nome reduzido existe
'         lErro = CF("Caixa_Le_NomeReduzido", objCaixa)
'         If lErro <> SUCESSO And lErro <> 79582 Then gError 116176
'
'         'sem dados
'         If lErro = 79582 Then gError 116177
'
'         'NomeControle.text recebe codigo - nome_reduzido
'         Caixa.Text = CStr(objCaixa.iCodigo) & SEPARADOR & sCaixa
'
'    End If
'
'    TP_Caixa_Le = SUCESSO
'
'    Exit Function
'
'Erro_TP_Caixa_Le:
'
'    TP_Caixa_Le = gErr
'
'    Select Case gErr
'
'        Case 116176, 116174 'Tratados nas rotinas chamadas
'
'        Case 116175, 116177 'Caixa com Codigo / NomeReduzido não cadastrado
'
'        Case Else
'             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172635)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim objOrcamento As ClassCupomFiscal

On Error GoTo Erro_objEventoOrcamento_evSelecao
    
    Set objOrcamento = obj1

    'Preenche campo Orcamento
    If giOrcamentoDe = 1 Then
        OrcamentoDe.PromptInclude = False
        OrcamentoDe.Text = CStr(objOrcamento.lNumOrcamento)
        OrcamentoDe.PromptInclude = True
        OrcamentoDe_Validate (bSGECancelDummy)
    Else
        OrcamentoAte.PromptInclude = False
        OrcamentoAte.Text = CStr(objOrcamento.lNumOrcamento)
        OrcamentoAte.PromptInclude = True
        OrcamentoAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoOrcamento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172636)

    End Select

    Exit Sub

End Sub
