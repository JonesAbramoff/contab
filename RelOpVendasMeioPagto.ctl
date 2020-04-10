VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpVendasMeioPagto 
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   6705
   Begin VB.Frame FrameNivelDetalhes 
      Caption         =   "Nível de Detalhes"
      Height          =   1815
      Left            =   240
      TabIndex        =   29
      Top             =   3915
      Width           =   4215
      Begin VB.OptionButton MeiosPagtoAdministradoraParc 
         Caption         =   "Meios de Pagto + Administradoras + Parcelamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   3495
      End
      Begin VB.OptionButton MeiosPagtoAdministradora 
         Caption         =   "Meios de Pagto + Administradoras"
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
         TabIndex        =   8
         Top             =   780
         Width           =   3255
      End
      Begin VB.OptionButton MeiosPagto 
         Caption         =   "Meios de Pagto"
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
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame FrameTipoMeioPagto 
      Caption         =   "Meios de Pagamento"
      Height          =   1290
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   4215
      Begin VB.ComboBox TipoMeioPagtoAte 
         Height          =   315
         ItemData        =   "RelOpVendasMeioPagto.ctx":0000
         Left            =   960
         List            =   "RelOpVendasMeioPagto.ctx":0002
         TabIndex        =   6
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   825
         Width           =   2880
      End
      Begin VB.ComboBox TipoMeioPagtoDe 
         Height          =   315
         ItemData        =   "RelOpVendasMeioPagto.ctx":0004
         Left            =   960
         List            =   "RelOpVendasMeioPagto.ctx":0006
         TabIndex        =   5
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   360
         Width           =   2880
      End
      Begin VB.Label LabelMeioPagtoAte 
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
         Left            =   480
         TabIndex        =   28
         Top             =   885
         Width           =   360
      End
      Begin VB.Label LabelMeioPagtoDe 
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
         Left            =   525
         TabIndex        =   27
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   240
      TabIndex        =   23
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
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
         Left            =   2280
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpVendasMeioPagto.ctx":0008
      Left            =   1080
      List            =   "RelOpVendasMeioPagto.ctx":000A
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2670
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
         Top             =   285
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
         Top             =   285
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
         TabIndex        =   21
         Top             =   345
         Width           =   315
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
         TabIndex        =   20
         Top             =   345
         Width           =   360
      End
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
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpVendasMeioPagto.ctx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpVendasMeioPagto.ctx":0166
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpVendasMeioPagto.ctx":02F0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "RelOpVendasMeioPagto.ctx":0822
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox ExibirPor 
      Caption         =   "Exibir valores por"
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
      Left            =   240
      TabIndex        =   10
      Top             =   5880
      Width           =   3060
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
      Left            =   4740
      Picture         =   "RelOpVendasMeioPagto.ctx":09A0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   945
      Width           =   1605
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
      TabIndex        =   22
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpVendasMeioPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giCaixaDe As Integer

Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Vendas x Meios de Pagamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpVendasMeioPagto"

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

Private Sub objEventoCaixa_evSelecao(obj1 As Object)
'evento de inclusao de item selecionado no browser caixa

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao

    Set objCaixa = obj1
    
    'Preenche campo Caixa
    If giCaixaDe = 1 Then
        CaixaDe.Text = objCaixa.iCodigo
        CaixaDe_Validate (bSGECancelDummy)
    Else
        CaixaAte.Text = objCaixa.iCodigo
        CaixaAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173622)

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

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 117046
    
    If (MeiosPagto.Value = True) And (ExibirPor.Value = Unchecked) Then gobjRelatorio.sNomeTsk = "VNDMPRES"
    If (MeiosPagto.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VNDMPFI"
    If (MeiosPagto.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VNDMPET"
    If (MeiosPagtoAdministradora.Value = True) And (ExibirPor.Value = Unchecked) Then gobjRelatorio.sNomeTsk = "VDMPADRE"
    If (MeiosPagtoAdministradora.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VDMPADFI"
    If (MeiosPagtoAdministradora.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VDMPADET"
    If (MeiosPagtoAdministradoraParc.Value = True) And (ExibirPor.Value = Unchecked) Then gobjRelatorio.sNomeTsk = "VDMPAPRE"
    If (MeiosPagtoAdministradoraParc.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VDMPAPFI"
    If (MeiosPagtoAdministradoraParc.Value = True) And (ExibirPor.Value = Checked) And giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "VDMPAPET"

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 117046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173623)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Nome da opção de Relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 117047

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 117048

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 117049

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 117050

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 117047
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 117048, 117049, 117050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173624)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giFilialEmpresa <> EMPRESA_TODA Then
        ExibirPor.Caption = "Exibir valores por caixa"
    Else
        ExibirPor.Caption = "Exibir valores por filial"
    End If
   
    Call Carrega_TipoMeioPagto

    Set objEventoCaixa = New AdmEvento

    giCaixaDe = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173625)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCaixa = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

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

Private Sub LabelMeioPagtoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMeioPagtoDe, Source, X, Y)
End Sub

Private Sub LabelMeioPagtoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMeioPagtoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelMeioPagtoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMeioPagtoAte, Source, X, Y)
End Sub

Private Sub LabelMeioPagtoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMeioPagtoAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        sDataFim = DataAte.Text

        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 117051

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 117051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173626)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        sDataInic = DataDe.Text

        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 117052

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 117052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173627)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 117053

    lErro = objRelOpcoes.ObterParametro("TSIGLADOC", sParam)
    If lErro <> SUCESSO Then gError 117054

    'Verifica qual OptionButton, está selecionada
    If sParam = "MP" Then MeiosPagto.Value = True

    If sParam = "MPA" Then MeiosPagtoAdministradora.Value = True

    If sParam = "MPAP" Then MeiosPagtoAdministradoraParc.Value = True

    'Exibe Meio de Pagto inicial
    lErro = objRelOpcoes.ObterParametro("NTMPAGTODE", sParam)
    If lErro <> SUCESSO Then gError 117055

    TipoMeioPagtoDe.Text = sParam
    Call TipoMeioPagtoDe_Validate(bSGECancelDummy)
    
    'Exibe Meio de Pagto final
    lErro = objRelOpcoes.ObterParametro("NTMPAGTOATE", sParam)
    If lErro <> SUCESSO Then gError 117056

    TipoMeioPagtoAte.Text = sParam
    Call TipoMeioPagtoAte_Validate(bSGECancelDummy)
    
    'Exibe Caixa inicial
    lErro = objRelOpcoes.ObterParametro("NCAIXADE", sParam)
    If lErro <> SUCESSO Then gError 117057

    CaixaDe.PromptInclude = False
    CaixaDe.Text = sParam
    CaixaDe.PromptInclude = True
    Call CaixaDe_Validate(bSGECancelDummy)

    'Exibe Caixa final
    lErro = objRelOpcoes.ObterParametro("NCAIXAATE", sParam)
    If lErro <> SUCESSO Then gError 117058

    CaixaAte.PromptInclude = False
    CaixaAte.Text = sParam
    CaixaAte.PromptInclude = True
    Call CaixaAte_Validate(bSGECancelDummy)

    'Exibe data inicial
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 117059

    Call DateParaMasked(DataDe, CDate(sParam))

    'Exibe data final
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 117060

    Call DateParaMasked(DataAte, CDate(sParam))

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 117053 To 117060

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173628)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sTMPagto_De As String
Dim sTMPagto_Ate As String
Dim sCaixa_De As String
Dim sCaixa_Ate As String
Dim sSigla As String, dValorBD1 As Double, dValorBD2 As Double

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sTMPagto_De, sTMPagto_Ate, sCaixa_De, sCaixa_Ate)
    If lErro <> SUCESSO Then gError 117061

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 117062

    If MeiosPagto.Value = True Then
        sSigla = "MP"
    ElseIf MeiosPagtoAdministradora.Value = True Then
        sSigla = "MPA"
    ElseIf MeiosPagtoAdministradoraParc.Value = True Then
        sSigla = "MPAP"
    End If

    lErro = objRelOpcoes.IncluirParametro("TSIGLADOC", sSigla)
    If lErro <> AD_BOOL_TRUE Then gError 117063

    lErro = objRelOpcoes.IncluirParametro("NCAIXADE", sCaixa_De)
    If lErro <> AD_BOOL_TRUE Then gError 117064
    
    lErro = objRelOpcoes.IncluirParametro("TCAIXADE", Trim(CaixaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117064

    lErro = objRelOpcoes.IncluirParametro("NCAIXAATE", sCaixa_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117065
    
    lErro = objRelOpcoes.IncluirParametro("TCAIXAATE", Trim(CaixaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117065

    lErro = objRelOpcoes.IncluirParametro("NTMPAGTODE", sTMPagto_De)
    If lErro <> AD_BOOL_TRUE Then gError 117066

    lErro = objRelOpcoes.IncluirParametro("TTMPAGTODE", TipoMeioPagtoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 117067

    lErro = objRelOpcoes.IncluirParametro("NTMPAGTOATE", sTMPagto_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117068

    lErro = objRelOpcoes.IncluirParametro("TTMPAGTOATE", TipoMeioPagtoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 117069

    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 117070

    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 117071
    
    If bExecutando Then
    
        lErro = CF("SldDiaMeioPagtoCx_Le_AcumPeriodo", giFilialEmpresa, StrParaInt(sCaixa_De), StrParaInt(sCaixa_Ate), StrParaInt(sTMPagto_De), StrParaInt(sTMPagto_Ate), StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), dValorBD1)
        If lErro <> SUCESSO Then gError 117068
    
        lErro = CF("SldDiaMeioPagtoCx_Le_AcumPeriodo", giFilialEmpresa, StrParaInt(sCaixa_De), StrParaInt(sCaixa_Ate), StrParaInt(sTMPagto_De), StrParaInt(sTMPagto_Ate), StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), dValorBD2, True)
        If lErro <> SUCESSO Then gError 117068
    
        lErro = objRelOpcoes.IncluirParametro("NVLRTOT1", CStr(dValorBD1))
        If lErro <> AD_BOOL_TRUE Then gError 117068
        
        lErro = objRelOpcoes.IncluirParametro("NVLRTOT2", CStr(dValorBD2))
        If lErro <> AD_BOOL_TRUE Then gError 117068
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sTMPagto_De, sTMPagto_Ate, sCaixa_De, sCaixa_Ate)
    If lErro <> SUCESSO Then gError 117072
    PreencherRelOp = SUCESSO
    
    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 117061 To 117072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173629)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sTMPagto_De As String, sTMPagto_Ate As String, sCaixa_De As String, sCaixa_Ate As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica TipoMeioPagto Inicial e Final
    If TipoMeioPagtoDe.Text <> "" Then
        sTMPagto_De = CStr(Codigo_Extrai(TipoMeioPagtoDe.Text))
    Else
        sTMPagto_De = ""
    End If

    If TipoMeioPagtoAte.Text <> "" Then
        sTMPagto_Ate = CStr(Codigo_Extrai(TipoMeioPagtoAte.Text))
    Else
        sTMPagto_Ate = ""
    End If

    If sTMPagto_De <> "" And sTMPagto_Ate <> "" Then

        'Se o TipoMeioPagto Inicial for maior que o final --> erro
        If CInt(sTMPagto_De) > CInt(sTMPagto_Ate) Then gError 117073

    End If

     'critica Caixa Inicial e Final
    If CaixaDe.ClipText <> "" Then
        sCaixa_De = CStr(Codigo_Extrai(CaixaDe.ClipText))
    Else
        sCaixa_De = ""
    End If

    If CaixaAte.ClipText <> "" Then
        sCaixa_Ate = CStr(Codigo_Extrai(CaixaAte.ClipText))
    Else
        sCaixa_Ate = ""
    End If

    If sCaixa_De <> "" And sCaixa_Ate <> "" Then

        'Se o Caixa Inicial for maior que o final --> erro
        If CInt(sCaixa_De) > CInt(sCaixa_Ate) Then gError 117074

    End If

    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 117075

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 117075
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 117073
            Call Rotina_Erro(vbOKOnly, "ERRO_MEIODEPAGTO_INICIAL_MAIOR", gErr)
            TipoMeioPagtoDe.SetFocus
        
        Case 117074
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INICIAL_MAIOR", gErr)
            CaixaDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173630)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sTMPagto_De As String, sTMPagto_Ate As String, sCaixa_De As String, sCaixa_Ate As String) As Long
'Monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sTMPagto_De <> "" Then sExpressao = "TipoMeioPagto >= " & Forprint_ConvInt(CInt(sTMPagto_De))

   If sTMPagto_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoMeioPagto <= " & Forprint_ConvInt(CInt(sTMPagto_Ate))

    End If

   If sCaixa_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(CInt(sCaixa_De))

   End If

   If sCaixa_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(CInt(sCaixa_Ate))

    End If

    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173631)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 117076

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 117077

        'Retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click

        ComboOpcoes.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 117076
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 117077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173632)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 117078

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    MeiosPagto.Value = True
    ExibirPor.Value = Unchecked

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 117078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173633)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 117089

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 117089
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173634)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 117088

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 117088
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173635)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 117087

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 117087
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173636)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 117086

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 117086
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173637)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoMeioPagto() As Long
'Carrega as combos com os dados lidos do BD

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTMPLoja
Dim colTipoMeioPagto As New Collection

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Lê os tipos de Meio de Pagamento
    lErro = CF("TipoMeioPagto_Le_Todas", colTipoMeioPagto)
    If lErro <> SUCESSO Then gError 117081
    
    'Carrega na combo
    For Each objTipoMeioPagto In colTipoMeioPagto
        TipoMeioPagtoDe.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
        TipoMeioPagtoDe.ItemData(TipoMeioPagtoDe.NewIndex) = objTipoMeioPagto.iTipo
        
        TipoMeioPagtoAte.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
        TipoMeioPagtoAte.ItemData(TipoMeioPagtoAte.NewIndex) = objTipoMeioPagto.iTipo
    Next
    
    Carrega_TipoMeioPagto = SUCESSO
    
    Exit Function
    
Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = gErr
    
    Select Case gErr
    
        Case 117081
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173638)
            
    End Select
    
    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 117082
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 117083
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 117083
        
        Case 117082
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173639)

    End Select

    Exit Function

End Function

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
    
    giCaixaDe = 1

    If Len(Trim(CaixaDe.Text)) > 0 Then
        
        'instancia o obj
        Set objCaixa = New ClassCaixa
        
        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116209

        'código inexistente
        If lErro = 116175 Then gError 117084

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116211

    End If
    
    Exit Sub

Erro_CaixaDe_Validate:

    Cancel = True

    Select Case gErr
        
        Case 116209
        
        Case 116211
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)

        Case 117084
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_EXISTE", gErr, Trim(CaixaDe.Text))
             CaixaDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173640)

    End Select

End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    giCaixaDe = 0

    If Len(Trim(CaixaAte.Text)) > 0 Then

        'instancia o obj
        Set objCaixa = New ClassCaixa

        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler a Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaAte, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116212

        'código inexistente
        If lErro = 116175 Then gError 117085

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116214

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116212
        
        Case 116214
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)

        Case 117085
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_EXISTE", gErr, Trim(CaixaAte.Text))
            CaixaAte.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173641)

    End Select

End Sub

Private Sub LabelCaixaDe_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCaixaDe_Click
    
    giCaixaDe = 1

    If Len(Trim(CaixaDe.ClipText)) > 0 Then
        'Preenche com o Caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If

    'Faz com que o browser não exiba o caixa central
    sSelecao = " CaixaCod <> ? "
    colSelecao.Add CODIGO_CAIXA_CENTRAL
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    Else
    
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    End If
    
    Exit Sub

Erro_LabelCaixaDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173642)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCaixaAte_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCaixaAte_Click
    
    giCaixaDe = 0
    
    If Len(Trim(CaixaAte.ClipText)) > 0 Then
        'Preenche com o Caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If

    'Faz com que o browser não exiba o caixa central
    sSelecao = " CaixaCod <> ? "
    colSelecao.Add CODIGO_CAIXA_CENTRAL
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    Else

        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)

    End If
    
    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173643)

    End Select

    Exit Sub
End Sub

Private Sub TipoMeioPagtoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagtoDe_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TipoMeioPagtoDe.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(TipoMeioPagtoDe.Text)) = 0 Then Exit Sub

    lErro = Combo_Seleciona(TipoMeioPagtoDe, iCodigo)
    If lErro <> SUCESSO Then gError 117079

    Exit Sub

Erro_TipoMeioPagtoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 117079

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173644)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagtoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagtoAte_Validate

    'se uma opcao da lista estiver selecionada, OK
    If TipoMeioPagtoAte.ListIndex <> -1 Then Exit Sub

    If Len(Trim(TipoMeioPagtoAte.Text)) = 0 Then Exit Sub

    lErro = Combo_Seleciona(TipoMeioPagtoAte, iCodigo)
    If lErro <> SUCESSO Then gError 117080

    Exit Sub

Erro_TipoMeioPagtoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 117080

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173645)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click

        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click

        End If

    End If

End Sub


