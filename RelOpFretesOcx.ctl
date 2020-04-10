VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFretes 
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ScaleHeight     =   4200
   ScaleWidth      =   8145
   Begin VB.ComboBox PlacaUF 
      Height          =   315
      ItemData        =   "RelOpFretesOcx.ctx":0000
      Left            =   750
      List            =   "RelOpFretesOcx.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1545
      Width           =   735
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   150
      TabIndex        =   20
      Top             =   690
      Width           =   5565
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1725
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   750
         TabIndex        =   22
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4095
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   330
         Left            =   3105
         TabIndex        =   24
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIniPrev 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   300
         Width           =   390
      End
      Begin VB.Label dFimPrev 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2715
         TabIndex        =   25
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5895
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFretesOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFretesOcx.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFretesOcx.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFretesOcx.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   6090
      Picture         =   "RelOpFretesOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   780
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFretesOcx.ctx":0A9A
      Left            =   1890
      List            =   "RelOpFretesOcx.ctx":0A9C
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   195
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   150
      TabIndex        =   7
      Top             =   3195
      Width           =   5565
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   300
         Left            =   630
         TabIndex        =   8
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   300
         Left            =   3255
         TabIndex        =   9
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Região"
      Height          =   1245
      Left            =   150
      TabIndex        =   0
      Top             =   1875
      Width           =   5565
      Begin MSMask.MaskEdBox RegiaoDe 
         Height          =   315
         Left            =   585
         TabIndex        =   1
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoAte 
         Height          =   315
         Left            =   585
         TabIndex        =   2
         Top             =   765
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelRegiaoAte 
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
         Height          =   255
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   810
         Width           =   435
      End
      Begin VB.Label LabelRegiaoDe 
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
         Height          =   255
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
      Begin VB.Label RegiaoDeDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2025
         TabIndex        =   4
         Top             =   315
         Width           =   3120
      End
      Begin VB.Label RegiaoAteDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2025
         TabIndex        =   3
         Top             =   765
         Width           =   3120
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "U.F.:"
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
      Index           =   45
      Left            =   300
      TabIndex        =   28
      Top             =   1590
      Width           =   435
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
      Index           =   0
      Left            =   1185
      TabIndex        =   19
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFretes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'Property Variables:
Dim m_Caption As String
Event Unload()

'Browses
Private WithEvents objEventoRegiaoVendaDe As AdmEvento
Attribute objEventoRegiaoVendaDe.VB_VarHelpID = -1
Private WithEvents objEventoRegiaoVendaAte As AdmEvento
Attribute objEventoRegiaoVendaAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoRegiaoVendaDe = New AdmEvento
    Set objEventoRegiaoVendaAte = New AdmEvento
    
    'Carrega a combo PlacaUF
    lErro = Carrega_PlacaUF()
    If lErro <> SUCESSO Then gError 125755
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125756

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125755, 125756

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179518)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125757

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125758

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 125757
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 125758

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179519)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO ***
Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO***
Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    'se está Preenchido
    If Len(Trim(ClienteDe.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lErro <> SUCESSO Then gError 125759

    End If

    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179520)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    'Se está Preenchido
    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lErro <> SUCESSO Then gError 125760

    End If

    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125760

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179521)

    End Select

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe estiver preenchida
    If Len(DataDe.ClipText) > 0 Then

        sDataInic = DataDe.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 125761

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179522)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'se a DataAte estiver preenchida
    If Len(DataAte.ClipText) > 0 Then

        sDataFim = DataAte.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 125762

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125762

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179523)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoAte_Validate
    
    lErro = RegiaoVenda_Perde_Foco(RegiaoAte, RegiaoAteDesc)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 125763
       
    If lErro = 87199 Then gError 125764
        
    Exit Sub

Erro_RegiaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125763
        
        Case 125764
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179524)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sRegiaoInicial As String
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoDe_Validate

    lErro = RegiaoVenda_Perde_Foco(RegiaoDe, RegiaoDeDesc)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 125765
       
    If lErro = 87199 Then gError 125766
    
    Exit Sub

Erro_RegiaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125765
        
        Case 125766
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179525)

    End Select

    Exit Sub

End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 125767

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 125768

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 125767

        Case 125768
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179526)

    End Select

    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125769

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 125769
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179527)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125769

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 125769
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179528)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125770

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 125770
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179529)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125771

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 125771
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179530)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub LabelRegiaoAte_Click()
    
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoAte.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoAte.Text)
    
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVendaAte)
    
End Sub

Private Sub LabelRegiaoDe_Click()

Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoDe.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoDe.Text)
        
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVendaDe)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 125772

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125772

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179531)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa a tela
    lErro = LimpaRelatorioFretes()
    If lErro <> SUCESSO Then gError 125773
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 125773
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179532)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125774

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125775

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioFretes()
        If lErro <> SUCESSO Then gError 125776
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125774
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125775, 125776

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179533)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 125777

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125778

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125779
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 125780
    
    'Limpa a tela
    lErro = LimpaRelatorioFretes()
    If lErro <> SUCESSO Then gError 125781
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125777
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125778 To 125781
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179534)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** FUNÇÕES DE APOIO A TELA - INÍCIO
Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179535)

    End Select

End Function

Private Function LimpaRelatorioFretes()
'Limpa a tela RelOpRentCliProduto

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioFretes

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125782
    
    ComboOpcoes.Text = ""
   
    RegiaoDeDesc.Caption = ""
    RegiaoAteDesc.Caption = ""
   
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125783
    
    LimpaRelatorioFretes = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioFretes:

    LimpaRelatorioFretes = gErr
    
    Select Case gErr
    
        Case 125782, 125783
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179536)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
        
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")
        
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179537)

    End Select

    Exit Function

End Function

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco

        
    If Len(Trim(Regiao.Text)) > 0 Then
        
        objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
    
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 125784
    
        If lErro = 16137 Then gError 125785

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr

        Case 125784
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)

        Case 125785
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179538)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long, lNumIntRel As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim lClienteDe As Long, lClienteAte As Long, dtDataDe As Date, dtDataAte As Date, sUF As String, iRegiaoDe As Integer, iRegiaoAte As Integer

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sCliente_De, sCliente_Ate)
    If lErro <> SUCESSO Then gError 125786
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125787
        
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 125788
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 125789
    
    'Inclui a região
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125790

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125791
    
    'Inclui a data
    If DataDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 125792
    
    If DataAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 125793
    
    'Inclui a UF
    lErro = objRelOpcoes.IncluirParametro("TUF", PlacaUF.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125805
    
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_De, sCliente_Ate)
    If lErro <> SUCESSO Then gError 125794
    
    If bExecutar Then
    
        lErro = RelAnaliseFretes_Prepara(lNumIntRel, lClienteDe, lClienteAte, MaskedParaDate(DataDe), MaskedParaDate(DataAte), sUF, iRegiaoDe, iRegiaoAte)
        If lErro <> SUCESSO Then gError 125794
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 125788
    
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125786 To 125794, 125805
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179539)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_De As String, sCliente_Ate As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
   
   'Verifica se o Cliente inicial foi preenchido
    If ClienteDe.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteDe.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteAte.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteAte.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        
        If CInt(sCliente_De) > CInt(sCliente_Ate) Then gError 125795
        
    End If
    
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoDe.Text)) > 0 And Len(Trim(RegiaoAte.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If CLng(RegiaoDe.Text) > CLng(RegiaoAte.Text) Then gError 125796
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 125797
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 125795
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteDe.SetFocus
        
        Case 125796
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)
        
        Case 125797
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179540)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_De As String, sCliente_Ate As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
'    'Verifica se o Cliente Inicial foi preenchido
'    If sCliente_De <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvInt(CInt(sCliente_De))
'
'    End If
'
'    'Verifica se o Cliente Final foi preenchido
'    If sCliente_Ate <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvInt(CInt(sCliente_Ate))
'
'    End If
'
'    'se a região estiver preenchida
'    If RegiaoDe.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "RegiaoVenda >= " & Forprint_ConvInt(Codigo_Extrai(RegiaoDe.Text))
'
'    End If
'
'    If RegiaoAte.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(Codigo_Extrai(RegiaoAte.Text))
'
'    End If
'
'    'se a data estiver preenchida
'    If Trim(DataDe.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))
'
'    End If
'
'    If Trim(DataAte.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))
'
'    End If
'
'    'se a UF estiver preenchida
'    If PlacaUF.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "UF = " & Forprint_ConvTexto(PlacaUF.Text)
'
'    End If
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If
    
    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179541)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125798
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 125799
    
    ClienteDe.Text = sParam
    Call ClienteDe_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 125800
    
    ClienteAte.Text = sParam
    Call ClienteAte_Validate(bSGECancelDummy)
    
    'pega Região de Venda Inicial e exibe
    'sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAINIC", sParam)
    If lErro <> SUCESSO Then gError 125801

    RegiaoDe.Text = sParam
    Call RegiaoDe_Validate(bSGECancelDummy)
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAFIM", sParam)
    If lErro <> SUCESSO Then gError 125802

    RegiaoAte.Text = sParam
    Call RegiaoAte_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 125803

    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 125804

    Call DateParaMasked(DataAte, CDate(sParam))
        
    'pega Frete e exibe
    lErro = objRelOpcoes.ObterParametro("TUF", sParam)
    If lErro <> SUCESSO Then gError 125806

    If sParam <> "" Then
        
        PlacaUF.Text = sParam
        Call PlacaUF_Validate(bSGECancelDummy)
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125798 To 125804, 125806
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179542)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

'*** FUNÇÕES DO BROWSER - INÍCIO ***
Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteDe.Text = CStr(objCliente.lCodigo)
    Call ClienteDe_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteAte.Text = CStr(objCliente.lCodigo)
    Call ClienteAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRegiaoVendaDe_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    RegiaoDe.Text = objRegiaoVenda.iCodigo
    RegiaoDeDesc.Caption = objRegiaoVenda.sDescricao

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRegiaoVendaAte_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    RegiaoAte.Text = objRegiaoVenda.iCodigo
    RegiaoAteDesc.Caption = objRegiaoVenda.sDescricao

    Me.Show

    Exit Sub

End Sub
'*** FUNÇÕES DO BROWSER - FIM ***

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
        
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Relação de Fretes por Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFretes"
    
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

Function RelAnaliseFretes_Prepara(lNumIntRel As Long, lClienteDe As Long, lClienteAte As Long, dtDataDe As Date, dtDataAte As Date, sUF As String, iRegiaoDe As Integer, iRegiaoAte As Integer) As Long

Dim lErro As Long, dValorFaturado As Double, dValorFaturadoAcum As Double, dValorFrete As Double, dFretesAcum As Double, dFretesAdicionaisAcum As Double
Dim lNumIntNF As Long, sUFAux As String, iRegiao As Integer, lCliente As Long
Dim lNumIntNFAnt As Long, sUFAuxAnt As String, iRegiaoAnt As Integer, lClienteAnt As Long
Dim lTransacao As Long, alComando(1 To 2) As Long, iIndice As Integer
Dim sSQL As String
Dim vsUFAux, viRegiao, vlCliente, vlNumIntNF, vdValorFaturado, vdValorFrete As Variant

On Error GoTo Erro_RelAnaliseFretes_Prepara

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 128180

    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 128181
    
    'Obtêm o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_RENTABILIDADECLI", lNumIntRel)
    If lErro <> SUCESSO Then gError 124259
    
    'Alterado por Wagner 05/08/04
    'Alteracão: Montagem do SQL dinamico
    'Motivo: Filtros opcionais
    Call RelAnaliseFretesParteFixa_Prepara(sUF, iRegiaoDe, iRegiaoAte, lClienteDe, lClienteAte, sSQL)
    
    vsUFAux = String(STRING_ESTADO_SIGLA, 0)
    viRegiao = CInt(0)
    vlCliente = CLng(0)
    vlNumIntNF = CLng(0)
    vdValorFaturado = CDbl(0)
    vdValorFrete = CDbl(0)
    
    lErro = RelAnaliseFretesParteDinamica_Prepara(alComando(1), vsUFAux, viRegiao, vlCliente, vlNumIntNF, vdValorFaturado, vdValorFrete, dtDataDe, dtDataAte, giFilialEmpresa, sUF, iRegiaoDe, iRegiaoAte, lClienteDe, lClienteAte, sSQL)
    If lErro <> SUCESSO Then gError 129000
       
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124261
    'Fim da alteracão
        
    sUFAux = vsUFAux
    iRegiao = viRegiao
    lCliente = vlCliente
    lNumIntNF = vlNumIntNF
    dValorFaturado = vdValorFaturado
    dValorFrete = vdValorFrete
       
    If lErro = AD_SQL_SUCESSO Then
        
        sUFAuxAnt = sUFAux
        iRegiaoAnt = iRegiao
        lClienteAnt = lCliente
        lNumIntNFAnt = 0
        
        Do While lErro = AD_SQL_SUCESSO
        
            If (sUFAuxAnt <> sUFAux) Or (iRegiaoAnt <> iRegiao) Or (lClienteAnt <> lCliente) Then
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO RelAnaliseFretes (NumIntRel,UF,Regiao,Cliente,ValorFaturado,Fretes,Despesas,ValorTotal) VALUES (?,?,?,?,?,?,?,?)", _
                    lNumIntRel, sUFAuxAnt, iRegiaoAnt, lClienteAnt, dValorFaturadoAcum, dFretesAcum, dFretesAdicionaisAcum, Round(dFretesAcum + dFretesAdicionaisAcum, 2))
                If lErro <> AD_SQL_SUCESSO Then gError 124262
        
                sUFAuxAnt = sUFAux
                iRegiaoAnt = iRegiao
                lClienteAnt = lCliente
                
                dValorFaturadoAcum = 0
                dFretesAcum = 0
                dFretesAdicionaisAcum = 0
            
            End If
            
            dValorFaturadoAcum = dValorFaturadoAcum + dValorFaturado
            
            If lNumIntNFAnt <> lNumIntNF Then
                
                dFretesAcum = dFretesAcum + dValorFrete
                lNumIntNFAnt = lNumIntNF
                
            Else
            
                dFretesAdicionaisAcum = dFretesAdicionaisAcum + dValorFrete
                
            End If
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124263
        
            sUFAux = vsUFAux
            iRegiao = viRegiao
            lCliente = vlCliente
            lNumIntNF = vlNumIntNF
            dValorFaturado = vdValorFaturado
            dValorFrete = vdValorFrete
       
        Loop
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelAnaliseFretes (NumIntRel,UF,Regiao,Cliente,ValorFaturado,Fretes,Despesas,ValorTotal) VALUES (?,?,?,?,?,?,?,?)", _
            lNumIntRel, sUFAux, iRegiao, lCliente, dValorFaturadoAcum, dFretesAcum, dFretesAdicionaisAcum, Round(dFretesAcum + dFretesAdicionaisAcum, 2))
        If lErro <> AD_SQL_SUCESSO Then gError 124264
    
    End If
        
    'Fecha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128198

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelAnaliseFretes_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseFretes_Prepara:

    RelAnaliseFretes_Prepara = gErr
     
    Select Case gErr
          
        Case 128180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 128181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 128198
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 124259 To 124264
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_RELANALISEFRETES", gErr)
            
        Case 129000
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_RELANALISEFRETES", gErr)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179543)
     
    End Select
     
    Call Transacao_Rollback
    
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Sub RelAnaliseFretesParteFixa_Prepara(ByVal vUF As Variant, ByVal vRegiaoDe As Variant, ByVal vRegiaoAte As Variant, ByVal vClienteDe As Variant, ByVal vClienteAte As Variant, sSQL As String)
'monta o comando SQL para obtencao das fretes dinamicamente e retorna.
Dim sSelect As String, sWhere As String, sFrom As String, sOrderBy As String

On Error GoTo Erro_RelAnaliseFretesParteFixa_Prepara

    sSelect = "SELECT Enderecos.SiglaEstado, " & _
                       "FiliaisClientes.Regiao, " & _
                       "NFVenda.Cliente, " & _
                       "NFVenda.NumIntDoc, " & _
                       "NFVenda.ValorTotal, " & _
                       "ConhecTransp.ValorTotal "
    
    sFrom = "   FROM  NFiscal NFVenda, " & _
                     "NFiscal ConhecTransp, " & _
                     "FiliaisClientes, " & _
                     "Enderecos "
                     
    sWhere = "  WHERE  FiliaisClientes.Endereco = Enderecos.Codigo AND " & _
                      "FiliaisClientes.CodCliente = NFVenda.Cliente AND " & _
                      "FiliaisClientes.CodFilial = NFVenda.FilialCli AND " & _
                      "NFVenda.Status <> 7 AND " & _
                      "ConhecTransp.Status <> 7 AND " & _
                      "ConhecTransp.NumIntNotaOriginal = NFVenda.NumIntDoc AND " & _
                      "ConhecTransp.TipoNFiscal = 105 AND " & _
                      "NFVenda.DataEmissao BETWEEN ? AND ? AND " & _
                      "NFVenda.FilialEmpresa = ? "
     
     sOrderBy = "ORDER BY Enderecos.SiglaEstado, " & _
                         "FiliaisClientes.Regiao, " & _
                         "NFVenda.Cliente, " & _
                         "NFVenda.NumIntDoc, " & _
                         "ConhecTransp.DataEmissao, " & _
                         "ConhecTransp.NumNotaFiscal "
                         
    If Len(Trim(vUF)) <> 0 Then
        sWhere = sWhere & "AND Enderecos.SiglaEstado = ? "
    End If
    
    If vRegiaoDe <> 0 Then
        sWhere = sWhere & "AND FiliaisClientes.Regiao >= ? "
    End If
    
    If vRegiaoAte <> 0 Then
        sWhere = sWhere & "AND FiliaisClientes.Regiao <= ? "
    End If
   
    If vClienteDe <> 0 Then
        sWhere = sWhere & "AND NFVenda.Cliente >= ? "
    End If
    
    If vClienteAte <> 0 Then
        sWhere = sWhere & "AND NFVenda.Cliente <= ? "
    End If
    
    sSQL = sSelect & sFrom & sWhere & sOrderBy

    Exit Sub

Erro_RelAnaliseFretesParteFixa_Prepara:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179544)

    End Select

    Exit Sub

End Sub

Private Function RelAnaliseFretesParteDinamica_Prepara(ByVal lComando As Long, vUFAux As Variant, vRegiao As Variant, vCliente As Variant, vNumIntNF As Variant, vValorFaturado As Variant, vValorFrete As Variant, vDataDe As Variant, vDataAte As Variant, vFilialEmpresa As Variant, vUF As Variant, vRegiaoDe As Variant, vRegiaoAte As Variant, vClienteDe As Variant, vClienteAte As Variant, ByVal sSQL As String) As Long

Dim lErro As Long

On Error GoTo Erro_RelAnaliseFretesParteDinamica_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129001

    lErro = Comando_BindVarInt(lComando, vUFAux)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129002

    lErro = Comando_BindVarInt(lComando, vRegiao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129003
    
    lErro = Comando_BindVarInt(lComando, vCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129004

    lErro = Comando_BindVarInt(lComando, vNumIntNF)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129005

    lErro = Comando_BindVarInt(lComando, vValorFaturado)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129006
    
    lErro = Comando_BindVarInt(lComando, vValorFrete)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129007

    lErro = Comando_BindVarInt(lComando, vDataDe)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129008

    lErro = Comando_BindVarInt(lComando, vDataAte)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129009

    lErro = Comando_BindVarInt(lComando, vFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129010

    If Len(Trim(vUF)) <> 0 Then
        lErro = Comando_BindVarInt(lComando, vUF)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129011
    End If
    
    If vRegiaoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vRegiaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129012
    End If
    
    If vRegiaoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vRegiaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129013
    End If
   
    If vClienteDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vClienteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129014
    End If
    
    If vClienteAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vClienteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129015
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129016
    
    RelAnaliseFretesParteDinamica_Prepara = SUCESSO

    Exit Function

Erro_RelAnaliseFretesParteDinamica_Prepara:

    RelAnaliseFretesParteDinamica_Prepara = gErr

    Select Case gErr
    
        Case 129001 To 129016
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_RELANALISEFRETES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179545)

    End Select

    Exit Function

End Function
