VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVExtraiDadosSigav 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   5070
   Begin VB.CommandButton BotaoVoucher 
      Caption         =   "Vouchers sem dados extraídos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   285
      TabIndex        =   4
      ToolTipText     =   "Vouchers sem dados extraídos do Sigav."
      Top             =   1845
      Width           =   1740
   End
   Begin VB.CommandButton BotaoExtrairSigav 
      Caption         =   "Extrair dados do Sigav"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3060
      TabIndex        =   5
      ToolTipText     =   "Extrai os dados dos vouchers do Sigav."
      Top             =   1845
      Width           =   1740
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3735
      ScaleHeight     =   495
      ScaleWidth      =   1005
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1065
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   525
         Picture         =   "TRVExtraiDadosSigav.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "TRVExtraiDadosSigav.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos Vouchers"
      Height          =   1005
      Left            =   285
      TabIndex        =   8
      Top             =   690
      Width           =   4530
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1740
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   720
         TabIndex        =   0
         Top             =   375
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3990
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   2985
         TabIndex        =   2
         Top             =   375
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelInicio 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelFim 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2550
         TabIndex        =   9
         Top             =   435
         Width           =   360
      End
   End
End
Attribute VB_Name = "TRVExtraiDadosSigav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192673)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long
'A tela não espera recebimento de parâmetros, portanto, essa função sempre retorna sucesso
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais

End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_ExtraiDados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192674)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 192675

    End If

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 192675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192676)
    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 192677

    End If

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 192677

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192678)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Se a data está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 192679

    End If

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 192679

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192680)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Se a data está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 192681

    End If

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 192681

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192682)

    End Select

    Exit Sub

End Sub

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub DataDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub DataAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 192684

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 192684
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192685)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 192686

Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 192686
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192687)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***


Private Sub Limpa_Tela_ExtraiDados()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ExtraiDados

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ExtraiDados:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192688)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Extração dos dados de Voucher do Sigav"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVExtraiDadosSigav"

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

'*** TRATAMENTO PARA MODO DE EDIÇÃO - INÍCIO ***
Private Sub LabelInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelInicio, Source, X, Y)
End Sub

Private Sub LabelFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFim, Button, Shift, X, Y)
End Sub

Private Sub LabelFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFim, Source, X, Y)
End Sub

'*** TRATAMENTO PARA MODO DE EDIÇÃO - FIM ***

Private Sub BotaoVoucher_Click()

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVoucher_Click

    sFiltro = "NOT EXISTS (SELECT V.NumVou FROM TRVVoucherInfoSigav AS V WHERE VoucherRapido.NumVou = V.NumVou AND VoucherRapido.Serie = V.Serie AND VoucherRapido.TipVou = V.Tipo)"
   
    Call Chama_Tela("VoucherRapidoLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoVoucher_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192689)
    
    End Select
    
    Exit Sub
    
End Sub

Sub BotaoExtrairSigav_Click()

Dim lErro As Long
Dim objTRVVoucherInfo As ClassTRVVoucherInfo
Dim objTRVVoucher As ClassTRVVouchers
Dim colVoucher As New Collection
Dim objSenha As New ClassSenha
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoExtrairSigav_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaDate(DataDe.Text) = DATA_NULA Then gError 192690
    If StrParaDate(DataAte.Text) = DATA_NULA Then gError 192691
    
    If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 192692
    
    Load SigavSenha
    lErro = SigavSenha.Trata_Parametros(objSenha)
    If lErro <> SUCESSO Then gError 192693
    SigavSenha.Show vbModal
    
    If Len(Trim(objSenha.sSenha)) = 0 Then gError 192694
    
    lErro = CF("TRVVouchers_Le_Periodo", StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), colVoucher)
    If lErro <> SUCESSO Then gError 192695
    
    For Each objTRVVoucher In colVoucher
    
        Set objTRVVoucherInfo = New ClassTRVVoucherInfo
        
        objTRVVoucherInfo.sTipo = objTRVVoucher.sTipVou
        objTRVVoucherInfo.lNumVou = objTRVVoucher.lNumVou
        objTRVVoucherInfo.sSerie = objTRVVoucher.sSerie
            
        lErro = Obter_Dados_Sigav(objTRVVoucherInfo, objSenha.sSenha)
        If lErro <> SUCESSO Then gError 192696
        
        lErro = CF("TRVVoucherInfoSigav_Grava", objTRVVoucherInfo)
        If lErro <> SUCESSO Then gError 192697
        
    Next
        
    GL_objMDIForm.MousePointer = vbDefault

    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")

    Exit Sub

Erro_BotaoExtrairSigav_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192690
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
            DataDe.SetFocus
            
        Case 192691
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
            DataAte.SetFocus

        Case 192692
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            DataDe.SetFocus
            
        Case 192693, 192695, 192696, 192697
        
        Case 192694
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192698)

    End Select

    Exit Sub

End Sub

