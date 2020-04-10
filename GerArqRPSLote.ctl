VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GerArqRPSLoteOcx 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LockControls    =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   7995
   Begin VB.Frame Frame2 
      Caption         =   "A partir do RPS ..."
      Height          =   780
      Left            =   4125
      TabIndex        =   22
      Top             =   150
      Width           =   1890
      Begin MSMask.MaskEdBox RPSDe 
         Height          =   315
         Left            =   480
         TabIndex        =   23
         Top             =   255
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "No:"
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
         Height          =   315
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.CheckBox optNomeAuto 
      Caption         =   "Preencher o nome do arquivo automaticamente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3450
      TabIndex        =   7
      Top             =   1710
      Width           =   4380
   End
   Begin VB.TextBox NomeArquivo 
      Height          =   285
      Left            =   885
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton BotaoHistRPS 
      Caption         =   "Histórico de RCP enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   6075
      TabIndex        =   12
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CommandButton BotaoArqGerados 
      Caption         =   "Arquivos de RPS em Lote Gerados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4140
      TabIndex        =   11
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CommandButton BotaoRPSNaoNFE 
      Caption         =   "RPS Não Convertidos em NF-e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2220
      TabIndex        =   10
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CommandButton BotaoRPSNaoEnviados 
      Caption         =   "RPS Não Enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   285
      TabIndex        =   9
      Top             =   3120
      Width           =   1650
   End
   Begin VB.CheckBox optAtualizaCliEnd 
      Caption         =   "Atualizar Informações Cadastrais do Cliente e Endereço nos Recibos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   900
      TabIndex        =   8
      Top             =   2385
      Width           =   6420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos RPS"
      Height          =   780
      Left            =   150
      TabIndex        =   18
      Top             =   150
      Width           =   3900
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1740
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
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
         Top             =   285
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
         Left            =   3540
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   2535
         TabIndex        =   2
         Top             =   285
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
         TabIndex        =   20
         Top             =   330
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
         Left            =   2100
         TabIndex        =   19
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   885
      TabIndex        =   4
      Top             =   1230
      Width           =   6225
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7155
      TabIndex        =   5
      Top             =   1185
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6165
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   255
      Width           =   1650
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   90
         Picture         =   "GerArqRPSLote.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gera o arquivo de RPS em lote"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "GerArqRPSLote.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "GerArqRPSLote.ctx":0974
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
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
      Left            =   180
      TabIndex        =   21
      Top             =   1830
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Diretório:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1245
      Width           =   795
   End
End
Attribute VB_Name = "GerArqRPSLoteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer

'Variável utilizada para manuseio do grid
Dim objGridArquivos As AdmGrid

Dim gcolArquivos As Collection

'Variáveis das colunas do grid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Item_Col As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    Call Default_Tela

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192488)

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
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
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

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGerar_Click()
'Dispara a geração dos arquivos e relatórios selecionados

Dim lErro As Long
Dim sNomeArqParam As String
Dim objRPSCab As New ClassRPSCab

On Error GoTo Erro_BotaoGerar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Move_Tela_Memoria(objRPSCab)
    If lErro <> SUCESSO Then gError 192489
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 192490
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 192491
    
    If objRPSCab.dtDataInicio = DATA_NULA Then gError 192492
    If objRPSCab.dtDataFim = DATA_NULA Then gError 192493
    If objRPSCab.dtDataInicio > objRPSCab.dtDataFim Then gError 192494
    If objRPSCab.dtDataFim > Date Then gError 192665
    
    'inicia o Batch
    lErro = CF("RPS_Gera_Arquivo_Lote", objRPSCab)
    If lErro <> SUCESSO Then gError 192495
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192489, 192495
               
        Case 192490
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
            
        Case 192491
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeArquivo.SetFocus

        Case 192492
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
            DataDe.SetFocus
            
        Case 192493
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
            DataAte.SetFocus

        Case 192494
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            DataDe.SetFocus
            
        Case 192665
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_MAIOR_QUE_DATAHOJE", gErr)
            DataAte.SetFocus
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192496)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

'    'Verifica se algum campo foi alterado e confirma se o usuário deseja
'    'salvar antes de limpar
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then gError 192497

    'limpa a tela
    Call Limpa_Tela_ExportarDados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 192497
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192498)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub NomeDiretorio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub optAtualizaCliEnd_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub optNomeAuto_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Calcula_NomeArquivo
End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 192499

        Call Calcula_NomeArquivo

    End If

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 192499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192500)
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
        If lErro <> SUCESSO Then gError 192501

        Call Calcula_NomeArquivo

    End If

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 192501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192502)

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
        If lErro <> SUCESSO Then gError 192503
        
        Call Calcula_NomeArquivo

    End If

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 192503

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192504)

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
        If lErro <> SUCESSO Then gError 192505

        Call Calcula_NomeArquivo

    End If

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 192505

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192506)

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
    If lErro <> SUCESSO Then gError 192507
    
    Call Calcula_NomeArquivo

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 192507
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192508)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 192509
    
    Call Calcula_NomeArquivo

Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 192509
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192510)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

Private Function Move_Tela_Memoria(ByVal objRPSCab As ClassRPSCab) As Long
'Transfere os dados tela para objIN86Modelo

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    objRPSCab.dtDataInicio = StrParaDate(DataDe.Text)
    objRPSCab.dtDataFim = StrParaDate(DataAte.Text)
    objRPSCab.iFilialEmpresa = giFilialEmpresa
    
    If optAtualizaCliEnd.Value = vbChecked Then
        objRPSCab.iAtualizaDadosCliEnd = MARCADO
    Else
        objRPSCab.iAtualizaDadosCliEnd = DESMARCADO
    End If
        
    If InStr(1, NomeArquivo.Text, ".") = 0 Then
        objRPSCab.sNomeArquivo = NomeDiretorio.Text & NomeArquivo.Text & ".txt"
    Else
        objRPSCab.sNomeArquivo = NomeDiretorio.Text & NomeArquivo.Text
    End If
    
    objRPSCab.lRPSDe = StrParaLong(RPSDe.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192511)

    End Select

End Function

Private Sub Limpa_Tela_ExportarDados()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ExportarDados

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    Call Default_Tela

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ExportarDados:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192512)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de arquivos em Lote de RPS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GerArqRPSLote"

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

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192513

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192513, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192514)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Diretório"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192515)

    End Select

    Exit Sub
  
End Sub

Private Sub BotaoRPSNaoEnviados_Click()

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoRPSNaoEnviados_Click

    sFiltro = "Enviado = ?"
    colSelecao.Add 0
   
    Call Chama_Tela("RPSLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoRPSNaoEnviados_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192575)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoArqGerados_Click()

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoArqGerados_Click
   
    Call Chama_Tela("RPSCabLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoArqGerados_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192576)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoHistRPS_Click()

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoArqGerados_Click
   
    Call Chama_Tela("RPSEnviadosLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoArqGerados_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192577)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Default_Tela() As Long

Dim lErro As Long
Dim objRPSCab As New ClassRPSCab

On Error GoTo Erro_Default_Tela

    optNomeAuto.Value = vbChecked
    optAtualizaCliEnd.Value = vbUnchecked

    objRPSCab.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("RPS_Preenche_CabPadrao", objRPSCab)
    If lErro <> SUCESSO Then gError 192663
    
    If objRPSCab.dtDataInicio <> DATA_NULA And objRPSCab.dtDataFim <> DATA_NULA Then
        
        DataDe.PromptInclude = False
        DataDe.Text = Format(objRPSCab.dtDataInicio, "dd/mm/yy")
        DataDe.PromptInclude = True
        
        DataAte.PromptInclude = False
        DataAte.Text = Format(objRPSCab.dtDataFim, "dd/mm/yy")
        DataAte.PromptInclude = True

    End If
    
    NomeDiretorio.Text = objRPSCab.sNomeArquivo
    Call NomeDiretorio_Validate(bSGECancelDummy)
    
    Call Calcula_NomeArquivo
    
    Default_Tela = SUCESSO
    
    Exit Function

Erro_Default_Tela:

    Default_Tela = gErr

    Select Case gErr
    
        Case 192663
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192664)

    End Select

    Exit Function

End Function

Private Function Calcula_NomeArquivo() As Long

Dim lErro As Long
Dim sNomeArquivo As String

On Error GoTo Erro_Calcula_NomeArquivo

    If optNomeAuto.Value = vbChecked Then

        sNomeArquivo = "RPS_" & CStr(giFilialEmpresa)
        
        If StrParaDate(DataDe.Text) <> DATA_NULA Then
            sNomeArquivo = sNomeArquivo & "_" & Format(StrParaDate(DataDe.Text), "ddmmyyyy")
        End If
        
        If StrParaDate(DataAte.Text) <> DATA_NULA Then
            sNomeArquivo = sNomeArquivo & "_" & Format(StrParaDate(DataAte.Text), "ddmmyyyy")
        End If
        
        NomeArquivo.Text = sNomeArquivo
        
        optNomeAuto.Value = vbChecked

    End If

    Calcula_NomeArquivo = SUCESSO
    
    Exit Function

Erro_Calcula_NomeArquivo:

    Calcula_NomeArquivo = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192665)

    End Select

    Exit Function

End Function

Private Sub NomeArquivo_Change()
    iAlterado = REGISTRO_ALTERADO
    optNomeAuto.Value = vbUnchecked
End Sub

Private Sub NomeArquivo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeArquivo_Validate

    If InStr(1, NomeArquivo.Text, ".") <> 0 Then gError 192667
    If InStr(1, NomeArquivo.Text, ",") <> 0 Then gError 192667
    If InStr(1, NomeArquivo.Text, "(") <> 0 Then gError 192667
    If InStr(1, NomeArquivo.Text, ")") <> 0 Then gError 192667
    If InStr(1, NomeArquivo.Text, "\") <> 0 Then gError 192667
    If InStr(1, NomeArquivo.Text, "/") <> 0 Then gError 192667

    Exit Sub

Erro_NomeArquivo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 192667
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEARQUIVO_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192666)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRPSNaoNFE_Click()

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoRPSNaoNFE_Click

    sFiltro = "NOT EXISTS (SELECT N.NumIntDoc FROM NFe AS N WHERE N.SerieRPS = RPS.Serie AND N.DataEmissaoRPS = RPS.DataEmissao AND N.NumeroRPS = RPS.Numero AND N.FilialEmpresa = RPS.FilialEmpresa )"
   
    Call Chama_Tela("RPSLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoRPSNaoNFE_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192575)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub RPSDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RPSDe_Validate

    'Verifica se Ano está preenchida
    If Len(Trim(RPSDe.Text)) <> 0 Then

       'Critica a Ano
       lErro = Inteiro_Critica(RPSDe.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_RPSDe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143967)

    End Select

    Exit Sub

End Sub

Private Sub RPSDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(RPSDe, iAlterado)
End Sub

Private Sub RPSDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
