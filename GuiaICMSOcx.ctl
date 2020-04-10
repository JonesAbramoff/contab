VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GuiaICMSOcx 
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7080
   Begin VB.Frame Frame3 
      Caption         =   "Dados Principais"
      Height          =   1410
      Left            =   75
      TabIndex        =   31
      Top             =   705
      Width           =   6915
      Begin VB.Frame FrameUF 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   525
         Left            =   2820
         TabIndex        =   40
         Top             =   105
         Visible         =   0   'False
         Width           =   1770
         Begin VB.ComboBox UF 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   135
            Width           =   885
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   420
            TabIndex        =   41
            Top             =   195
            Width           =   720
         End
      End
      Begin VB.CommandButton BotaoConsulta 
         Caption         =   "Guias Cadastradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4620
         TabIndex        =   2
         Top             =   240
         Width           =   2220
      End
      Begin VB.TextBox LocalEntrega 
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1005
         Width           =   5505
      End
      Begin VB.TextBox OrgaoArrecadador 
         Height          =   285
         Left            =   4620
         MaxLength       =   20
         TabIndex        =   4
         Top             =   615
         Width           =   2235
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   1350
         TabIndex        =   0
         Top             =   255
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "############"
         PromptChar      =   " "
      End
      Begin VB.Label LabelValor 
         Caption         =   "Valor:"
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
         Left            =   825
         TabIndex        =   35
         Top             =   675
         Width           =   540
      End
      Begin VB.Label LabelLocalEntrega 
         Caption         =   "Local Entrega:"
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
         Height          =   270
         Left            =   60
         TabIndex        =   34
         Top             =   1035
         Width           =   1305
      End
      Begin VB.Label LabelOrgaoArrecadador 
         Caption         =   "Órgão Arrecadador:"
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
         Height          =   270
         Left            =   2805
         TabIndex        =   33
         Top             =   645
         Width           =   1740
      End
      Begin VB.Label LabelNumero 
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
         Height          =   180
         Left            =   615
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datas"
      Height          =   1095
      Left            =   75
      TabIndex        =   27
      Top             =   2160
      Width           =   6930
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   285
         Left            =   1365
         TabIndex        =   10
         Top             =   675
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataVencimento 
         Height          =   300
         Left            =   2340
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEntrega 
         Height          =   300
         Left            =   4605
         TabIndex        =   8
         Top             =   240
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataEntrega 
         Height          =   300
         Left            =   5625
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2325
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1350
         TabIndex        =   6
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelData 
         Caption         =   "Data:"
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
         Height          =   240
         Left            =   840
         TabIndex        =   30
         Top             =   285
         Width           =   525
      End
      Begin VB.Label LabelDataEntrega 
         Caption         =   "Data Entrega:"
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
         Left            =   3300
         TabIndex        =   29
         Top             =   285
         Width           =   1260
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
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
         Left            =   255
         TabIndex        =   28
         Top             =   705
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "SPED"
      Height          =   1020
      Left            =   105
      TabIndex        =   26
      Top             =   4095
      Width           =   6900
      Begin VB.Frame Frame6 
         Caption         =   "Obrigação a Recolher"
         Height          =   660
         Left            =   195
         TabIndex        =   37
         Top             =   225
         Width           =   2910
         Begin MSMask.MaskEdBox CodObrigRecolher 
            Height          =   315
            Left            =   1140
            TabIndex        =   16
            Top             =   255
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Format          =   "000"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   390
            TabIndex        =   38
            Top             =   315
            Width           =   645
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Receita"
         Height          =   660
         Left            =   3330
         TabIndex        =   36
         Top             =   240
         Width           =   3465
         Begin MSMask.MaskEdBox CodReceita 
            Height          =   315
            Left            =   1275
            TabIndex        =   17
            Top             =   240
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   555
            TabIndex        =   39
            Top             =   300
            Width           =   645
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4800
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GuiaICMSOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "GuiaICMSOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "GuiaICMSOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "GuiaICMSOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de Apuração"
      Height          =   675
      Left            =   90
      TabIndex        =   23
      Top             =   3345
      Width           =   6915
      Begin MSComCtl2.UpDown UpDownApuracaoDe 
         Height          =   300
         Left            =   2295
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox ApuracaoDe 
         Height          =   300
         Left            =   1335
         TabIndex        =   12
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownApuracaoAte 
         Height          =   300
         Left            =   5580
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox ApuracaoAte 
         Height          =   300
         Left            =   4590
         TabIndex        =   14
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelApuracaoDe 
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
         Height          =   240
         Left            =   915
         TabIndex        =   22
         Top             =   300
         Width           =   345
      End
      Begin VB.Label LabelApuracaoAte 
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
         Height          =   240
         Left            =   4140
         TabIndex        =   24
         Top             =   315
         Width           =   525
      End
   End
End
Attribute VB_Name = "GuiaICMSOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis do Browse
Private WithEvents objEventoBotaoConsulta As AdmEvento
Attribute objEventoBotaoConsulta.VB_VarHelpID = -1

Dim giICMSST As Integer

'Variáveis globais
Dim iAlterado As Integer

'*** FUNÇÕES DE INICIALIZAÇÃO DA TELA - INÍCIO ***

Public Sub Form_Load()
'Função que carregará a tela

Dim lErro As Long
Dim objEstado As New ClassEstado
Dim colEstado As New Collection

On Error GoTo Erro_Form_Load

    Set objEventoBotaoConsulta = New AdmEvento

    Call CarregaTela
    
    'Le os estados
    lErro = CF("Estados_Le_Todos", colEstado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Preenche a combo de Estados
    For Each objEstado In colEstado
        UF.AddItem objEstado.sSigla
    Next
    
    UF.ListIndex = 0
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161747)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(Optional objGuiasICMS As ClassGuiasICMS, Optional ByVal iICMSST As Long = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giICMSST = iICMSST
    
    'Se foi passado um Item como parâmetro
    If Not objGuiasICMS Is Nothing Then
    
        If Len(Trim(objGuiasICMS.sUF)) <> 0 Then
            objGuiasICMS.iICMSST = MARCADO
            giICMSST = objGuiasICMS.iICMSST
        End If
    
        '??? falta tratar o caso da guia ser passada
        lErro = CF("GuiaICMS_Le", objGuiasICMS)
        If lErro <> SUCESSO And lErro <> 125236 Then gError 125332
        
        If lErro = 125236 Then gError 125333
        
        lErro = Traz_GuiaICMS_Tela(objGuiasICMS)
        If lErro <> SUCESSO Then gError 125334
        
    End If
    
    If giICMSST = MARCADO Then
        Caption = "Guia ICMS ST"
        FrameUF.Visible = True
    Else
        Caption = "Guia ICMS"
    End If
        
    Trata_Parametros = SUCESSO
       
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 125332, 125334
         
        Case 125333
            Call Rotina_Erro(vbOKOnly, "ERRO_GUIASICMS_NAO_ENCONTRADO", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161748)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE INICIALIZAÇÃO DA TELA - FIM ***

'*** EVENTOS CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o conteúdo da tela no bd
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 125209

    Call Limpa_GuiaICMS_Tela

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161749)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objGuiasICMS As New ClassGuiasICMS

On Error GoTo Erro_BotaoExcluir_Click

    If Len(Trim(Data.ClipText)) = 0 Then gError 125211

    objGuiasICMS.sNumero = Trim(Numero.Text)
    objGuiasICMS.dtData = Data.Text
    objGuiasICMS.iFilialEmpresa = giFilialEmpresa
    If giICMSST = MARCADO Then objGuiasICMS.sUF = UF.Text
    objGuiasICMS.iICMSST = giICMSST
    
    'Pede a confirmação da exclusão
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_GUIAICMS", objGuiasICMS.sNumero, objGuiasICMS.dtData)
    
    'se a resposta for não
    If vbMsgResp = vbNo Then Exit Sub

    'transforma o mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    lErro = CF("GuiaICMS_Exclui", objGuiasICMS)
    If lErro <> SUCESSO And lErro <> 125246 Then gError 125214

    If lErro = 125246 Then gError 125215
    
    Call Limpa_GuiaICMS_Tela
    
    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
                
        Case 125210
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_GUIASICMS_NAO_PREECHIDO", gErr)
                
        Case 125211
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_GUIASICMS_NAO_PREENCHIDO", gErr)
        
        Case 125215
            Call Rotina_Erro(vbOKOnly, "ERRO_GUIASICMS_NAO_ENCONTRADO", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
                
        Case 125214
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161750)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa alterações
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 125216

    'limpa a tela toda
    Call Limpa_GuiaICMS_Tela
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 125216
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161751)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Call Unload(Me)

End Sub

Private Sub UpDownApuracaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownApuracaoDe_DownClick

    'Se a data está preenchida
    If Len(Trim(ApuracaoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(ApuracaoDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 125867

    End If

    Exit Sub

Erro_UpDownApuracaoDe_DownClick:

    Select Case gErr

        Case 125867

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161752)

    End Select

    Exit Sub

End Sub

Private Sub UpDownApuracaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownApuracaoDe_UpClick

    'Se a data está preenchida
    If Len(Trim(ApuracaoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(ApuracaoDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 125868

    End If

    Exit Sub

Erro_UpDownApuracaoDe_UpClick:

    Select Case gErr

        Case 125868

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161753)

    End Select

    Exit Sub

End Sub

Private Sub UpDownApuracaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownApuracaoAte_DownClick

    'Se a data está preenchida
    If Len(Trim(ApuracaoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(ApuracaoAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 125869

    End If

    Exit Sub

Erro_UpDownApuracaoAte_DownClick:

    Select Case gErr

        Case 125869

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161754)

    End Select

    Exit Sub

End Sub

Private Sub UpDownApuracaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownApuracaoAte_UpClick

    'Se a data está preenchida
    If Len(Trim(ApuracaoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(ApuracaoAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 125870

    End If

    Exit Sub

Erro_UpDownApuracaoAte_UpClick:

    Select Case gErr

        Case 125870

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161755)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 125217

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 125217

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161756)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 125218

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 125218

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161757)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_DownClick

    'Se a data está preenchida
    If Len(Trim(DataEntrega.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataEntrega, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 125219

    End If

    Exit Sub

Erro_UpDownDataEntrega_DownClick:

    Select Case gErr

        Case 125219

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161758)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_UpClick

    'Se a data está preenchida
    If Len(Trim(DataEntrega.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataEntrega, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 125220

    End If

    Exit Sub

Erro_UpDownDataEntrega_UpClick:

    Select Case gErr

        Case 125220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161759)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objGuiasICMS As New ClassGuiasICMS
Dim colSelecao As Collection

On Error GoTo Erro_BotaoConsulta_Click

    If giICMSST = MARCADO Then
        Call Chama_Tela("GuiasICMSSTLista", colSelecao, objGuiasICMS, objEventoBotaoConsulta)
    Else
        Call Chama_Tela("GuiaICMSLista", colSelecao, objGuiasICMS, objEventoBotaoConsulta)
    End If
    
    Exit Sub
    
Erro_BotaoConsulta_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161760)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim lErro As Long
Dim objGuiasICMS As New ClassGuiasICMS
Dim colSelecao As Collection

On Error GoTo Erro_LabelNumero_Click

    If giICMSST = MARCADO Then
        Call Chama_Tela("GuiasICMSSTLista", colSelecao, objGuiasICMS, objEventoBotaoConsulta)
    Else
        Call Chama_Tela("GuiaICMSLista", colSelecao, objGuiasICMS, objEventoBotaoConsulta)
    End If
    
    Exit Sub
    
Erro_LabelNumero_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161761)

    End Select

    Exit Sub

End Sub
'*** EVENTOS CLICK DOS CONTROLES - FIM ***

'*** EVENTOS CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LocalEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OrgaoArrecadador_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ApuracaoDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ApuracaoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UF_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTOS CHANGE DOS CONTROLES - FIM ***

'*** EVENTOS VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Critica o valor data
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 125221

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 125221
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161762)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ApuracaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ApuracaoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(ApuracaoDe.Text)
    If lErro <> SUCESSO Then gError 125871

    Exit Sub

Erro_ApuracaoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 125871
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161763)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ApuracaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ApuracaoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(ApuracaoAte.Text)
    If lErro <> SUCESSO Then gError 125872

    Exit Sub

Erro_ApuracaoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 125872
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161764)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrega_Validate

    'Critica o valor data
    lErro = Data_Critica(DataEntrega.Text)
    If lErro <> SUCESSO Then gError 125222

    Exit Sub

Erro_DataEntrega_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 125222
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161765)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor.Text)) = 0 Then Exit Sub
    
    'Critica o valor informado
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then gError 125223
    
    Exit Sub
    
Erro_Valor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 125223
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161766)
            
    End Select
    
    Exit Sub
    
End Sub
'*** EVENTOS VALIDATE DOS CONTROLES - FIM ***

'*** FUNÇÕES DE APOIO A TELA - INÍCIO ***
Private Function CarregaTela()

    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrega.Text = Format(gdtDataAtual, "dd/mm/yy")
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long, objGuiasICMS As New ClassGuiasICMS

On Error GoTo Erro_Gravar_Registro

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Data.ClipText)) = 0 Then gError 125225
    If Len(Trim(DataEntrega.ClipText)) = 0 Then gError 125226
    If Len(Trim(Valor.ClipText)) = 0 Then gError 125227
    If Len(Trim(OrgaoArrecadador.Text)) = 0 Then gError 125228
    If Len(Trim(LocalEntrega)) = 0 Then gError 125229
    If Len(Trim(ApuracaoDe.ClipText)) = 0 Then gError 125874
    If Len(Trim(ApuracaoAte.ClipText)) = 0 Then gError 125875
    
    'data não pode ser maior que a data Entrega --> ERRO
    If Trim(Data.ClipText) <> "" And Trim(DataEntrega.ClipText) <> "" Then
    
         If CDate(Data.Text) > CDate(DataEntrega.Text) Then gError 125230
    
    End If
    
    'data inicial não pode ser maior que a data final --> ERRO
    If Trim(ApuracaoDe.ClipText) <> "" And Trim(ApuracaoAte.ClipText) <> "" Then
    
         If CDate(ApuracaoDe.Text) > CDate(ApuracaoAte.Text) Then gError 125873
    
    End If
    
    lErro = Move_Tela_Memoria(objGuiasICMS)
    If lErro <> SUCESSO Then gError 125231
    
    lErro = CF("GuiaICMS_Grava", objGuiasICMS)
    If lErro <> SUCESSO Then gError 125232
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 125225
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_GUIASICMS_NAO_PREENCHIDO", gErr)
        
        Case 125226
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTREGA_GUIASICMS_NAO_PREENCHIDO", gErr)
        
        Case 125227
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_GUIASICMS_NAO_PREENCHIDO", gErr)
        
        Case 125228
            Call Rotina_Erro(vbOKOnly, "ERRO_ORGAOARRECAD_GUIASICMAS_NAO_PREENCHIDO", gErr)
            
        Case 125229
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALENTRADA_GUIASICMS_NAO_PREENCHIDO", gErr)
        
        Case 125230
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_MAIOR_DATAENTREGA_GUIAICMS", gErr)
        
        Case 125231
        
        Case 125232
        
        Case 125873
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_INICIAL_MAIOR", gErr)
        
        Case 125874
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_INICIAL_NAO_PREENCHIDA", gErr)
        
        Case 125875
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_FINAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161767)
            
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objGuiasICMS As ClassGuiasICMS) As Long

On Error GoTo Erro_Move_Tela_Memoria

    objGuiasICMS.dtData = MaskedParaDate(Data)
    objGuiasICMS.dtDataEntrega = MaskedParaDate(DataEntrega)
    objGuiasICMS.dValor = StrParaDbl(Valor.ClipText)
    objGuiasICMS.sLocalEntrega = LocalEntrega.Text
    objGuiasICMS.sNumero = Trim(Numero.ClipText)
    objGuiasICMS.sOrgaoArrecadador = OrgaoArrecadador.Text
    objGuiasICMS.iFilialEmpresa = giFilialEmpresa
    objGuiasICMS.dtVencimento = MaskedParaDate(DataVencimento)
    objGuiasICMS.sCodReceita = CodReceita.Text 'Format(CodReceita.Text, CodReceita.Format)
    objGuiasICMS.sCodObrigRecolher = Format(CodObrigRecolher.Text, CodObrigRecolher.Format)
    If giICMSST = MARCADO Then objGuiasICMS.sUF = UF.Text
    
    objGuiasICMS.dtApuracaoDe = MaskedParaDate(ApuracaoDe)
    objGuiasICMS.dtApuracaoAte = MaskedParaDate(ApuracaoAte)
    
    objGuiasICMS.iICMSST = giICMSST
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161768)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Limpa_GuiaICMS_Tela()

    Call Limpa_Tela(Me)
    
    Call CarregaTela
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
End Sub

Private Function Traz_GuiaICMS_Tela(objGuiasICMS As ClassGuiasICMS) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_GuiaICMS_Tela

    Numero.PromptInclude = False
    Numero.Text = objGuiasICMS.sNumero
    Numero.PromptInclude = True
    
    Call DateParaMasked(Data, objGuiasICMS.dtData)
    Call DateParaMasked(DataEntrega, objGuiasICMS.dtDataEntrega)
    
    Call DateParaMasked(ApuracaoDe, objGuiasICMS.dtApuracaoDe)
    Call DateParaMasked(ApuracaoAte, objGuiasICMS.dtApuracaoAte)
    
    Valor.PromptInclude = False
    Valor.Text = Format(objGuiasICMS.dValor, "#,##0.00")
    Valor.PromptInclude = True
    
    LocalEntrega.Text = objGuiasICMS.sLocalEntrega
    OrgaoArrecadador.Text = objGuiasICMS.sOrgaoArrecadador
    
    Call DateParaMasked(DataVencimento, objGuiasICMS.dtVencimento)
    CodReceita.PromptInclude = False
    CodReceita.Text = objGuiasICMS.sCodReceita 'Format(objGuiasICMS.sCodReceita, CodReceita.Format)
    CodReceita.PromptInclude = True
    CodObrigRecolher.PromptInclude = False
    CodObrigRecolher.Text = Format(objGuiasICMS.sCodObrigRecolher, CodObrigRecolher.Format)
    CodObrigRecolher.PromptInclude = True
    
    If objGuiasICMS.iICMSST = MARCADO Then Call CF("sCombo_Seleciona2", UF, objGuiasICMS.sUF)
    
    iAlterado = 0
    
    Traz_GuiaICMS_Tela = SUCESSO
    
    Exit Function

Erro_Traz_GuiaICMS_Tela:

    Traz_GuiaICMS_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161769)
            
    End Select
    
    Exit Function
    
End Function
'*** FUNÇÕES DE APOIO A TELA - FIM ***

'*** FUNÇÕES DO SISTEMA DE SETA ***
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objGuiasICMS As New ClassGuiasICMS

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objGuiasICMS.sNumero = colCampoValor.Item("Numero").vValor
    objGuiasICMS.dtData = CStr(colCampoValor.Item("Data").vValor)
    objGuiasICMS.dtDataEntrega = CStr(colCampoValor.Item("DataEntrega").vValor)
    objGuiasICMS.dValor = CStr(colCampoValor.Item("Valor").vValor)
    objGuiasICMS.sLocalEntrega = colCampoValor.Item("LocalEntrega").vValor
    objGuiasICMS.sOrgaoArrecadador = colCampoValor.Item("OrgaoArrecadador").vValor
    objGuiasICMS.dtApuracaoDe = CStr(colCampoValor.Item("ApuracaoDe").vValor)
    objGuiasICMS.dtApuracaoAte = CStr(colCampoValor.Item("ApuracaoAte").vValor)
    objGuiasICMS.iFilialEmpresa = giFilialEmpresa
    objGuiasICMS.iICMSST = giICMSST
    If giICMSST = MARCADO Then objGuiasICMS.sUF = colCampoValor.Item("UF").vValor

    lErro = CF("GuiaICMS_Le", objGuiasICMS)
    If lErro <> SUCESSO And lErro <> 125236 Then gError 125332

    lErro = Traz_GuiaICMS_Tela(objGuiasICMS)
    If lErro <> SUCESSO Then gError 125249

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 125249, 125332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161770)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objGuiasICMS As New ClassGuiasICMS

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    If giICMSST = MARCADO Then
        sTabela = "GuiasICMSST"
    Else
        sTabela = "GuiasICMS"
    End If

    lErro = Move_Tela_Memoria(objGuiasICMS)
    If lErro <> SUCESSO Then gError 125250

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Numero", objGuiasICMS.sNumero, STRING_GUIASICMS_NUMERO, "Numero"
    colCampoValor.Add "Data", objGuiasICMS.dtData, 0, "Data"
    colCampoValor.Add "DataEntrega", objGuiasICMS.dtDataEntrega, 0, "DataEntrega"
    colCampoValor.Add "Valor", objGuiasICMS.dValor, 0, "Valor"
    colCampoValor.Add "LocalEntrega", objGuiasICMS.sLocalEntrega, STRING_GUIASICMS_LOCALENTREGA, "LocalEntrega"
    colCampoValor.Add "OrgaoArrecadador", objGuiasICMS.sOrgaoArrecadador, STRING_GUIASICMS_ORGAOARRECADADOR, "OrgaoArrecadador"
    colCampoValor.Add "ApuracaoDe", objGuiasICMS.dtApuracaoDe, 0, "ApuracaoDe"
    colCampoValor.Add "ApuracaoAte", objGuiasICMS.dtApuracaoAte, 0, "ApuracaoAte"
    If giICMSST = MARCADO Then colCampoValor.Add "UF", objGuiasICMS.sUF, STRING_UM_SIGLA, "UF"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 125250

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161771)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub
'*** FUNÇÕES DO SISTEMA DE SETA - FIM ***

'*** FUNÇÕES DO BROWSE - INÍCIO ***
Private Sub objEventoBotaoConsulta_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objGuiasICMS As New ClassGuiasICMS

On Error GoTo Erro_objEventoBotaoConsulta_evSelecao

    Set objGuiasICMS = obj1
    
    objGuiasICMS.iICMSST = giICMSST
    
    lErro = CF("GuiaICMS_Le", objGuiasICMS)
    If lErro <> SUCESSO And lErro <> 125236 Then gError 125332
    
    lErro = Traz_GuiaICMS_Tela(objGuiasICMS)
    If lErro <> SUCESSO Then gError 125251
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoConsulta_evSelecao:

    Select Case gErr
        
        Case 125251, 125332
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161772)

    End Select

    Exit Sub
    
End Sub
'*** FUNÇÕES DO BROWSE - FIM ***

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoBotaoConsulta = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    If giICMSST = MARCADO Then
        Caption = "Guia ICMS ST"
    Else
        Caption = "Guia ICMS"
    End If
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GuiaICMS"
    
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
'**** fim do trecho a ser copiado *****

Private Sub CodObrigRecolher_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodReceita_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodObrigRecolher_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodObrigRecolher, iAlterado)
End Sub

Private Sub DataVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataVencimento_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 70078

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True

    Select Case gErr

        Case 70078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144051)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Se a data está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70079

    End If

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case gErr

        Case 70079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144052)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Se a data está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70080

    End If

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case gErr

        Case 70080

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144053)

    End Select

    Exit Sub

End Sub

