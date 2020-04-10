VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl TRVGerComiInt 
   ClientHeight    =   5052
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   ScaleHeight     =   5055.189
   ScaleMode       =   0  'User
   ScaleWidth      =   6503.613
   Begin VB.Frame Frame3 
      Caption         =   "Prévia"
      Height          =   1044
      Left            =   96
      TabIndex        =   28
      Top             =   3876
      Width           =   6048
      Begin VB.CommandButton BotaoPlanilhaPrevia 
         Caption         =   "Planilha"
         Height          =   735
         Left            =   4136
         Picture         =   "TRVGerComiInt.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Consultar"
         Top             =   204
         Width           =   825
      End
      Begin VB.CommandButton BotaoExcluirPrevia 
         Caption         =   "Excluir"
         Height          =   735
         Left            =   5124
         Picture         =   "TRVGerComiInt.ctx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Consultar"
         Top             =   204
         Width           =   825
      End
      Begin VB.CommandButton BotaoGerarPrevia 
         Caption         =   "Gerar"
         Height          =   735
         Left            =   3148
         Picture         =   "TRVGerComiInt.ctx":159C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Consultar"
         Top             =   204
         Width           =   825
      End
      Begin VB.CommandButton BotaoConsultarPrevia 
         Caption         =   "Consultar"
         Height          =   735
         Left            =   2160
         Picture         =   "TRVGerComiInt.ctx":206A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Consultar"
         Top             =   204
         Width           =   825
      End
      Begin MSMask.MaskEdBox Previa 
         Height          =   300
         Left            =   996
         TabIndex        =   33
         Top             =   408
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPrevia 
         AutoSize        =   -1  'True
         Caption         =   "Prévia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   336
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   432
         Width           =   588
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2265
      Picture         =   "TRVGerComiInt.ctx":2B38
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Geração"
      Height          =   825
      Left            =   120
      TabIndex        =   19
      Top             =   2910
      Width           =   6105
      Begin VB.Label Hora 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5100
         TabIndex        =   25
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3240
         TabIndex        =   24
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Usuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1080
         TabIndex        =   23
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   4590
         TabIndex        =   22
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2685
         TabIndex        =   21
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   20
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos vouchers"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1965
      Width           =   6105
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   300
         Left            =   2175
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissaoDe 
         Height          =   300
         Left            =   1050
         TabIndex        =   13
         Top             =   330
         Width           =   1170
         _ExtentX        =   2053
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissaoAte 
         Height          =   300
         Left            =   3225
         TabIndex        =   14
         Top             =   315
         Width           =   1170
         _ExtentX        =   2053
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   300
         Left            =   4395
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   17
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2820
         TabIndex        =   16
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Período de pagamento dos vouchers"
      Height          =   795
      Left            =   105
      TabIndex        =   4
      Top             =   900
      Width           =   6105
      Begin MSComCtl2.UpDown UpDownBaixaDe 
         Height          =   300
         Left            =   2175
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   345
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixaDe 
         Height          =   300
         Left            =   1050
         TabIndex        =   6
         Top             =   345
         Width           =   1170
         _ExtentX        =   2053
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataBaixaAte 
         Height          =   300
         Left            =   3225
         TabIndex        =   7
         Top             =   330
         Width           =   1170
         _ExtentX        =   2053
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownBaixaAte 
         Height          =   300
         Left            =   4395
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2820
         TabIndex        =   10
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   9
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4425
      ScaleHeight     =   504
      ScaleWidth      =   1680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1725
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   105
         Picture         =   "TRVGerComiInt.ctx":2C22
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gerar o MRP"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TRVGerComiInt.ctx":3064
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "TRVGerComiInt.ctx":31EE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1170
      TabIndex        =   27
      Top             =   255
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   285
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   315
      Width           =   810
   End
End
Attribute VB_Name = "TRVGerComiInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private WithEvents objEventoGerComissao As AdmEvento
Attribute objEventoGerComissao.VB_VarHelpID = -1
Private WithEvents objEventoRelComissao As AdmEvento
Attribute objEventoRelComissao.VB_VarHelpID = -1
Private WithEvents objEventoRelGerComiInt As AdmEvento
Attribute objEventoRelGerComiInt.VB_VarHelpID = -1

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Padrao_Tela()

Dim iMes As Integer
Dim iAno As Integer
Dim dtData As Date
Dim dtData1 As Date

On Error GoTo Erro_Padrao_Tela

    iMes = Month(gdtDataAtual)
    iAno = Year(gdtDataAtual)

    dtData = CDate("20/" & iMes & "/" & iAno)

    DataBaixaAte.PromptInclude = False
    DataBaixaAte.Text = Format(dtData, "dd/mm/yy")
    DataBaixaAte.PromptInclude = True

    dtData = CDate("1/" & iMes & "/" & iAno)
    
    dtData1 = DateAdd("d", -1, dtData)
    
    DataEmissaoAte.PromptInclude = False
    DataEmissaoAte.Text = Format(dtData1, "dd/mm/yy")
    DataEmissaoAte.PromptInclude = True
    
    Exit Sub

Erro_Padrao_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select


    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoGerComissao = New AdmEvento
    Set objEventoRelComissao = New AdmEvento
    Set objEventoRelGerComiInt = New AdmEvento
    

    Call Padrao_Tela

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 197259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoGerComissao = Nothing
    Set objEventoRelComissao = Nothing
    Set objEventoRelGerComiInt = Nothing
    

    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
   Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Comissão Interna"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVGerComiInt"

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


Private Sub BotaoPrevia_Click()

Dim iRelatorio As Integer
Dim objTRVGerComiInt As New ClassTRVGerComiInt
Dim lErro As Long
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoPrevia_Click

    lErro = Formata_E_Critica_Dados(objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197477
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 197478
    
    iRelatorio = 1
    
    lErro = CF("Rotina_GerComiInt", sNomeArqParam, objTRVGerComiInt, iRelatorio)
    If lErro <> SUCESSO Then gError 197479
        
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoPrevia_Click:

    Select Case gErr

        Case 197477 To 197479
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197480)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultarPrevia_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim lNumIntRel As Long
Dim objTRVRelComissao As New ClassTRVRelComissao
Dim sFiltro As String

On Error GoTo Erro_BotaoConsultarPrevia_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Previa.Text)) <> 0 Then

        lNumIntRel = StrParaLong(Previa.Text)

    End If
    
    sFiltro = "NumIntRel=" & lNumIntRel

    Call Chama_Tela("TRVRelComissaoLista", colSelecao, objTRVRelComissao, objEventoRelComissao, sFiltro)

    Exit Sub

Erro_BotaoConsultarPrevia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197289)

    End Select

    Exit Sub

End Sub


Private Sub BotaoGerarPrevia_Click()

Dim iRelatorio As Integer
Dim objTRVGerComiInt As New ClassTRVGerComiInt
Dim lErro As Long
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoGerarPrevia_Click

    lErro = Formata_E_Critica_Dados(objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197481
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 197482
    
    iRelatorio = 2
    
    Set objTRVGerComiInt.objTela = Me
    
    lErro = CF("Rotina_GerComiInt", sNomeArqParam, objTRVGerComiInt, iRelatorio)
    If lErro <> SUCESSO Then gError 197483
        
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGerarPrevia_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 197481 To 197483
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197484)

    End Select

    Exit Sub

End Sub



Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 197285

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 197285
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197286)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "TRVConfig", "NUM_PROX_TRVGERACOMISSAOINT", "TRVGerComiInt", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 197287
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 197287

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197288)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DataEmissaoDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 197261

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 197261

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197262)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 197263

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197264)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixaDe_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataBaixaDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataBaixaDe, iAlterado)

End Sub

Private Sub DataBaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBaixaDe_Validate

    'Verifica se a Data de Baixa foi digitada
    If Len(Trim(DataBaixaDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataBaixaDe.Text)
    If lErro <> SUCESSO Then gError 197265

    Exit Sub

Erro_DataBaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 197265

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197266)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixaAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataBaixaAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataBaixaAte, iAlterado)

End Sub

Private Sub DataBaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBaixaAte_Validate

    'Verifica se a Data de Baixa foi digitada
    If Len(Trim(DataBaixaAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataBaixaAte.Text)
    If lErro <> SUCESSO Then gError 197267

    Exit Sub

Erro_DataBaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197267

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197268)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197269

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 197269

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197270)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197271

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 197271

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197272)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197277

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 197277

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197279

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 197279

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197280)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownBaixaDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataBaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197273

    Exit Sub

Erro_UpDownBaixaDe_DownClick:

    Select Case gErr

        Case 197273

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197274)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataBaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197275

    Exit Sub

Erro_UpDownBaixaDe_UpClick:

    Select Case gErr

        Case 197275

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197276)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownBaixaAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataBaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197281

    Exit Sub

Erro_UpDownBaixaAte_DownClick:

    Select Case gErr

        Case 197281

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197282)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataBaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197283

    Exit Sub

Erro_UpDownBaixaAte_UpClick:

    Select Case gErr

        Case 197283

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197284)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTRVGerComiInt As New ClassTRVGerComiInt
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTRVGerComiInt.lCodigo = StrParaLong(Codigo.Text)

    End If

    Call Chama_Tela("TRVGerComiIntLista", colSelecao, objTRVGerComiInt, objEventoGerComissao)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197289)

    End Select

    Exit Sub

End Sub

Private Sub objEventoGerComissao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRVGerComiInt As ClassTRVGerComiInt

On Error GoTo Erro_objEventoGerComissao_evSelecao

    Set objTRVGerComiInt = obj1

    'Mostra os dados do TRVGerComiInt na tela
    lErro = Traz_TRVGerComiInt_Tela(objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197290

    Me.Show

    Exit Sub

Erro_objEventoGerComissao_evSelecao:

    Select Case gErr

        Case 197290

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197291)

    End Select

    Exit Sub

End Sub

Function Traz_TRVGerComiInt_Tela(objTRVGerComiInt As ClassTRVGerComiInt) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_TRVGerComiInt_Tela

    Call Limpa_Tela_TRVGerComiInt
    
    If objTRVGerComiInt.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = objTRVGerComiInt.lCodigo
        Codigo.PromptInclude = True
    End If

    'Lê o TRVGerComiInt que está sendo Passado
    lErro = CF("TRVGerComiInt_Le", objTRVGerComiInt)
    If lErro <> SUCESSO And lErro <> 197295 Then gError 197297
    
    If lErro = SUCESSO Then
        
        If objTRVGerComiInt.dtDataPagtoDe <> DATA_NULA Then
            DataBaixaDe.PromptInclude = False
            DataBaixaDe.Text = Format(objTRVGerComiInt.dtDataPagtoDe, "dd/mm/yy")
            DataBaixaDe.PromptInclude = True
        End If

        If objTRVGerComiInt.dtDataPagtoAte <> DATA_NULA Then
            DataBaixaAte.PromptInclude = False
            DataBaixaAte.Text = Format(objTRVGerComiInt.dtDataPagtoAte, "dd/mm/yy")
            DataBaixaAte.PromptInclude = True
        End If

        If objTRVGerComiInt.dtDataEmiDe <> DATA_NULA Then
            DataEmissaoDe.PromptInclude = False
            DataEmissaoDe.Text = Format(objTRVGerComiInt.dtDataEmiDe, "dd/mm/yy")
            DataEmissaoDe.PromptInclude = True
        End If

        If objTRVGerComiInt.dtDataEmiAte <> DATA_NULA Then
            DataEmissaoAte.PromptInclude = False
            DataEmissaoAte.Text = Format(objTRVGerComiInt.dtDataEmiAte, "dd/mm/yy")
            DataEmissaoAte.PromptInclude = True
        End If


        Usuario.Caption = objTRVGerComiInt.sUsuario
        
        Data.Caption = Format(objTRVGerComiInt.dtDataGeracao, "dd/mm/yyyy")
        
        Hora.Caption = Format(objTRVGerComiInt.dHoraGeracao, "hh:mm:ss")
        
    End If

    iAlterado = 0

    Traz_TRVGerComiInt_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVGerComiInt_Tela:

    Traz_TRVGerComiInt_Tela = gErr

    Select Case gErr

        Case 197297

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197298)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRVGerComiInt

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 197299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197300)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TRVGerComiInt()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVGerComiInt

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
          
    Usuario.Caption = ""
    Data.Caption = ""
    Hora.Caption = ""
    
    Call Padrao_Tela
    
    iAlterado = 0
 
    Exit Sub

Erro_Limpa_Tela_TRVGerComiInt:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197301)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Previa Then
            Call LabelPrevia_Click
        End If
          
    End If

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

Private Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objTRVGerComiInt As New ClassTRVGerComiInt
    
On Error GoTo Erro_BotaoExcluir_Click
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 197419

    objTRVGerComiInt.lCodigo = StrParaLong(Codigo.Text)
    
    lErro = CF("TRVGerComiInt_Exclui", objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197420
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 197419
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 197420
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197420)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim objTRVGerComiInt As New ClassTRVGerComiInt

On Error GoTo Erro_BotaoGerar_Click

    lErro = Formata_E_Critica_Dados(objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197304
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 197305
    
    lErro = CF("Rotina_GerComiInt", sNomeArqParam, objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 197306
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Limpa_Tela_TRVGerComiInt
    
    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 197304 To 197306
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197307)

    End Select

    Exit Sub

End Sub

Public Function Formata_E_Critica_Dados(objTRVGerComiInt As ClassTRVGerComiInt) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Dados

    If Len(Trim(Codigo.Text)) = 0 Then gError 197308

    objTRVGerComiInt.lCodigo = StrParaLong(Codigo.Text)

    'Lê o TRVGerComiInt que está sendo Passado
    lErro = CF("TRVGerComiInt_Le", objTRVGerComiInt)
    If lErro <> SUCESSO And lErro <> 197895 Then gError 197309

    If lErro = SUCESSO Then gError 197310

    If Len(Trim(DataEmissaoDe.Text)) > 0 And Len(Trim(DataEmissaoAte.Text)) > 0 Then
        If StrParaDate(DataEmissaoDe.Text) > StrParaDate(DataEmissaoAte.Text) Then gError 197311
    End If
    
    If Len(Trim(DataBaixaDe.Text)) > 0 And Len(Trim(DataBaixaAte.Text)) > 0 Then
        If StrParaDate(DataBaixaDe.Text) > StrParaDate(DataBaixaAte.Text) Then gError 197312
    End If
    
    objTRVGerComiInt.lCodigo = StrParaLong(Codigo.Text)
    objTRVGerComiInt.sUsuario = gsUsuario
    objTRVGerComiInt.dtDataGeracao = gdtDataAtual
    objTRVGerComiInt.dHoraGeracao = Time
    objTRVGerComiInt.dtDataPagtoDe = StrParaDate(DataBaixaDe.Text)
    objTRVGerComiInt.dtDataPagtoAte = StrParaDate(DataBaixaAte.Text)
    objTRVGerComiInt.dtDataEmiDe = StrParaDate(DataEmissaoDe.Text)
    objTRVGerComiInt.dtDataEmiAte = StrParaDate(DataEmissaoAte.Text)

    Formata_E_Critica_Dados = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Dados:

    Formata_E_Critica_Dados = gErr

    Select Case gErr
        
        Case 197308
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 197309
        
        Case 197310
            Call Rotina_Erro(vbOKOnly, "ERRO_GERCOMIINT_CODIGO_EXISTENTE", gErr, objTRVGerComiInt.lCodigo)

        Case 197311
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAODE_MAIOR_DATAEMISSAOATE", gErr, DataEmissaoDe.Text, DataEmissaoAte.Text)
        
        Case 197312
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPAGTODE MAIOR DATAPAGTOATE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197313)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVGerComiInt"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", StrParaLong(Codigo.Text), 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197454)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTRVGerComiInt As New ClassTRVGerComiInt

On Error GoTo Erro_Tela_Preenche

    objTRVGerComiInt.lCodigo = colCampoValor.Item("Codigo").vValor

    If objTRVGerComiInt.lCodigo <> 0 Then
    
        lErro = Traz_TRVGerComiInt_Tela(objTRVGerComiInt)
        If lErro <> SUCESSO Then gError 197455
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 197455

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197456)

    End Select

    Exit Function

End Function

Private Sub LabelPrevia_Click()

Dim lErro As Long
Dim objTRVRelGerComiInt As New ClassTRVRelGerComiInt
Dim colSelecao As New Collection

On Error GoTo Erro_LabelPrevia_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTRVRelGerComiInt.lNumIntRel = StrParaLong(Previa.Text)

    End If

    Call Chama_Tela("TRVRelGerComiIntLista", colSelecao, objTRVRelGerComiInt, objEventoRelGerComiInt)

    Exit Sub

Erro_LabelPrevia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199489)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRelGerComiInt_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRVRelGerComiInt As ClassTRVRelGerComiInt

On Error GoTo Erro_objEventoRelGerComiInt_evSelecao

    Set objTRVRelGerComiInt = obj1

    'Mostra os dados do TRVGerComiInt na tela
    lErro = Traz_TRVRelGerComiInt_Tela(objTRVRelGerComiInt)
    If lErro <> SUCESSO Then gError 199490

    Me.Show

    Exit Sub

Erro_objEventoRelGerComiInt_evSelecao:

    Select Case gErr

        Case 199490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199491)

    End Select

    Exit Sub

End Sub

Function Traz_TRVRelGerComiInt_Tela(objTRVRelGerComiInt As ClassTRVRelGerComiInt) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_TRVRelGerComiInt_Tela

    Call Limpa_Tela_TRVGerComiInt
    
    
    'Lê o TRVGerComiInt que está sendo Passado
    lErro = CF("TRVRelGerComiInt_Le", objTRVRelGerComiInt)
    If lErro <> SUCESSO And lErro <> 199497 Then gError 199499
    
    If lErro <> SUCESSO Then gError 199500
        
    If objTRVRelGerComiInt.lNumIntRel <> 0 Then
        Previa.PromptInclude = False
        Previa.Text = objTRVRelGerComiInt.lNumIntRel
        Previa.PromptInclude = True
    End If
    
        
    If objTRVRelGerComiInt.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = objTRVRelGerComiInt.lCodigo
        Codigo.PromptInclude = True
    End If

        
    If objTRVRelGerComiInt.dtDataPagtoDe <> DATA_NULA Then
        DataBaixaDe.PromptInclude = False
        DataBaixaDe.Text = Format(objTRVRelGerComiInt.dtDataPagtoDe, "dd/mm/yy")
        DataBaixaDe.PromptInclude = True
    End If

    If objTRVRelGerComiInt.dtDataPagtoAte <> DATA_NULA Then
        DataBaixaAte.PromptInclude = False
        DataBaixaAte.Text = Format(objTRVRelGerComiInt.dtDataPagtoAte, "dd/mm/yy")
        DataBaixaAte.PromptInclude = True
    End If

    If objTRVRelGerComiInt.dtDataEmiDe <> DATA_NULA Then
        DataEmissaoDe.PromptInclude = False
        DataEmissaoDe.Text = Format(objTRVRelGerComiInt.dtDataEmiDe, "dd/mm/yy")
        DataEmissaoDe.PromptInclude = True
    End If

    If objTRVRelGerComiInt.dtDataEmiAte <> DATA_NULA Then
        DataEmissaoAte.PromptInclude = False
        DataEmissaoAte.Text = Format(objTRVRelGerComiInt.dtDataEmiAte, "dd/mm/yy")
        DataEmissaoAte.PromptInclude = True
    End If


    Usuario.Caption = objTRVRelGerComiInt.sUsuario
    
    Data.Caption = Format(objTRVRelGerComiInt.dtDataGeracao, "dd/mm/yyyy")
    
    Hora.Caption = Format(objTRVRelGerComiInt.dHoraGeracao, "hh:mm:ss")
        
    iAlterado = 0

    Traz_TRVRelGerComiInt_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVRelGerComiInt_Tela:

    Traz_TRVRelGerComiInt_Tela = gErr

    Select Case gErr

        Case 199499
        
        Case 199500
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVIA_TRVCOMIINT_NAO_CADASTRADA", gErr, objTRVRelGerComiInt.lNumIntRel)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199501)

    End Select

    Exit Function

End Function

Private Sub Previa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Previa_GotFocus()

    Call MaskEdBox_TrataGotFocus(Previa, iAlterado)

End Sub

Private Sub Previa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Previa_Validate

    If Len(Trim(Previa.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Previa.Text)
    If lErro <> SUCESSO Then gError 199492

    Exit Sub

Erro_Previa_Validate:

    Cancel = True

    Select Case gErr

        Case 199492
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199493)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluirPrevia_Click()

Dim lErro As Long
Dim objTRVRelGerComiInt As New ClassTRVRelGerComiInt
    
On Error GoTo Erro_BotaoExcluirPrevia_Click
    
    If Len(Trim(Previa.Text)) = 0 Then gError 199502

    objTRVRelGerComiInt.lNumIntRel = StrParaLong(Previa.Text)
    
    lErro = CF("TRVRelGerComiInt_Exclui", objTRVRelGerComiInt)
    If lErro <> SUCESSO Then gError 199503
    
    Exit Sub
    
Erro_BotaoExcluirPrevia_Click:

    Select Case gErr

        Case 199502
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVIA_TRVCOMIINT_NAO_PREENCHIDA", gErr)
        
        Case 199503
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199504)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRelComissao_evSelecao(obj1 As Object)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoPlanilhaPrevia_Click()

Dim iRelatorio As Integer
Dim objTRVGerComiInt As New ClassTRVGerComiInt
Dim lErro As Long
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoPlanilhaPrevia_Click

    lErro = Formata_E_Critica_Dados(objTRVGerComiInt)
    If lErro <> SUCESSO Then gError 199616
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 199617
    
    iRelatorio = 3
    
    Set objTRVGerComiInt.objTela = Me
    
    lErro = CF("Rotina_GerComiInt", sNomeArqParam, objTRVGerComiInt, iRelatorio)
    If lErro <> SUCESSO Then gError 199618
        
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoPlanilhaPrevia_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 199616 To 199618
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197484)

    End Select

    Exit Sub

End Sub

