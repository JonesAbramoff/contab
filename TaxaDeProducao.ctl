VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TaxaDeProducao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoVerTaxas 
      Caption         =   "Taxas de Produção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      Picture         =   "TaxaDeProducao.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Abre o Browse das Taxas de Produção cadastradas"
      Top             =   5190
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Taxa de Produção"
      Height          =   2205
      Left            =   135
      TabIndex        =   32
      Top             =   2805
      Width           =   4170
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "TaxaDeProducao.ctx":030A
         Left            =   1395
         List            =   "TaxaDeProducao.ctx":030C
         TabIndex        =   3
         Top             =   360
         Width           =   2625
      End
      Begin VB.ComboBox UMTempo 
         Height          =   315
         Left            =   2895
         TabIndex        =   7
         Top             =   1275
         Width           =   1125
      End
      Begin VB.ComboBox UMProduto 
         Height          =   315
         Left            =   2895
         TabIndex        =   5
         Top             =   840
         Width           =   1125
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Top             =   840
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TempoOperacao 
         Height          =   315
         Left            =   1395
         TabIndex        =   6
         Top             =   1275
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label LabelUMTempo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2460
         TabIndex        =   41
         Top             =   1290
         Width           =   390
      End
      Begin VB.Label LabelUMProduto 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2460
         TabIndex        =   40
         Top             =   870
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Taxa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         TabIndex        =   39
         Top             =   1725
         Width           =   585
      End
      Begin VB.Label LabelTaxaDeProducao 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   22
         Top             =   1695
         Width           =   2625
      End
      Begin VB.Label LabelTipo 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
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
         Height          =   300
         Left            =   600
         TabIndex        =   37
         Top             =   390
         Width           =   660
      End
      Begin VB.Label LabelTempoOperacao 
         Alignment       =   1  'Right Justify
         Caption         =   "Tempo:"
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
         Height          =   315
         Left            =   570
         TabIndex        =   36
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label LabelQuantidade 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   210
         TabIndex        =   33
         Top             =   870
         Width           =   1065
      End
   End
   Begin VB.Frame FrameLotes 
      Caption         =   "Lotes"
      Height          =   2205
      Left            =   7215
      TabIndex        =   29
      Top             =   2805
      Width           =   2160
      Begin MSMask.MaskEdBox LoteMinimo 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   450
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteMaximo 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   990
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LotePadrao 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   1545
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Padrão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   38
         Top             =   1575
         Width           =   735
      End
      Begin VB.Label LabelLoteMin 
         Caption         =   "Mínimo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   31
         Top             =   480
         Width           =   750
      End
      Begin VB.Label LabelLoteMax 
         Caption         =   "Máximo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   30
         Top             =   1020
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tempos (Horas)"
      Height          =   2205
      Left            =   4455
      TabIndex        =   26
      Top             =   2805
      Width           =   2610
      Begin MSMask.MaskEdBox TempoPreparacao 
         Height          =   315
         Left            =   1620
         TabIndex        =   8
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TempoDescarga 
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   1545
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TempoMovimentacao 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   1020
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label LabelTempoMovimentacao 
         Caption         =   "Movimentação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   35
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label LabelTempoPreparacao 
         Caption         =   "Preparação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   28
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label LabelTempoDescarga 
         Caption         =   "Descarga:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   585
         TabIndex        =   27
         Top             =   1575
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   150
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TaxaDeProducao.ctx":030E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TaxaDeProducao.ctx":0468
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TaxaDeProducao.ctx":05F2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TaxaDeProducao.ctx":0B24
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   1395
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Maquina 
      Height          =   315
      Left            =   1515
      TabIndex        =   2
      Top             =   1890
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Competencia 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   885
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Data 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1500
      TabIndex        =   43
      Top             =   2370
      Width           =   1320
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Atualizado em:"
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
      Left            =   135
      TabIndex        =   42
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label LabelCompetencia 
      Alignment       =   1  'Right Justify
      Caption         =   "Competência:"
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
      Height          =   315
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   34
      Top             =   915
      Width           =   1260
   End
   Begin VB.Label DescCompetencia 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3540
      TabIndex        =   19
      Top             =   885
      Width           =   5805
   End
   Begin VB.Label DescMaquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3540
      TabIndex        =   21
      Top             =   1890
      Width           =   5805
   End
   Begin VB.Label LabelProduto 
      Alignment       =   1  'Right Justify
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   25
      Top             =   1425
      Width           =   870
   End
   Begin VB.Label DescProd 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3540
      TabIndex        =   20
      Top             =   1395
      Width           =   5805
   End
   Begin VB.Label LabelMaquina 
      Alignment       =   1  'Right Justify
      Caption         =   "Máquina:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   24
      Top             =   1890
      Width           =   900
   End
End
Attribute VB_Name = "TaxaDeProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoTaxaDeProducao As AdmEvento
Attribute objEventoTaxaDeProducao.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoMaquina As AdmEvento
Attribute objEventoMaquina.VB_VarHelpID = -1
Private WithEvents objEventoCompetencia As AdmEvento
Attribute objEventoCompetencia.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Taxa de Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TaxaDeProducao"

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

Private Sub BotaoVerTaxas_Click()

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMaquinas As ClassMaquinas
Dim objCompetencias As ClassCompetencias
Dim sFiltro As String

On Error GoTo Erro_BotaoVerTaxas_Click

    If Len(Competencia.Text) > 0 Then
    
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = Competencia.Text
        
        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 134464

        objTaxaDeProducao.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    End If

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134465
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objTaxaDeProducao.sProduto = sProdutoFormatado
    
    End If
    
    If Len(Maquina.Text) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 134466
            
        objTaxaDeProducao.lNumIntDocMaq = objMaquinas.lNumIntDoc
     
    End If
    
    sFiltro = "Ativo = ? "
    colSelecao.Add TAXA_ATIVA

    Call Chama_Tela("TaxaDeProducaoLista", colSelecao, objTaxaDeProducao, objEventoTaxaDeProducao, sFiltro)

    Exit Sub

Erro_BotaoVerTaxas_Click:

    Select Case gErr
    
        Case 134464, 134465, 134466
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174500)

    End Select

    Exit Sub

End Sub

Private Sub Competencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Competencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Competencia, iAlterado)
    
End Sub


Private Sub Competencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_Competencia_Validate

    DescCompetencia.Caption = ""

    'Verifica se Competencia não está preenchida
    If Len(Trim(Competencia.Text)) = 0 Then

        Exit Sub
    
    End If
    
    Set objCompetencias = New ClassCompetencias
    
    'Verifica sua existencia
    lErro = CF("TP_Competencia_Le", Competencia, objCompetencias)
    If lErro <> SUCESSO Then gError 134467
    
    DescCompetencia.Caption = objCompetencias.sDescricao
       
    Exit Sub

Erro_Competencia_Validate:

    Cancel = True

    Select Case gErr
    
        Case 134467
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174501)

    End Select

    Exit Sub

End Sub

Private Sub LabelCompetencia_Click()

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCompetencia_Click

    'Verifica se a Competencia foi preenchida
    If Len(Trim(Competencia.Text)) <> 0 Then
    
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = Competencia.Text
        
        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 137921

    End If

    Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompetencia)

    Exit Sub

Erro_LabelCompetencia_Click:

    Select Case gErr
    
        Case 137921

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174502)

    End Select

    Exit Sub

End Sub

Private Sub LabelMaquina_Click()

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelMaquina_Click

    'Verifica se a Maquina foi preenchida
    If Len(Trim(Maquina.Text)) <> 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137922
        
    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaquina)

    Exit Sub

Erro_LabelMaquina_Click:

    Select Case gErr
    
        Case 137922

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174503)

    End Select

    Exit Sub

End Sub

Private Sub LotePadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LotePadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteMaximo, iAlterado)
    
End Sub

Private Sub LotePadrao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LotePadrao_Validate

    'Verifica se LotePadrao está preenchido
    If Len(Trim(LotePadrao.Text)) <> 0 Then

       'Critica a LotePadrao
       lErro = Valor_Positivo_Critica(LotePadrao.Text)
       If lErro <> SUCESSO Then gError 134468
       
       LotePadrao.Text = Formata_Estoque(LotePadrao.Text)

    End If

    Exit Sub

Erro_LotePadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 134468
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174504)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCompetencia_evSelecao(obj1 As Object)

Dim objCompetencias As New ClassCompetencias
Dim lErro As Long

On Error GoTo Erro_objEventoCompetencia_evSelecao

    Set objCompetencias = obj1

    Competencia.Text = objCompetencias.sNomeReduzido
    DescCompetencia.Caption = objCompetencias.sDescricao
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174505)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMaquina_evSelecao(obj1 As Object)

Dim objMaquinas As New ClassMaquinas
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objMaquinas = obj1

    Maquina.Text = objMaquinas.sNomeReduzido
    DescMaquina.Caption = objMaquinas.sDescricao
    If objMaquinas.dTempoMovimentacao <> 0 Then TempoMovimentacao.Text = CStr(objMaquinas.dTempoMovimentacao)
    If objMaquinas.dTempoPreparacao <> 0 Then TempoPreparacao.Text = CStr(objMaquinas.dTempoPreparacao)
    If objMaquinas.dTempoDescarga <> 0 Then TempoDescarga.Text = CStr(objMaquinas.dTempoDescarga)
    
    Call Seleciona_Recurso(objMaquinas.iRecurso)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174506)

    End Select

    Exit Sub

End Sub

Private Sub TempoMovimentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoMovimentacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoMovimentacao, iAlterado)
    
End Sub


Private Sub TempoMovimentacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoMovimentacao_Validate

    'Veifica se TempoMovimentacao está preenchida
    If Len(Trim(TempoMovimentacao.Text)) <> 0 Then

        'Critica a TempoMovimentacao
        lErro = Valor_Positivo_Critica(TempoMovimentacao.Text)
        If lErro <> SUCESSO Then gError 134469
        
        TempoMovimentacao.Text = Formata_Estoque(TempoMovimentacao.Text)

    End If

    Exit Sub

Erro_TempoMovimentacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134469
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174507)

    End Select

    Exit Sub

End Sub

Private Sub TempoOperacao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TempoOperacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoOperacao, iAlterado)
    
End Sub


Private Sub TempoOperacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoOperacao_Validate

    'Verifica se TempoOperacao está preenchido
    If Len(Trim(TempoOperacao.Text)) <> 0 Then
    
        'Critica o Tempo
        lErro = Valor_Positivo_Critica(TempoOperacao.Text)
        If lErro <> SUCESSO Then gError 134470
        
        TempoOperacao.Text = Formata_Estoque(TempoOperacao.Text)
    
    End If
    
    Call Preenche_LabelTaxaDeProducao

    Exit Sub

Erro_TempoOperacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134470
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174508)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
   
On Error GoTo Erro_Tipo_Validate
    
    If Len(Tipo.Text) <> 0 Then
    
       lErro = Habilita_Quantidade()
       If lErro <> SUCESSO Then gError 134471
    
    End If
    
    Call Preenche_LabelTaxaDeProducao

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case gErr

        Case 134471
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174509)

    End Select

    Exit Sub

End Sub

Private Sub UMTempo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UMTempo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUnidadeDeMedida As ClassUnidadeDeMedida

On Error GoTo Erro_UMTempo_Validate

    'Verifica se Taxa está preenchida
    If Len(Trim(UMTempo.Text)) <> 0 Then
    
        Set objUnidadeDeMedida = New ClassUnidadeDeMedida
        
        objUnidadeDeMedida.sSigla = UMTempo.Text
        
        'Verifica se a Taxa digitada está cadastrada
        lErro = CF("UnidadesDeMedidas_Le", objUnidadeDeMedida)
        If lErro <> SUCESSO And lErro <> 134463 Then gError 137550
        
        'Caso não esteja cadastrada -> Erro
        If lErro = 134463 Then gError 137551

        Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_UMTempo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137550
            'erro tratado na rotina chamada
            
        Case 137551
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEDIDA_TEMPO_NAO_CADASTRADA", gErr, objUnidadeDeMedida.sSigla)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174510)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Competencia Then Call LabelCompetencia_Click
    
        If Me.ActiveControl Is Produto Then Call LabelProduto_Click
    
        If Me.ActiveControl Is Maquina Then Call LabelMaquina_Click
    
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoTaxaDeProducao = Nothing
    Set objEventoProduto = Nothing
    Set objEventoMaquina = Nothing
    Set objEventoCompetencia = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174511)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoTaxaDeProducao = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoMaquina = New AdmEvento
    Set objEventoCompetencia = New AdmEvento
        
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 134472
    
    lErro = CarregaComboTipo(Tipo)
    If lErro <> SUCESSO Then gError 134473
    
    lErro = Preenche_Combo_UMTempo()
    If lErro <> SUCESSO Then gError 134474
    
    Call Inicializa_Padrao
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 134472, 134473, 134474
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174512)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTaxaDeProducao Is Nothing) Then

        lErro = Traz_TaxaDeProducao_Tela(objTaxaDeProducao)
        If lErro <> SUCESSO Then gError 134475

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 134475
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174513)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMaquinas As ClassMaquinas
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_Move_Tela_Memoria

    objTaxaDeProducao.dtData = StrParaDate(Data.Caption)

    If Len(Competencia.Text) > 0 Then
    
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = Competencia.Text
        
        'Verifica a Competencia no BD a partir do NomeReduzido
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 134476

        objTaxaDeProducao.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    End If

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134477
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objTaxaDeProducao.sProduto = sProdutoFormatado
    
    End If
    
    If Len(Maquina.Text) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 134478
        
        objTaxaDeProducao.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If
    
    objTaxaDeProducao.iTipo = Codigo_Extrai(Tipo.Text)
    objTaxaDeProducao.dQuantidade = StrParaDbl(Quantidade.Text)
    objTaxaDeProducao.sUMProduto = UMProduto.Text
    objTaxaDeProducao.dTempoOperacao = StrParaDbl(TempoOperacao.Text)
    objTaxaDeProducao.sUMTempo = UMTempo.Text
    objTaxaDeProducao.dTempoPreparacao = StrParaDbl(TempoPreparacao.Text)
    objTaxaDeProducao.dTempoMovimentacao = StrParaDbl(TempoMovimentacao.Text)
    objTaxaDeProducao.dTempoDescarga = StrParaDbl(TempoDescarga.Text)
    objTaxaDeProducao.dLoteMax = StrParaDbl(LoteMaximo.Text)
    objTaxaDeProducao.dLoteMin = StrParaDbl(LoteMinimo.Text)
    objTaxaDeProducao.dLotePadrao = StrParaDbl(LotePadrao.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 134476, 134477, 134478
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174514)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TaxaDeProducao"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134479

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocCompet", objTaxaDeProducao.lNumIntDocCompet, 0, "NumIntDocCompet"
    colCampoValor.Add "Produto", objTaxaDeProducao.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "NumIntDocMaq", objTaxaDeProducao.lNumIntDocMaq, 0, "NumIntDocMaq"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Ativo", OP_IGUAL, TAXA_ATIVA
    
    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174515)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao

On Error GoTo Erro_Tela_Preenche

    objTaxaDeProducao.lNumIntDocCompet = colCampoValor.Item("NumIntDocCompet").vValor
    objTaxaDeProducao.sProduto = colCampoValor.Item("Produto").vValor
    objTaxaDeProducao.lNumIntDocMaq = colCampoValor.Item("NumIntDocMaq").vValor

    lErro = Traz_TaxaDeProducao_Tela(objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134480

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174516)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTaxaDeProducao As New ClassTaxaDeProducao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Data está preenchida
'    If Len(Data.ClipText) = 0 Then gError 137154

    'Verifica se Competência está preenchido
    If Len(Competencia.Text) = 0 Then gError 134481

    'Verifica se Tipo está preenchido
    If Len(Tipo.Text) = 0 Then gError 134482
    
    If Codigo_Extrai(Tipo.Text) <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
    
        'Verifica se Quantidade está preenchido
        If Len(Quantidade.Text) = 0 Then gError 134483
            
        'Verifica se UMProduto está preenchido
        If Len(UMProduto.Text) = 0 Then gError 134484
    
    End If
    
    'Verifica se TempoOperacao está preenchido
    If Len(TempoOperacao.Text) = 0 Then gError 134485
        
    'Verifica se UMTempo está preenchido
    If Len(UMTempo.Text) = 0 Then gError 134486
    
    'Preenche o objTaxaDeProducao
    lErro = Move_Tela_Memoria(objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134487

    lErro = Trata_Alteracao(objTaxaDeProducao, objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> SUCESSO Then gError 137685

    'Grava o/a TaxaDeProducao no Banco de Dados
    lErro = CF("TaxaDeProducao_Grava", objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134488

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134481
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            Competencia.SetFocus

        Case 134482
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TAXADEOPERACAO_NAO_PREENCHIDO", gErr)
            Tipo.SetFocus
        
        Case 134483
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_TAXADEOPERACAO_NAO_PREENCHIDO", gErr)
            Quantidade.SetFocus
        
        Case 134484
            Call Rotina_Erro(vbOKOnly, "ERRO_UMPRODUTO_TAXADEOPERACAO_NAO_PREENCHIDO", gErr)
            UMProduto.SetFocus
        
        Case 134485
            Call Rotina_Erro(vbOKOnly, "ERRO_TEMPOOPERACAO_TAXADEOPERACAO_NAO_PREENCHIDO", gErr)
            TempoOperacao.SetFocus
        
        Case 134486
            Call Rotina_Erro(vbOKOnly, "ERRO_UMTEMPO_TAXADEOPERACAO_NAO_PREENCHIDO", gErr)
            UMTempo.SetFocus

        Case 134487, 134488, 137685
            'erros tratados nas rotinas chamadas
        
'        Case 137154
'            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_TAXADEOPERACAO_NAO_PREENCHIDA", gErr)
'            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174517)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TaxaDeProducao() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TaxaDeProducao
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    DescProd.Caption = ""
    DescMaquina.Caption = ""
    DescCompetencia.Caption = ""
    LabelTaxaDeProducao = ""
    UMProduto.ListIndex = -1
    UMProduto.Clear
    
    Call Inicializa_Padrao

    iAlterado = 0
    
    Limpa_Tela_TaxaDeProducao = SUCESSO

    Exit Function

Erro_Limpa_Tela_TaxaDeProducao:

    Limpa_Tela_TaxaDeProducao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174518)

    End Select

    Exit Function

End Function

Function Traz_TaxaDeProducao_Tela(objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim sProduto As String
Dim sProdutoMascarado As String
Dim objMaquinas As ClassMaquinas
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_Traz_TaxaDeProducao_Tela

    'Limpa a Tela
    Call Limpa_Tela_TaxaDeProducao

    'Lê o TaxaDeProducao que está sendo Passado
    lErro = CF("TaxaDeProducao_Le", objTaxaDeProducao)
    If lErro <> SUCESSO And lErro <> 134543 Then gError 134489

    If lErro = SUCESSO Then

        Data.Caption = Format(objTaxaDeProducao.dtData, "dd/mm/yyyy")
                       
        If objTaxaDeProducao.dLoteMax <> 0 Then LoteMaximo.Text = Formata_Estoque(objTaxaDeProducao.dLoteMax)
        If objTaxaDeProducao.dLoteMin <> 0 Then LoteMinimo.Text = Formata_Estoque(objTaxaDeProducao.dLoteMin)
        If objTaxaDeProducao.dLotePadrao <> 0 Then LotePadrao.Text = Formata_Estoque(objTaxaDeProducao.dLotePadrao)
        If objTaxaDeProducao.dTempoPreparacao <> 0 Then TempoPreparacao.Text = Formata_Estoque(objTaxaDeProducao.dTempoPreparacao)
        If objTaxaDeProducao.dTempoMovimentacao <> 0 Then TempoMovimentacao.Text = Formata_Estoque(objTaxaDeProducao.dTempoMovimentacao)
        If objTaxaDeProducao.dTempoDescarga <> 0 Then TempoDescarga.Text = Formata_Estoque(objTaxaDeProducao.dTempoDescarga)
        If objTaxaDeProducao.dQuantidade <> 0 Then Quantidade.Text = Formata_Estoque(objTaxaDeProducao.dQuantidade)
        If Len(objTaxaDeProducao.sUMProduto) > 0 Then UMProduto.Text = objTaxaDeProducao.sUMProduto
        If objTaxaDeProducao.dTempoOperacao <> 0 Then TempoOperacao.Text = Formata_Estoque(objTaxaDeProducao.dTempoOperacao)
        If Len(objTaxaDeProducao.sUMTempo) > 0 Then UMTempo.Text = objTaxaDeProducao.sUMTempo
        
        Call Combo_Seleciona_ItemData(Tipo, objTaxaDeProducao.iTipo)
                
    End If
    
    If objTaxaDeProducao.lNumIntDocCompet <> 0 Then

        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objTaxaDeProducao.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 134490
        
        Competencia.Text = objCompetencias.sNomeReduzido
        DescCompetencia.Caption = objCompetencias.sDescricao
               
    End If

    If Len(objTaxaDeProducao.sProduto) > 0 Then
    
        sProduto = objTaxaDeProducao.sProduto
        
        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134491

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        Call Produto_Validate(bSGECancelDummy)
        
    End If
    
    If objTaxaDeProducao.lNumIntDocMaq <> 0 Then

        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.lNumIntDoc = objTaxaDeProducao.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 134492
        
        Maquina.Text = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
        Call Seleciona_Recurso(objMaquinas.iRecurso)
        
    End If

    lErro = Habilita_Quantidade()
    If lErro <> SUCESSO Then gError 134493

    iAlterado = 0
    
    Traz_TaxaDeProducao_Tela = SUCESSO

    Exit Function

Erro_Traz_TaxaDeProducao_Tela:

    Traz_TaxaDeProducao_Tela = gErr

    Select Case gErr

        Case 134489 To 134493
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174519)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134494

    'Limpa Tela
    Call Limpa_Tela_TaxaDeProducao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134494

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174520)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174521)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134495

    Call Limpa_Tela_TaxaDeProducao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174522)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMaquinas As ClassMaquinas
Dim objCompetencias As ClassCompetencias
Dim objTaxaDeProducao As New ClassTaxaDeProducao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Competencia.Text) = 0 Then gError 134496
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.sNomeReduzido = Competencia.Text
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 134497

    objTaxaDeProducao.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134498
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objTaxaDeProducao.sProduto = sProdutoFormatado
    
    End If
    
    If Len(Maquina.Text) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Text
        
        'Le a Máquina no BD a partir do Código
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 134499
            
        objTaxaDeProducao.lNumIntDocMaq = objMaquinas.lNumIntDoc
     
    End If

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TAXADEPRODUCAO", objCompetencias.sDescricao)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a Taxa de Producao
    lErro = CF("TaxaDeProducao_Exclui", objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134500

    'Limpa Tela
    Call Limpa_Tela_TaxaDeProducao

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134496
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)
            Competencia.SetFocus
            
        Case 134497 To 134500
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174523)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Maquina_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Maquina_Validate

    DescMaquina.Caption = ""

    'Verifica se Maquina não está preenchida
    If Len(Trim(Maquina.Text)) = 0 Then
    
        TempoDescarga.Enabled = True
        TempoMovimentacao.Enabled = True
        TempoPreparacao.Enabled = True
        LabelTempoDescarga.Enabled = True
        LabelTempoMovimentacao.Enabled = True
        LabelTempoPreparacao.Enabled = True
    
        Exit Sub
    
    End If

    Set objMaquinas = New ClassMaquinas
    
    'Verifica sua existencia
    lErro = CF("TP_Maquina_Le", Maquina, objMaquinas)
    If lErro <> SUCESSO Then gError 134501
    
    DescMaquina.Caption = objMaquinas.sDescricao
    If objMaquinas.dTempoMovimentacao <> 0 Then TempoMovimentacao.Text = CStr(objMaquinas.dTempoMovimentacao)
    If objMaquinas.dTempoPreparacao <> 0 Then TempoPreparacao.Text = CStr(objMaquinas.dTempoPreparacao)
    If objMaquinas.dTempoDescarga <> 0 Then TempoDescarga.Text = CStr(objMaquinas.dTempoDescarga)
    
    Call Seleciona_Recurso(objMaquinas.iRecurso)
    
    Exit Sub

Erro_Maquina_Validate:

    Cancel = True

    Select Case gErr

        Case 134501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174524)

    End Select

    Exit Sub

End Sub

Private Sub Maquina_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Maquina, iAlterado)
    
End Sub

Private Sub Maquina_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

        'Critica a Quantidade
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 134502
        
        Quantidade.Text = Formata_Estoque(Quantidade.Text)

        Call Preenche_LabelTaxaDeProducao

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 134502

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174525)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UMProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_UMProduto_Validate

    'Verifica se Taxa está preenchida
    If Len(Trim(UMProduto.Text)) <> 0 Then
    
        Set objUnidadeDeMedida = New ClassUnidadeDeMedida
        
        objUnidadeDeMedida.sSigla = UMProduto.Text
        
        'Verifica se a Taxa digitada está cadastrada
        lErro = CF("UnidadesDeMedidas_Le", objUnidadeDeMedida)
        If lErro <> SUCESSO And lErro <> 134463 Then gError 134503
        
        'Caso não esteja cadastrada -> Erro
        If lErro = 134463 Then gError 134504
        
        Call Preenche_LabelTaxaDeProducao
        
        sCodProduto = Produto.Text
        
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 134505
        
        If Len(sProdutoFormatado) = 0 Then
        
            FrameLotes.Caption = "Lotes (" & UMProduto.Text & ")"
    
        End If
        
    End If

    Exit Sub

Erro_UMProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 134503, 134505
            'erro tratado na rotina chamada
            
        Case 134504
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_CADASTRADA1", gErr, objUnidadeDeMedida.sSigla)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174526)

    End Select

    Exit Sub

End Sub

Private Sub UMProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoTaxaDeProducao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTaxaDeProducao As ClassTaxaDeProducao

On Error GoTo Erro_objEventoTaxaDeProducao_evSelecao

    Set objTaxaDeProducao = obj1

    'Mostra os dados do TaxaDeProducao na tela
    lErro = Traz_TaxaDeProducao_Tela(objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134506

    Me.Show

    Exit Sub

Erro_objEventoTaxaDeProducao_evSelecao:

    Select Case gErr

        Case 134506

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174527)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_LabelProduto_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134507

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    sFiltro = "Ativo = ? And Compras <> ?"
    
    colSelecao.Add PRODUTO_ATIVO
    colSelecao.Add PRODUTO_COMPRAVEL
        
    'Lista de produtos
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, sFiltro)
    
    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 134507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174528)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String


On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134508

    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True
    
    Call Produto_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 134508
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174529)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String
Dim sUMProduto As String

On Error GoTo Erro_Produto_Validate

    If Len(Trim(Produto.ClipText)) > 0 Then
    
        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 134509
                        
        'se o produto não estiver cadastrado ==> Erro
        If lErro = 51381 Then gError 134511
        
        'se o produto não for ativo ==> Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 134550
    
        'se o produto for comprável ==> Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 134551
        
        'então podemos continuar ...
        If Codigo_Extrai(Tipo.Text) <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
            
            lErro = CarregaComboUM(objProduto)
            If lErro <> SUCESSO Then gError 134512
                    
        End If
        
        FrameLotes.Caption = "Lotes (" & objProduto.sSiglaUMEstoque & ")"
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134513
    
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        DescProd.Caption = objProduto.sDescricao
        
    Else
        
        sUMProduto = UMProduto.Text
        UMProduto.Clear
        
        If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
        
            FrameLotes.Caption = "Lotes"
        
        Else
        
            UMProduto.Text = sUMProduto
            FrameLotes.Caption = "Lotes (" & UMProduto.Text & ")"
        
        End If

        DescProd.Caption = ""
        
    End If
    
    Call Preenche_LabelTaxaDeProducao
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 134509, 134510, 134512, 134513

        Case 134511
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 134550
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)
        
        Case 134551
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174530)

    End Select

    Exit Sub

End Sub

Private Sub TempoPreparacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoPreparacao_Validate

    'Verifica se TempoPreparacao está preenchida
    If Len(Trim(TempoPreparacao.Text)) <> 0 Then

        'Critica a TempoPreparacao
        lErro = Valor_Positivo_Critica(TempoPreparacao.Text)
        If lErro <> SUCESSO Then gError 134514
        
        TempoPreparacao.Text = Formata_Estoque(TempoPreparacao.Text)

    End If

    Exit Sub

Erro_TempoPreparacao_Validate:

    Cancel = True

    Select Case gErr

        Case 134514
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174531)

    End Select

    Exit Sub

End Sub

Private Sub TempoPreparacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoPreparacao, iAlterado)
    
End Sub

Private Sub TempoPreparacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoDescarga_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TempoDescarga_Validate

    'Verifica se TempoDescarga está preenchida
    If Len(Trim(TempoDescarga.Text)) <> 0 Then

       'Critica a TempoDescarga
       lErro = Valor_Positivo_Critica(TempoDescarga.Text)
       If lErro <> SUCESSO Then gError 134515
       
       TempoDescarga.Text = Formata_Estoque(TempoDescarga.Text)

    End If

    Exit Sub

Erro_TempoDescarga_Validate:

    Cancel = True

    Select Case gErr

        Case 134515
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174532)

    End Select

    Exit Sub

End Sub

Private Sub TempoDescarga_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TempoDescarga, iAlterado)
    
End Sub

Private Sub TempoDescarga_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMinimo_Validate

    'Verifica se LoteMinimo está preenchido
    If Len(Trim(LoteMinimo.Text)) <> 0 Then

       'Critica a LoteMinimo
       lErro = Valor_Positivo_Critica(LoteMinimo.Text)
       If lErro <> SUCESSO Then gError 134516
       
       LoteMinimo.Text = Formata_Estoque(LoteMinimo.Text)

    End If

    Exit Sub

Erro_LoteMinimo_Validate:

    Cancel = True

    Select Case gErr

        Case 134516

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174533)

    End Select

    Exit Sub

End Sub

Private Sub LoteMinimo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteMinimo, iAlterado)
    
End Sub

Private Sub LoteMinimo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMaximo_Validate

    'Verifica se LoteMaximo está preenchido
    If Len(Trim(LoteMaximo.Text)) <> 0 Then

       'Critica a LoteMaximo
       lErro = Valor_Positivo_Critica(LoteMaximo.Text)
       If lErro <> SUCESSO Then gError 134517
       
       LoteMaximo.Text = Formata_Estoque(LoteMaximo.Text)

    End If

    Exit Sub

Erro_LoteMaximo_Validate:

    Cancel = True

    Select Case gErr

        Case 134517

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174534)

    End Select

    Exit Sub

End Sub

Private Sub LoteMaximo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(LoteMaximo, iAlterado)
    
End Sub

Private Sub LoteMaximo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function CarregaComboTipo(objCombo As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CarregaComboTipo

    
    objCombo.AddItem ITEM_TIPO_TAXAPRODUCAO_VARIAVEL & SEPARADOR & STRING_ITEM_TIPO_TAXAPRODUCAO_VARIAVEL
    objCombo.ItemData(objCombo.NewIndex) = ITEM_TIPO_TAXAPRODUCAO_VARIAVEL
    
    objCombo.AddItem ITEM_TIPO_TAXAPRODUCAO_FIXO & SEPARADOR & STRING_ITEM_TIPO_TAXAPRODUCAO_FIXO
    objCombo.ItemData(objCombo.NewIndex) = ITEM_TIPO_TAXAPRODUCAO_FIXO
    
    CarregaComboTipo = SUCESSO

    Exit Function

Erro_CarregaComboTipo:

    CarregaComboTipo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174535)

    End Select

    Exit Function

End Function

Private Function Preenche_LabelTaxaDeProducao() As Long

Dim lErro As Long
Dim dQtdeUnitaria As Double
Dim sTaxaDeProducao As String

On Error GoTo Erro_Preenche_LabelTaxaDeProducao

    sTaxaDeProducao = ""
    
    If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
    
        If Len(TempoOperacao.Text) <> 0 And Len(UMTempo.Text) <> 0 Then
            sTaxaDeProducao = Formata_Estoque(StrParaDbl(TempoOperacao.Text)) & " " & UMTempo.Text
        End If
    
    Else
    
        If Len(Quantidade.Text) <> 0 And Len(UMProduto.Text) <> 0 Then
            
            If Len(TempoOperacao.Text) <> 0 And Len(UMTempo.Text) <> 0 Then
                
                dQtdeUnitaria = CDbl(Quantidade.Text) / CDbl(TempoOperacao.Text)
                sTaxaDeProducao = Formata_Estoque(dQtdeUnitaria) & " " & UMProduto.Text
                sTaxaDeProducao = sTaxaDeProducao & "/" & UMTempo.Text
            
            End If
    
        End If
    
    End If
            
    LabelTaxaDeProducao.Caption = sTaxaDeProducao
    
    Exit Function

Erro_Preenche_LabelTaxaDeProducao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174536)

    End Select

    Exit Function

End Function

Private Function Preenche_Combo_UMTempo()

Dim objClasseUM As ClassClasseUM
Dim colSiglas As New Collection
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim lErro As Long

On Error GoTo Erro_Preenche_Combo_UMTempo

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = gobjEST.iClasseUMTempo

    'Preenche a List da Combo UnidadeMed com as UM's de Tempo
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then gError 134518

    'Se tem algum valor para UMTempo
    If Len(UMTempo.Text) > 0 Then
        'Guardo o valor da UMTempo da Linha
        sUnidadeMed = UMTempo.Text
    Else
        'Senão coloco o Padrão UMTempo
        Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)
        'sUnidadeMed = TAXA_CONSUMO_TEMPO_PADRAO
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UMTempo.Clear

    For Each objUnidadeDeMedida In colSiglas
        UMTempo.AddItem objUnidadeDeMedida.sSigla
    Next

    UMTempo.AddItem ""

    'Tento selecionar na Combo a Unidade anterior
    If UMTempo.ListCount <> 0 Then

        For iIndice = 0 To UMTempo.ListCount - 1

            If UMTempo.List(iIndice) = sUnidadeMed Then
                UMTempo.ListIndex = iIndice
                Exit For
            End If
        Next
    End If

    Exit Function

Erro_Preenche_Combo_UMTempo:

    Select Case gErr
    
        Case 134518

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174537)

    End Select

    Exit Function

End Function

Private Function Inicializa_Padrao() As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Padrao

    Data.Caption = Format(gdtDataAtual, "dd/mm/yyyy")
    
    Tipo.ListIndex = 0

    LabelQuantidade.Enabled = True
    Quantidade.Enabled = True
    
    LabelUMProduto.Enabled = True
    UMProduto.Clear
    
    TempoOperacao.Text = Formata_Estoque(1)
    
    'sUnidadeMed = TAXA_CONSUMO_TEMPO_PADRAO
    Call CF("Taxa_Producao_UM_Padrao_Obtem", sUnidadeMed)

    'Tento selecionar na Combo a Unidade Padrão
    If UMTempo.ListCount <> 0 Then

        For iIndice = 0 To UMTempo.ListCount - 1

            If UMTempo.List(iIndice) = sUnidadeMed Then
                UMTempo.ListIndex = iIndice
                Exit For
            End If
        
        Next
    
    End If
    
    FrameLotes.Caption = "Lotes"
    
    TempoDescarga.Enabled = True
    TempoMovimentacao.Enabled = True
    TempoPreparacao.Enabled = True
    LabelTempoDescarga.Enabled = True
    LabelTempoMovimentacao.Enabled = True
    LabelTempoPreparacao.Enabled = True

    Call Preenche_LabelTaxaDeProducao

    Exit Function
        
Erro_Inicializa_Padrao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174538)

    End Select

    Exit Function
    
End Function

Private Function CarregaComboUM(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_CarregaComboUM

    Set objClasseUM = New ClassClasseUM
    
    objClasseUM.iClasse = objProduto.iClasseUM
    
    'Preenche a List da Combo UnidadeMed com as UM's da Competencia
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then gError 134519

    'Se tem algum valor para UM na Tela
    If Len(UMProduto.Text) > 0 Then
        'Guardo o valor da UMProduto da Tela
        sUnidadeMed = UMProduto.Text
    Else
        'Senão coloco a do Estoque do Produto
        sUnidadeMed = objProduto.sSiglaUMEstoque
    End If
    
    'Limpar as Unidades utilizadas anteriormente
    UMProduto.Clear

    For Each objUnidadeDeMedida In colSiglas
        UMProduto.AddItem objUnidadeDeMedida.sSigla
    Next

    UMProduto.AddItem ""

    'Tento selecionar na Combo a Unidade anterior
    If UMProduto.ListCount <> 0 Then

        For iIndice = 0 To UMProduto.ListCount - 1

            If UMProduto.List(iIndice) = sUnidadeMed Then
                UMProduto.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    CarregaComboUM = SUCESSO
    
    Exit Function

Erro_CarregaComboUM:

    CarregaComboUM = gErr

    Select Case gErr

        Case 134519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174539)

    End Select

    Exit Function

End Function

Private Function Habilita_Quantidade() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String
Dim sUMProduto As String
   
On Error GoTo Erro_Habilita_Quantidade
        
    If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
    
        Quantidade.Text = ""
        LabelQuantidade.Enabled = False
        Quantidade.Enabled = False
        
        UMProduto.Clear
        LabelUMProduto.Enabled = False
        UMProduto.Enabled = False
                
    Else
        
        Quantidade.Enabled = True
        LabelQuantidade.Enabled = True
        
        UMProduto.Enabled = True
        LabelUMProduto.Enabled = True
        
    End If
        
    sCodProduto = Produto.Text
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sCodProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 134521
        
        If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
        
            Quantidade.Text = ""
            LabelQuantidade.Enabled = False
            Quantidade.Enabled = False
            
            UMProduto.Clear
            LabelUMProduto.Enabled = False
            UMProduto.Enabled = False
        
        Else
            
            Quantidade.Enabled = True
            LabelQuantidade.Enabled = True
            
            UMProduto.Enabled = True
            LabelUMProduto.Enabled = True
                    
            lErro = CarregaComboUM(objProduto)
            If lErro <> SUCESSO Then gError 134200
        
        End If
        
        FrameLotes.Caption = "Lotes (" & objProduto.sSiglaUMEstoque & ")"
          
    Else
    
        If Codigo_Extrai(Tipo.Text) = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
        
            Quantidade.Text = ""
            LabelQuantidade.Enabled = False
            Quantidade.Enabled = False
            
            UMProduto.Clear
            LabelUMProduto.Enabled = False
            UMProduto.Enabled = False
        
            FrameLotes.Caption = "Lotes"
        
        Else
            
            Quantidade.Enabled = True
            LabelQuantidade.Enabled = True
            
            UMProduto.Enabled = True
            LabelUMProduto.Enabled = True
    
            sUMProduto = UMProduto.Text
            UMProduto.Clear
            
            UMProduto.Text = sUMProduto
            If Len(Trim(UMProduto.Text)) <> 0 Then
            
                FrameLotes.Caption = "Lotes (" & UMProduto.Text & ")"
                
            Else
            
                FrameLotes.Caption = "Lotes"

            End If
        
        End If

    End If
           
    Call Preenche_LabelTaxaDeProducao
    
    Habilita_Quantidade = SUCESSO

    Exit Function

Erro_Habilita_Quantidade:

    Habilita_Quantidade = gErr

    Select Case gErr

        Case 134519, 134520

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174540)

    End Select

    Exit Function

End Function

Private Sub Seleciona_Recurso(ByVal iRecurso As Integer)
        
    Select Case iRecurso
    
        Case ITEMCT_RECURSO_MAQUINA
        
            TempoDescarga.Text = ""
            TempoDescarga.Enabled = False
            TempoMovimentacao.Enabled = True
            TempoPreparacao.Enabled = True
            LabelTempoDescarga.Enabled = False
            LabelTempoMovimentacao.Enabled = True
            LabelTempoPreparacao.Enabled = True
        
        Case ITEMCT_RECURSO_HABILIDADE
        
            TempoDescarga.Text = ""
            TempoPreparacao.Text = ""
            TempoDescarga.Enabled = False
            TempoMovimentacao.Enabled = True
            TempoPreparacao.Enabled = False
            LabelTempoDescarga.Enabled = False
            LabelTempoMovimentacao.Enabled = True
            LabelTempoPreparacao.Enabled = False
        
        Case ITEMCT_RECURSO_PROCESSO
        
            TempoMovimentacao.Text = ""
            TempoDescarga.Enabled = True
            TempoMovimentacao.Enabled = False
            TempoPreparacao.Enabled = True
            LabelTempoDescarga.Enabled = True
            LabelTempoMovimentacao.Enabled = False
            LabelTempoPreparacao.Enabled = True
            
    End Select

End Sub

