VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ParcelasRecDifOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   7650
   Begin VB.Frame Frame3 
      Caption         =   "Diferença"
      Height          =   2100
      Left            =   165
      TabIndex        =   31
      Top             =   3765
      Width           =   7260
      Begin VB.TextBox Observacao 
         Height          =   315
         Left            =   1710
         MaxLength       =   250
         TabIndex        =   8
         Top             =   1215
         Width           =   5415
      End
      Begin MSMask.MaskEdBox CodTipoDif 
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   300
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorDiferenca 
         Height          =   315
         Left            =   1725
         TabIndex        =   7
         Top             =   750
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   11
         Format          =   "##,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1725
         TabIndex        =   37
         Top             =   1665
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Registrada em :"
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
         Left            =   255
         TabIndex        =   36
         Top             =   1695
         Width           =   1440
      End
      Begin VB.Label DescricaoTipoDif 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2460
         TabIndex        =   35
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label LabelObservacao 
         Alignment       =   1  'Right Justify
         Caption         =   "Observação:"
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
         Left            =   300
         TabIndex        =   34
         Top             =   1245
         Width           =   1395
      End
      Begin VB.Label LabelValorDiferenca 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   300
         TabIndex        =   33
         Top             =   795
         Width           =   1395
      End
      Begin VB.Label LabelCodTipoDif 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo da Diferença:"
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
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   330
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situação Atual"
      Height          =   1080
      Left            =   180
      TabIndex        =   22
      Top             =   2550
      Width           =   7275
      Begin VB.Label ValorTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   30
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Título:"
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
         Left            =   3480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label SaldoTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo do Título:"
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
         Left            =   3435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label ValorParc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   26
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Parcela:"
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
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   735
         Width           =   1485
      End
      Begin VB.Label SaldoParc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   24
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo da Parcela:"
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   285
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   1785
      Left            =   180
      TabIndex        =   14
      Top             =   675
      Width           =   7275
      Begin VB.CommandButton BotaoLimparSeq 
         Height          =   300
         Left            =   5040
         Picture         =   "ParcelasRecDif.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar o Número"
         Top             =   1305
         Width           =   345
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4575
         TabIndex        =   1
         Top             =   345
         Width           =   2310
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "ParcelasRecDif.ctx":0532
         Left            =   960
         List            =   "ParcelasRecDif.ctx":0534
         TabIndex        =   2
         Top             =   840
         Width           =   2190
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   345
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   4050
         TabIndex        =   3
         Top             =   840
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   6810
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   810
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   5715
         TabIndex        =   41
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label16 
         Caption         =   "Emissão:"
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
         Left            =   4905
         TabIndex        =   42
         Top             =   855
         Width           =   750
      End
      Begin VB.Label Label5 
         Caption         =   "Vcto.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1455
         TabIndex        =   39
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Vencimento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2025
         TabIndex        =   38
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Sequencial 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4365
         TabIndex        =   21
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label ClienteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   405
         Width           =   660
      End
      Begin VB.Label NumeroLabel 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   885
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
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
         Left            =   3975
         TabIndex        =   18
         Top             =   405
         Width           =   525
      End
      Begin VB.Label LabelParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label LabelTipo 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   900
         Width           =   450
      End
      Begin VB.Label LabelSequencial 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial:"
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
         Left            =   3270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   1380
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5355
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "ParcelasRecDif.ctx":0536
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "ParcelasRecDif.ctx":0690
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "ParcelasRecDif.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "ParcelasRecDif.ctx":0D4C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "ParcelasRecDifOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFilialAlterada As Integer
Dim iTipoAlterado As Integer
Dim iNumeroAlterado As Integer
Dim iParcelaAlterada As Integer
Dim iSequencialAlterado As Integer

Dim lClienteAnterior As Long
Dim iFilialAnterior As Integer
Dim sTipoAnterior As String
Dim lNumeroAnterior As Long
Dim iParcelaAnterior As Integer

Dim glNumIntParc As Long

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1
Private WithEvents objEventoSequencial As AdmEvento
Attribute objEventoSequencial.VB_VarHelpID = -1
Private WithEvents objEventoTipoDoc As AdmEvento
Attribute objEventoTipoDoc.VB_VarHelpID = -1
Private WithEvents objEventoCodTipoDif As AdmEvento
Attribute objEventoCodTipoDif.VB_VarHelpID = -1

Private gobjTituloReceber As New ClassTituloReceber
Private gobjParcelaReceber As New ClassParcelaReceber

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Diferenças nas Parcelas a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ParcelasRecDif"

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

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoParcela = Nothing
    Set objEventoSequencial = Nothing
    Set objEventoTipoDoc = Nothing
    Set objEventoCodTipoDif = Nothing
    
    Set gobjTituloReceber = Nothing
    Set gobjParcelaReceber = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177818)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega os Tipos de Documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then gError 28520

    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoParcela = New AdmEvento
    Set objEventoSequencial = New AdmEvento
    Set objEventoTipoDoc = New AdmEvento
    Set objEventoCodTipoDif = New AdmEvento
    
    Data.Caption = Format(gdtDataAtual, "dd/mm/yy")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177819)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objParcelasRecDif As ClassParcelasRecDif) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objParcelasRecDif Is Nothing) Then

        'se a diferenca existe e está identificada
        If objParcelasRecDif.lNumIntParc <> 0 And objParcelasRecDif.iSeq <> 0 Then
        
            lErro = Traz_ParcelasRecDif_Tela(objParcelasRecDif)
            If lErro <> SUCESSO Then gError 177852

        Else
        
            'se apenas a parcela está identificada
            If objParcelasRecDif.lNumIntParc <> 0 Then
            
                lErro = Traz_ParcelasRecDif_Tela2(objParcelasRecDif)
                If lErro <> SUCESSO Then gError 177852
            
            End If
            
        End If
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 177852

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177820)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objParcelasRecDif As ClassParcelasRecDif) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objParcelasRecDif.lNumIntParc = glNumIntParc
    objParcelasRecDif.iSeq = StrParaInt(Sequencial.Caption)
    objParcelasRecDif.dtDataRegistro = gdtDataAtual
    objParcelasRecDif.iCodTipoDif = StrParaInt(CodTipoDif.Text)
    objParcelasRecDif.dValorDiferenca = StrParaDbl(ValorDiferenca.Text)
    objParcelasRecDif.sObservacao = Observacao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177821)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objParcelasRecDif As New ClassParcelasRecDif

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ParcelasRecDif"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objParcelasRecDif)
    If lErro <> SUCESSO Then gError 177853

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntParc", objParcelasRecDif.lNumIntParc, 0, "NumIntParc"
    colCampoValor.Add "Seq", objParcelasRecDif.iSeq, 0, "Seq"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 177853

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177822)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objParcelasRecDif As New ClassParcelasRecDif

On Error GoTo Erro_Tela_Preenche

    objParcelasRecDif.lNumIntParc = colCampoValor.Item("NumIntParc").vValor
    objParcelasRecDif.iSeq = colCampoValor.Item("Seq").vValor

    If objParcelasRecDif.lNumIntParc <> 0 And objParcelasRecDif.iSeq <> 0 Then
    
        lErro = Traz_ParcelasRecDif_Tela(objParcelasRecDif)
        If lErro <> SUCESSO Then gError 177854
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 177854

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177823)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objParcelasRecDif As New ClassParcelasRecDif

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If glNumIntParc = 0 Then gError 177855

    'Preenche o objParcelasRecDif
    lErro = Move_Tela_Memoria(objParcelasRecDif)
    If lErro <> SUCESSO Then gError 177856

    lErro = Trata_Alteracao(objParcelasRecDif, objParcelasRecDif.lNumIntParc, objParcelasRecDif.iSeq)
    If lErro <> SUCESSO Then gError 177857

    'Grava o/a ParcelasRecDif no Banco de Dados
    lErro = CF("ParcelasRecDif_Grava", objParcelasRecDif)
    If lErro <> SUCESSO Then gError 177858

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177855
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", gErr)
            Parcela.SetFocus

        Case 177856, 177857, 177858

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177824)

    End Select

    Exit Function

End Function

Function Limpa_Tela_ParcelasRecDif() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ParcelasRecDif

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Data.Caption = Format(gdtDataAtual, "dd/mm/yy")
    
    DescricaoTipoDif.Caption = ""
    Vencimento.Caption = ""
    SaldoParc.Caption = ""
    ValorParc.Caption = ""
    SaldoTitulo.Caption = ""
    ValorTitulo.Caption = ""
    Sequencial.Caption = ""
    
     iAlterado = 0
     iClienteAlterado = 0
     iFilialAlterada = 0
     iTipoAlterado = 0
     iNumeroAlterado = 0
     iParcelaAlterada = 0
     iSequencialAlterado = 0
    
     lClienteAnterior = 0
     iFilialAnterior = 0
     sTipoAnterior = ""
     lNumeroAnterior = 0
     iParcelaAnterior = 0
    
    Filial.Clear
    Tipo.ListIndex = -1

    iAlterado = 0

    Limpa_Tela_ParcelasRecDif = SUCESSO

    Exit Function

Erro_Limpa_Tela_ParcelasRecDif:

    Limpa_Tela_ParcelasRecDif = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177825)

    End Select

    Exit Function

End Function

Function Traz_ParcelasRecDif_Tela(objParcelasRecDif As ClassParcelasRecDif) As Long

Dim lErro As Long
Dim objParcRec As New ClassParcelaReceber
Dim objTitRec As New ClassTituloReceber

On Error GoTo Erro_Traz_ParcelasRecDif_Tela

    'Lê o ParcelasRecDif que está sendo Passado
    lErro = CF("ParcelasRecDif_Le", objParcelasRecDif)
    If lErro <> SUCESSO And lErro <> 177862 Then gError 177863

    If lErro = SUCESSO Then

        objParcRec.lNumIntDoc = objParcelasRecDif.lNumIntParc
        
        lErro = CF("ParcelaReceber_Le", objParcRec)
        If lErro <> SUCESSO And lErro <> 19147 Then gError 177908
        
        If lErro <> SUCESSO Then

            lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
            If lErro <> SUCESSO And lErro <> 58559 Then gError 177908

        End If
        
        objTitRec.lNumIntDoc = objParcRec.lNumIntTitulo
        
        lErro = Traz_TitReceber_Tela(objTitRec)
        If lErro <> SUCESSO Then gError 177909
        
        glNumIntParc = objParcelasRecDif.lNumIntParc

        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcRec.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bSGECancelDummy)
        
        Vencimento.Caption = Format(objParcRec.dtDataVencimento, "dd/mm/yyyy")
        SaldoParc.Caption = Format(objParcRec.dSaldo, "STANDARD")
        ValorParc.Caption = Format(objParcRec.dValor, "STANDARD")

        Sequencial.Caption = CStr(objParcelasRecDif.iSeq)
        Data.Caption = Format(objParcelasRecDif.dtDataRegistro, "dd/mm/yy")

        If objParcelasRecDif.iCodTipoDif <> 0 Then
            
            CodTipoDif.Text = CStr(objParcelasRecDif.iCodTipoDif)
            
            Call CodTipoDif_Validate(bSGECancelDummy)
            
        Else
        
            DescricaoTipoDif.Caption = ""
            
        End If
        
        If objParcelasRecDif.dValorDiferenca <> 0 Then
            ValorDiferenca.Text = Format(objParcelasRecDif.dValorDiferenca, "STANDARD")
        End If
        
        Observacao.Text = objParcelasRecDif.sObservacao

    End If

    iAlterado = 0

    Traz_ParcelasRecDif_Tela = SUCESSO

    Exit Function

Erro_Traz_ParcelasRecDif_Tela:

    Traz_ParcelasRecDif_Tela = gErr

    Select Case gErr

        Case 177863, 177908, 177909
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177826)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 177866

    'Limpa Tela
    Call Limpa_Tela_ParcelasRecDif

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 177866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177827)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177828)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 177866

    Call Limpa_Tela_ParcelasRecDif

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 177866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177829)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objParcelasRecDif As New ClassParcelasRecDif
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If glNumIntParc = 0 Then gError 177867
    If Len(Trim(Sequencial.Caption)) = 0 Then gError 177868

    objParcelasRecDif.lNumIntParc = glNumIntParc
    objParcelasRecDif.iSeq = StrParaInt(Sequencial.Caption)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PARCELASRECDIF", objParcelasRecDif.iSeq)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("ParcelasRecDif_Exclui", objParcelasRecDif)
        If lErro <> SUCESSO Then gError 177869

        'Limpa Tela
        Call Limpa_Tela_ParcelasRecDif

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177867
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", gErr)
            Parcela.SetFocus

        Case 177868
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case 177869

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177830)

    End Select

    Exit Sub

End Sub

Private Sub CodTipoDif_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDif As New ClassTiposDifParcRec

On Error GoTo Erro_CodTipoDif_Validate

    'Verifica se CodTipoDif está preenchida
    If Len(Trim(CodTipoDif.Text)) <> 0 Then

       'Critica a CodTipoDif
       lErro = Inteiro_Critica(CodTipoDif.Text)
       If lErro <> SUCESSO Then gError 177870
       
        objTipoDif.iCodigo = StrParaInt(CodTipoDif.Text)
        
        lErro = CF("TiposDifParcRec_Le", objTipoDif)
        If lErro <> SUCESSO And lErro <> 177657 Then gError 177871
        
        If lErro <> SUCESSO Then gError 177872
        
        DescricaoTipoDif.Caption = objTipoDif.sDescricao
    
    Else
       
        DescricaoTipoDif.Caption = ""

    End If

    Exit Sub

Erro_CodTipoDif_Validate:

    Cancel = True

    Select Case gErr

        Case 177870, 177871
        
        Case 177872
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDIFPARCREC_NAO_CADASTRADO", gErr, objTipoDif.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177831)

    End Select

    Exit Sub

End Sub

Private Sub CodTipoDif_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodTipoDif, iAlterado)
    
End Sub

Private Sub CodTipoDif_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorDiferenca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDiferenca_Validate

    'Verifica se ValorDiferenca está preenchida
    If Len(Trim(ValorDiferenca.Text)) <> 0 Then

       'Critica a ValorDiferenca
       lErro = Valor_Critica(ValorDiferenca.Text)
       If lErro <> SUCESSO Then gError 177873

    End If

    Exit Sub

Erro_ValorDiferenca_Validate:

    Cancel = True

    Select Case gErr

        Case 177873

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177832)

    End Select

    Exit Sub

End Sub

Private Sub ValorDiferenca_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValorDiferenca, iAlterado)
    
End Sub

Private Sub ValorDiferenca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177833)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Change()

   iClienteAlterado = 1
   iAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente, lCliente As Long
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 177873

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 177874

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
        lCliente = objCliente.lCodigo

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear
        
        lCliente = 0

    End If
    
    If lClienteAnterior <> lCliente Then
    
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
        SaldoTitulo.Caption = ""
        ValorTitulo.Caption = ""
        Vencimento.Caption = ""
        SaldoParc.Caption = ""
        ValorParc.Caption = ""
        
        lNumeroAnterior = 0
        iParcelaAnterior = 0
        
        'Se Cliente foi alterado zera glNumIntParc
        glNumIntParc = 0

        iClienteAlterado = 0
        
        Call Verifica_Alteracao

    End If
    
    lClienteAnterior = lCliente

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 177873
            
        Case 177874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177834)

    End Select

    Exit Sub

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    If IsNumeric(Cliente.Text) Then
        objCliente.lCodigo = StrParaLong(Cliente.Text)
    End If
    
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Filial_Change()

    iFilialAlterada = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    If Filial.ListIndex = -1 Then Exit Sub

    Call Filial_Validate(bSGECancelDummy)

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177835)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    If iFilialAnterior <> Codigo_Extrai(Filial.Text) Then
    
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
        SaldoTitulo.Caption = ""
        ValorTitulo.Caption = ""
        Vencimento.Caption = ""
        SaldoParc.Caption = ""
        ValorParc.Caption = ""
        
    End If

    If iFilialAlterada Then

        'Verifica se a filial foi preenchida
        If Len(Trim(Filial.Text)) = 0 Then Exit Sub

        'Verifica se é uma filial selecionada
        If Filial.Text = Filial.List(Filial.ListIndex) Then
            Call Verifica_Alteracao
            Exit Sub
        End If
        
        'Tenta selecionar na combo
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 177875

        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then

            'Verifica se o Cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then gError 177876

            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo

            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 177877

            If lErro = 17660 Then gError 177878

            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        End If

        'Não encontrou a STRING
        If lErro = 6731 Then gError 177879
        
        'Se Filial foi alterado zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iFilialAlterada = 0

    End If

    iFilialAnterior = Codigo_Extrai(Filial.Text)

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 177875, 177877

        Case 177876
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 177878
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 177879
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177836)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Lê os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 177880

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 177880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177837)

    End Select

    Exit Function

End Function

Private Sub LabelParcela_Click()
'Lista as parcelas do titulo selecionado

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objParcelaReceber As ClassParcelaReceber
Dim colSelecao As New Collection

On Error GoTo Erro_LabelParcela_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 177881
    If Len(Trim(Filial.Text)) = 0 Then gError 177882
    If Len(Trim(Tipo.Text)) = 0 Then gError 177883
    If Len(Trim(Numero.ClipText)) = 0 Then gError 177884
    
    objCliente.sNomeReduzido = Cliente.Text
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 177885
    
    'Se não achou o Cliente --> erro
    If lErro <> SUCESSO Then gError 177886
    
    colSelecao.Add objCliente.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)
    colSelecao.Add SCodigo_Extrai(Tipo.Text)
    colSelecao.Add StrParaLong(Numero.Text)
    
    'Chama a tela
    Call Chama_Tela("ParcelasRecLista", colSelecao, objParcelaReceber, objEventoParcela)
    
    Exit Sub
    
Erro_LabelParcela_Click:

    Select Case gErr
    
        Case 177881
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
    
        Case 177882
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 177883
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            
        Case 177884
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case 177885
    
        Case 177886
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177838)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelSequencial_Click()

Dim lErro As Long
Dim objParcelasRecDif As New ClassParcelasRecDif
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSequencial_Click

    If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Or _
    Len(Trim(Numero.Text)) = 0 Or Len(Trim(Parcela.Text)) = 0 Then gError 177887
    
    If glNumIntParc = 0 Then gError 177888

    objParcelasRecDif.lNumIntParc = glNumIntParc
    
    colSelecao.Add objParcelasRecDif.lNumIntParc
        
    'Chama a tela
    Call Chama_Tela("ParcelasRecDifLista", colSelecao, objParcelasRecDif, objEventoSequencial)
    
    Exit Sub
    
Erro_LabelSequencial_Click:

    Select Case gErr
    
        Case 177887
            Call Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", gErr)
    
        Case 177888
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177839)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelTipo_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)
    
    'Chama a tela TipoDocTituloRecLista
    Call Chama_Tela("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDoc)

End Sub

Private Sub Numero_Change()

    iNumeroAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

Dim iNumAux As Integer
    
    iNumAux = iNumeroAlterado
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
    iNumeroAlterado = iNumAux
    
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Número está preenchido
    If Len(Trim(Numero.ClipText)) = 0 Then Exit Sub
    
    If lNumeroAnterior <> StrParaLong(Numero.Text) Then
    
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
    
    End If

    'Critica se é Long positivo
    lErro = Long_Critica(Numero.ClipText)
    If lErro <> SUCESSO Then gError 177889

    If iNumeroAlterado Then
        
        'Se Número foi alterado zera glNUmIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iNumeroAlterado = 0

    End If
    
    lNumeroAnterior = StrParaLong(Numero.Text)

    Exit Sub

Erro_Numero_Validate:

    Cancel = True


    Select Case gErr

        Case 177889

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177840)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloReceber As New ClassTituloReceber
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_NumeroLabel_Click

    If Len(Trim(Cliente.Text)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 177889
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 177890

    End If
    
    'Guarda o código no objTituloReceber
    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)

    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If

    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If

    If Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND SiglaDocumento = ?"
        Else
            iPreenchido = 1
            sSelecao = "SiglaDocumento = ?"
        End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)
    End If

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 177889

        Case 177890
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177841)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1
    
    lErro = Traz_TitReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then gError 177891
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 177891
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177842)

    End Select

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long, bCancela As Boolean
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaReceber = obj1

    If Not (objParcelaReceber Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcelaReceber.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bCancela)
    End If

    Me.Show

    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177843)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSequencial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objParcelasRecDif As ClassParcelasRecDif

On Error GoTo Erro_objEventoSequencial_evSelecao

    Set objParcelasRecDif = obj1
    
    Sequencial.Caption = CInt(objParcelasRecDif.iSeq)
    
    lErro = Traz_ParcelasRecDif_Tela(objParcelasRecDif)
    If lErro <> SUCESSO Then gError 177892
        
    Me.Show
    
    Exit Sub
    
Erro_objEventoSequencial_evSelecao:

    Select Case gErr
    
        Case 177892
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177844)
     
     End Select
     
     Exit Sub

End Sub

Private Sub objEventoTipoDoc_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177845)
     
     End Select
     
     Exit Sub

End Sub

Private Sub Parcela_Change()

    iParcelaAlterada = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcela_GotFocus()
Dim iParcelaAux As Integer
    
    iParcelaAux = iParcelaAlterada
    Call MaskEdBox_TrataGotFocus(Parcela, iAlterado)
    iParcelaAlterada = iParcelaAux

End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Verifica se está preenchido
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Sub
    
    If iParcelaAnterior <> StrParaInt(Parcela.Text) Then
    
        Sequencial.Caption = ""
    
    End If

    'Critica se é Long positivo
    lErro = Valor_Positivo_Critica(Parcela.ClipText)
    If lErro <> SUCESSO Then gError 177892

    If iParcelaAlterada Then
    
        'Se Parcela foi alterada zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iParcelaAlterada = 0
    
    End If
    
    iParcelaAnterior = StrParaInt(Parcela.Text)

    Exit Sub

Erro_Parcela_Validate:

    Select Case gErr

        Case 177892
            Cancel = True

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177846)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()

    iTipoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

Dim lErro As Long

On Error GoTo Erro_Tipo_Click

    iAlterado = REGISTRO_ALTERADO
    
    If Tipo.ListIndex = -1 Then Exit Sub

    'Se o Tipo foi alterado zera glNumIntParc
    glNumIntParc = 0
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_Tipo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177847)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub
    
    If sTipoAnterior <> SCodigo_Extrai(Tipo.Text) Then

        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
        SaldoTitulo.Caption = ""
        ValorTitulo.Caption = ""
        Vencimento.Caption = ""
        SaldoParc.Caption = ""
        ValorParc.Caption = ""
        
    End If

    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then
        Call Verifica_Alteracao
        
        sTipoAnterior = SCodigo_Extrai(Tipo.Text)
        Exit Sub
    End If
    
    'Tenta localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then gError 177893

    'Se não encontrar -> Erro
    If lErro = 60483 Then gError 177894

    If iTipoAlterado Then
        
        'Se Tipo foi alterado zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iTipoAlterado = 0

    End If
    
    sTipoAnterior = SCodigo_Extrai(Tipo.Text)

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case gErr

        Case 177893

        Case 177894
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, Tipo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177848)

    End Select

    Exit Sub

End Sub

Private Function Verifica_Alteracao() As Long
'tenta obter o NumInt da parcela e trazer seus dados para a tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objParcelasRecDif As New ClassParcelasRecDif

On Error GoTo Erro_Verifica_Alteracao

    Vencimento.Caption = ""
    SaldoParc.Caption = ""
    ValorParc.Caption = ""

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Text)) = 0 Then Exit Function

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Function

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Function

    'Verifica preenchimento de NumTítulo
    If Len(Trim(Numero.Text)) = 0 Then Exit Function

    objCliente.sNomeReduzido = Cliente.Text

    'Lê Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 177895

    'Se não encontrou o Cliente --> Erro
    If lErro <> SUCESSO Then gError 177896

   'Preenche objTituloReceber
    gobjTituloReceber.iFilialEmpresa = giFilialEmpresa
    gobjTituloReceber.lCliente = objCliente.lCodigo
    gobjTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    gobjTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    gobjTituloReceber.lNumTitulo = CLng(Numero.Text)
    gobjTituloReceber.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Pesquisa no BD o Título Receber
    lErro = CF("TituloReceber_Le_SemNumIntDoc", gobjTituloReceber)
    If lErro <> SUCESSO And lErro <> 28574 Then gError 177897

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le_Numero", gobjTituloReceber)
        If lErro <> SUCESSO And lErro <> 28574 Then gError 177897
    
    End If

    'Se não encontrou o Título --> Erro
    If lErro <> SUCESSO Then gError 177898

    SaldoTitulo.Caption = Format(gobjTituloReceber.dSaldo, "STANDARD")
    ValorTitulo.Caption = Format(gobjTituloReceber.dValor, "STANDARD")

    'Verifica preenchimento da Parcela
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Function

    'Preenche objParcelaReceber
    gobjParcelaReceber.lNumIntTitulo = gobjTituloReceber.lNumIntDoc
    gobjParcelaReceber.iNumParcela = CInt(Parcela.Text)

    'Pesquisa no BD a Parcela
    lErro = CF("ParcelaReceber_Le_SemNumIntDoc", gobjParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28590 Then gError 177899

    If lErro <> SUCESSO Then

        'Verifica se é uma Parcela Baixada
        lErro = CF("ParcelaReceberBaixada_Le_SemNumIntDoc", gobjParcelaReceber)
        If lErro <> SUCESSO And lErro <> 28567 Then gError 177901
        
    End If

    'Se encontrou a Parcela Receber Baixada --> Erro
    If lErro <> SUCESSO Then gError 177900

    Vencimento.Caption = Format(gobjParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
    SaldoParc.Caption = Format(gobjParcelaReceber.dSaldo, "STANDARD")
    ValorParc.Caption = Format(gobjParcelaReceber.dValor, "STANDARD")
        
    If Len(Trim(Sequencial.Caption)) = 0 Then

        glNumIntParc = gobjParcelaReceber.lNumIntDoc

        iSequencialAlterado = 0

    Else
        If gobjParcelaReceber.lNumIntDoc <> 0 Then
            glNumIntParc = gobjParcelaReceber.lNumIntDoc
        
            objParcelasRecDif.lNumIntParc = glNumIntParc
            objParcelasRecDif.iSeq = StrParaInt(Sequencial.Caption)
                       
            lErro = Traz_ParcelasRecDif_Tela(objParcelasRecDif)
            If lErro <> SUCESSO Then gError 177903
        
        End If
    End If

    Verifica_Alteracao = SUCESSO

    Exit Function

Erro_Verifica_Alteracao:

    Verifica_Alteracao = gErr

    Select Case gErr

        Case 177895, 177897, 177899, 177901, 177903

        Case 177896
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case 177898
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO2", gErr, gobjTituloReceber.iFilialEmpresa, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, gobjTituloReceber.sSiglaDocumento, gobjTituloReceber.lNumTitulo)

        Case 177900
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA", gErr, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case 177902
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_BAIXADA", gErr, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177849)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Parcela Then
            Call LabelParcela_Click
        ElseIf Me.ActiveControl Is Sequencial Then
            Call LabelSequencial_Click
        ElseIf Me.ActiveControl Is CodTipoDif Then
            Call LabelCodTipoDif_Click
        End If
    
    End If
    
End Sub

Function Traz_TitReceber_Tela(objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Traz_TitReceber_Tela
    
    'Lê o Título à Receber
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 177904

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 56568 Then gError 177904

    End If

    'Não encontrou o Título à Receber --> erro
    If lErro <> SUCESSO Then gError 177905
'
'    If Len(Trim(Cliente.Text)) <> 0 Then
'        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
'        If lErro <> SUCESSO Then gError 177922
'    End If

    'Coloca o Cliente na Tela
    If objCliente.lCodigo <> objTituloReceber.lCliente Then
    
        Filial.Clear
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
    
        Cliente.Text = CStr(objTituloReceber.lCliente)
        Call Cliente_Validate(bSGECancelDummy)
    End If

    'Coloca a Filial na Tela
    If Codigo_Extrai(Filial.Text) <> objTituloReceber.iFilial Then
        
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
        
        Filial.Text = CStr(objTituloReceber.iFilial)
        Call Filial_Validate(bSGECancelDummy)
    End If
    
    'Coloca o Tipo na tela
    If SCodigo_Extrai(Tipo.Text) <> objTituloReceber.sSiglaDocumento Then
        
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
        
        Tipo.Text = objTituloReceber.sSiglaDocumento
        Call Tipo_Validate(bSGECancelDummy)
    End If

    If StrParaLong(Numero.Text) <> objTituloReceber.lNumTitulo Then
    
        Parcela.PromptInclude = False
        Parcela.Text = ""
        Parcela.PromptInclude = True
        Sequencial.Caption = ""
    
        If objTituloReceber.lNumTitulo = 0 Then
        
            Numero.PromptInclude = False
            Numero.Text = ""
            Numero.PromptInclude = True
            
            SaldoTitulo.Caption = ""
            ValorTitulo.Caption = ""
            
        Else
            Numero.PromptInclude = False
            Numero.Text = CStr(objTituloReceber.lNumTitulo)
            Numero.PromptInclude = True
            
            SaldoTitulo.Caption = Format(objTituloReceber.dSaldo, "STANDARD")
            ValorTitulo.Caption = Format(objTituloReceber.dValor, "STANDARD")
    
        End If
    
        Call Numero_Validate(bSGECancelDummy)
    
    End If
    
    Me.Show

Traz_TitReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TitReceber_Tela:

    Traz_TitReceber_Tela = gErr

    Select Case gErr

        Case 177904
        
        Case 177905
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177850)

    End Select

    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 177906

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 177906

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177851)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCodTipoDif_Click()

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodTipoDif_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(CodTipoDif.Text)) <> 0 Then

        objTiposDifParcRec.iCodigo = StrParaInt(CodTipoDif.Text)

    End If

    Call Chama_Tela("TiposDifParcRecLista", colSelecao, objTiposDifParcRec, objEventoCodTipoDif)

    Exit Sub

Erro_LabelCodTipoDif_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177907)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCodTipoDif_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDifParcRec As ClassTiposDifParcRec

On Error GoTo Erro_objEventoCodTipoDif_evSelecao

    Set objTiposDifParcRec = obj1

    CodTipoDif.Text = objTiposDifParcRec.iCodigo
    
    Call CodTipoDif_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCodTipoDif_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177908)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimparSeq_Click()
    Sequencial.Caption = ""
End Sub

Function Traz_ParcelasRecDif_Tela2(objParcelasRecDif As ClassParcelasRecDif) As Long

Dim lErro As Long
Dim objParcRec As New ClassParcelaReceber
Dim objTitRec As New ClassTituloReceber

On Error GoTo Erro_Traz_ParcelasRecDif_Tela2

    objParcRec.lNumIntDoc = objParcelasRecDif.lNumIntParc
    
    lErro = CF("ParcelaReceber_Le", objParcRec)
    If lErro <> SUCESSO And lErro <> 19147 Then gError 177908
    
    If lErro <> SUCESSO Then

        lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
        If lErro <> SUCESSO And lErro <> 58559 Then gError 177908

    End If
    
    objTitRec.lNumIntDoc = objParcRec.lNumIntTitulo
    
    lErro = Traz_TitReceber_Tela(objTitRec)
    If lErro <> SUCESSO Then gError 177909
    
    glNumIntParc = objParcelasRecDif.lNumIntParc

    Parcela.PromptInclude = False
    Parcela.Text = CStr(objParcRec.iNumParcela)
    Parcela.PromptInclude = True
    
    Vencimento.Caption = Format(objParcRec.dtDataVencimento, "dd/mm/yyyy")
    SaldoParc.Caption = Format(objParcRec.dSaldo, "STANDARD")
    ValorParc.Caption = Format(objParcRec.dValor, "STANDARD")

    Sequencial.Caption = ""
    Data.Caption = ""

    DescricaoTipoDif.Caption = ""
    
    Observacao.Text = ""

    iAlterado = 0

    Traz_ParcelasRecDif_Tela2 = SUCESSO

    Exit Function

Erro_Traz_ParcelasRecDif_Tela2:

    Traz_ParcelasRecDif_Tela2 = gErr

    Select Case gErr

        Case 177863, 177908, 177909
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177826)

    End Select

    Exit Function

End Function

