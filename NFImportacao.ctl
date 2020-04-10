VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl NFImportacao 
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   ScaleHeight     =   6345
   ScaleWidth      =   10095
   Begin VB.TextBox Fabricante 
      Height          =   285
      Left            =   3555
      TabIndex        =   33
      Top             =   180
      Width           =   1410
   End
   Begin VB.CheckBox FiltrarFabricante 
      Caption         =   "Filtrar Itens do Fabricante:"
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
      Left            =   930
      TabIndex        =   32
      Top             =   210
      Width           =   2670
   End
   Begin VB.CommandButton BotaoTrazerDados 
      Caption         =   "Trazer Dados"
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
      Left            =   5475
      TabIndex        =   18
      Top             =   600
      Width           =   1470
   End
   Begin VB.Frame FrameCusto 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   5010
      TabIndex        =   29
      Top             =   525
      Visible         =   0   'False
      Width           =   2235
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Left            =   885
         TabIndex        =   30
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   180
         Width           =   510
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Caption         =   "Itens da Nota Fiscal"
      Height          =   4650
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1395
      Visible         =   0   'False
      Width           =   9675
      Begin MSMask.MaskEdBox IPIValorUnitario 
         Height          =   300
         Left            =   2325
         TabIndex        =   28
         Top             =   1500
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ICMSPercRedBase 
         Height          =   300
         Left            =   7125
         TabIndex        =   27
         Top             =   585
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DespImpValorRateado 
         Height          =   225
         Left            =   7860
         TabIndex        =   26
         Top             =   1020
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ICMSAliquota 
         Height          =   300
         Left            =   6945
         TabIndex        =   25
         Top             =   1035
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox IPIAliquota 
         Height          =   300
         Left            =   6075
         TabIndex        =   24
         Top             =   990
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin VB.TextBox ItemNFAdicao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2940
         MaxLength       =   50
         TabIndex        =   23
         Top             =   630
         Width           =   1155
      End
      Begin VB.TextBox ItemNFItemAdicao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1785
         MaxLength       =   50
         TabIndex        =   22
         Top             =   630
         Width           =   1020
      End
      Begin VB.TextBox UnidadeMed 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   465
         MaxLength       =   50
         TabIndex        =   21
         Top             =   615
         Width           =   1215
      End
      Begin MSMask.MaskEdBox ValorUnitario 
         Height          =   225
         Left            =   4440
         TabIndex        =   20
         Top             =   1035
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   2865
         TabIndex        =   19
         Top             =   1065
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   10
         Top             =   225
         Width           =   2265
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   165
         TabIndex        =   11
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemNFValorTotal 
         Height          =   225
         Left            =   7740
         TabIndex        =   12
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemNFValorAduaneiro 
         Height          =   225
         Left            =   3900
         TabIndex        =   13
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemNFValorII 
         Height          =   225
         Left            =   6420
         TabIndex        =   14
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItensNF 
         Height          =   4545
         Left            =   45
         TabIndex        =   9
         Top             =   75
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   8017
         _Version        =   393216
      End
   End
   Begin VB.Frame FrameOpcao 
      BorderStyle     =   0  'None
      Caption         =   "Valores Complementares para a Nota Fiscal"
      Height          =   4635
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1410
      Width           =   9645
      Begin VB.TextBox ComplNFTipo 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   435
         Width           =   1035
      End
      Begin VB.TextBox ComplNFDescricao 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1755
         TabIndex        =   4
         Top             =   405
         Width           =   4710
      End
      Begin MSMask.MaskEdBox ComplNFValor 
         Height          =   225
         Left            =   6585
         TabIndex        =   6
         Top             =   465
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridComplNF 
         Height          =   4530
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   7990
         _Version        =   393216
      End
   End
   Begin VB.CommandButton BotaoRetornar 
      Caption         =   "Retornar"
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
      Left            =   7350
      TabIndex        =   17
      Top             =   600
      Width           =   960
   End
   Begin VB.TextBox DI 
      Height          =   330
      Left            =   705
      TabIndex        =   0
      Top             =   645
      Width           =   2130
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5100
      Left            =   90
      TabIndex        =   7
      Top             =   1035
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   8996
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resumo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens NF"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Ao mexer no layout da tela verifique a função Trata_Parametros pois ela reposiciona os controles para complemento de custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   5085
      TabIndex        =   34
      Top             =   -45
      Visible         =   0   'False
      Width           =   4965
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   2985
      TabIndex        =   16
      Top             =   690
      Width           =   510
   End
   Begin VB.Label DataDI 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3555
      TabIndex        =   15
      Top             =   645
      Width           =   1425
   End
   Begin VB.Label LabelDI 
      Caption         =   "D.I.:"
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
      Left            =   225
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   705
      Width           =   420
   End
End
Attribute VB_Name = "NFImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? pedir confirmacao para sobrepor se já houver dados

Private Const NUM_MAX_LINHAS_GRID_COMPLNF = 50
Private Const NUM_MAX_LINHAS_GRID_ITENSNF = 990

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Dim iAlterado As Integer

Dim objGridComplNF As AdmGrid
Dim iGrid_ComplNFTipo_Col As Integer
Dim iGrid_ComplNFDescricao_Col As Integer
Dim iGrid_ComplNFValor_Col As Integer

Dim objGridItensNF As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_ItemNFValorAduaneiro_Col As Integer
Dim iGrid_ItemNFValorII_Col As Integer
Dim iGrid_ItemNFValorTotal_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_ValorUnitario_Col As Integer
Dim iGrid_ItemNFAdicao_Col As Integer
Dim iGrid_ItemNFItemAdicao_Col As Integer
Dim iGrid_IPIAliquota_Col As Integer
Dim iGrid_ICMSAliquota_Col As Integer
Dim iGrid_ICMSPercRedBase_Col As Integer
Dim iGrid_DespImpValorRateado_Col As Integer
Dim iGrid_IPIValorUnitario_Col As Integer

Private WithEvents objEventoDI As AdmEvento
Attribute objEventoDI.VB_VarHelpID = -1

Private gobjNFiscal As ClassNFiscal

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Informações sobre a Importação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFImportacao"

End Function

Public Sub Show()
'???? comentei para nao dar erro nesta tela pq é modal
'    Parent.Show
 '   Parent.SetFocus
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

Private Sub BotaoRetornar_Click()
    
    Call Move_Tela_Memoria
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK

    Unload Me
    
End Sub

Private Sub BotaoTrazerDados_Click()

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo

On Error GoTo Erro_BotaoTrazerDados_Click

    If Len(Trim(DI.Text)) <> 0 Then

        objDIInfo.sNumero = DI.Text
        
        If gobjNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
            'If StrParaDbl(Valor.Text) = 0 Then gError 211321
        End If
        
        'Mostra os dados do DIInfo na tela
        lErro = Traz_DIInfo_Tela(objDIInfo)
        If lErro <> SUCESSO Then gError 184590

    End If
    
    Exit Sub

Erro_BotaoTrazerDados_Click:

    Select Case gErr

        Case 184590
        
        Case 211321
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORRATEIO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184591)

    End Select

    Exit Sub

End Sub


Private Sub LabelDI_Click()

Dim lErro As Long
Dim objDIInfo As New ClassDIInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(DI.Text)) <> 0 Then

        objDIInfo.sNumero = DI.Text

    End If

    Call Chama_Tela_Modal("DIInfoLista", colSelecao, objDIInfo, objEventoDI)

    Exit Sub

Erro_LabelNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184592)

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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    iFrameAtual = 1

    Set objEventoDI = Nothing
    
    Set objGridComplNF = Nothing
    Set objGridItensNF = Nothing

    Set gobjNFiscal = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184529)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoDI = New AdmEvento
    
    lErro = Inicializa_GridComplNF(objGridComplNF)
    If lErro <> SUCESSO Then gError 184530
    
    lErro = Inicializa_GridItensNF(objGridItensNF)
    If lErro <> SUCESSO Then gError 184531

    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 196566

    iFrameAtual = 1
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 184530, 184531, 196566

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184532)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objNFiscal Is Nothing) Then

        Set gobjNFiscal = objNFiscal
        
        If gobjNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
            BotaoTrazerDados.left = 7350
            BotaoRetornar.left = 8955
            FrameCusto.Visible = True
            
            If objNFiscal.dValorProdutos <> 0 Then Valor.Text = Format(objNFiscal.dValorProdutos, "STANDARD")
        End If
        
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 184533

    Else
    
        gError 184701
        
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel
    
    Trata_Parametros = gErr

    Select Case gErr

        Case 184533

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184534)

    End Select

    Exit Function

End Function

Private Function Traz_NFiscal_Tela(ByVal objNFiscal As ClassNFiscal) As Long

    Call Traz_NFImportacao_Tela(objNFiscal.objNFImportacao)

End Function

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameOpcao(Opcao.SelectedItem.Index).Visible = True
        FrameOpcao(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

End Sub

Private Function Inicializa_GridItensNF(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descricao")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Valor Unitário")
    objGrid.colColuna.Add ("Valor Total")
    objGrid.colColuna.Add ("Rateio Desp.")
    objGrid.colColuna.Add ("Total CIF")
    objGrid.colColuna.Add ("Valor I.I.")
    objGrid.colColuna.Add ("IPI %")
    objGrid.colColuna.Add ("IPI Unit.")
    objGrid.colColuna.Add ("ICMS %")
    objGrid.colColuna.Add ("ICMS % Red.")
    objGrid.colColuna.Add ("Adição")
    objGrid.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (ValorUnitario.Name)
    objGrid.colCampo.Add (ItemNFValorTotal.Name)
    objGrid.colCampo.Add (DespImpValorRateado.Name)
    objGrid.colCampo.Add (ItemNFValorAduaneiro.Name)
    objGrid.colCampo.Add (ItemNFValorII.Name)
    objGrid.colCampo.Add (IPIAliquota.Name)
    objGrid.colCampo.Add (IPIValorUnitario.Name)
    objGrid.colCampo.Add (ICMSAliquota.Name)
    objGrid.colCampo.Add (ICMSPercRedBase.Name)
    objGrid.colCampo.Add (ItemNFAdicao.Name)
    objGrid.colCampo.Add (ItemNFItemAdicao.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_ValorUnitario_Col = 5
    iGrid_ItemNFValorTotal_Col = 6
    iGrid_DespImpValorRateado_Col = 7
    iGrid_ItemNFValorAduaneiro_Col = 8
    iGrid_ItemNFValorII_Col = 9
    iGrid_IPIAliquota_Col = 10
    iGrid_IPIValorUnitario_Col = 11
    iGrid_ICMSAliquota_Col = 12
    iGrid_ICMSPercRedBase_Col = 13
    iGrid_ItemNFAdicao_Col = 14
    iGrid_ItemNFItemAdicao_Col = 15

    objGrid.objGrid = GridItensNF

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_LINHAS_GRID_ITENSNF + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 11

    'Largura da primeira coluna
    GridItensNF.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItensNF = SUCESSO

End Function

Private Sub GridItensNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensNF, iAlterado)
    End If

End Sub

Private Sub GridItensNF_GotFocus()
    Call Grid_Recebe_Foco(objGridItensNF)
End Sub

Private Sub GridItensNF_EnterCell()
    Call Grid_Entrada_Celula(objGridItensNF, iAlterado)
End Sub

Private Sub GridItensNF_LeaveCell()
    Call Saida_Celula(objGridItensNF)
End Sub

Private Sub GridItensNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensNF, iAlterado)
    End If

End Sub

Private Sub GridItensNF_RowColChange()
    Call Grid_RowColChange(objGridItensNF)
End Sub

Private Sub GridItensNF_Scroll()
    Call Grid_Scroll(objGridItensNF)
End Sub

Private Sub GridItensNF_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItensNF)
End Sub

Private Sub GridItensNF_LostFocus()
    Call Grid_Libera_Foco(objGridItensNF)
End Sub

Private Sub GridComplNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComplNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComplNF, iAlterado)
    End If

End Sub

Private Sub GridComplNF_GotFocus()
    Call Grid_Recebe_Foco(objGridComplNF)
End Sub

Private Sub GridComplNF_EnterCell()
    Call Grid_Entrada_Celula(objGridComplNF, iAlterado)
End Sub

Private Sub GridComplNF_LeaveCell()
    Call Saida_Celula(objGridComplNF)
End Sub

Private Sub GridComplNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComplNF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComplNF, iAlterado)
    End If

End Sub

Private Sub GridComplNF_RowColChange()
    Call Grid_RowColChange(objGridComplNF)
End Sub

Private Sub GridComplNF_Scroll()
    Call Grid_Scroll(objGridComplNF)
End Sub

Private Sub GridComplNF_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridComplNF)
End Sub

Private Sub GridComplNF_LostFocus()
    Call Grid_Libera_Foco(objGridComplNF)
End Sub

Private Function Inicializa_GridComplNF(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ComplNFTipo.Name)
    objGrid.colCampo.Add (ComplNFDescricao.Name)
    objGrid.colCampo.Add (ComplNFValor.Name)

    'Colunas do Grid
    iGrid_ComplNFTipo_Col = 1
    iGrid_ComplNFDescricao_Col = 2
    iGrid_ComplNFValor_Col = 3
    
    objGrid.objGrid = GridComplNF

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_LINHAS_GRID_COMPLNF + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridComplNF.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridComplNF = SUCESSO

End Function

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescricaoItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemNFValorAduaneiro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemNFValorAduaneiro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ItemNFValorAduaneiro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ItemNFValorAduaneiro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ItemNFValorAduaneiro
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemNFValorII_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemNFValorII_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ItemNFValorII_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ItemNFValorII_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ItemNFValorII
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemNFValorTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemNFValorTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ItemNFValorTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ItemNFValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ItemNFValorTotal
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemNFAdicao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemNFAdicao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ItemNFAdicao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ItemNFAdicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ItemNFAdicao
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemNFItemAdicao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemNFItemAdicao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ItemNFItemAdicao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ItemNFItemAdicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ItemNFItemAdicao
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplNFTipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplNFTipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComplNF)
End Sub

Private Sub ComplNFTipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComplNF)
End Sub

Private Sub ComplNFTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComplNF.objControle = ComplNFTipo
    lErro = Grid_Campo_Libera_Foco(objGridComplNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplNFDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplNFDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComplNF)
End Sub

Private Sub ComplNFDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComplNF)
End Sub

Private Sub ComplNFDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComplNF.objControle = ComplNFDescricao
    lErro = Grid_Campo_Libera_Foco(objGridComplNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ComplNFValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComplNFValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComplNF)
End Sub

Private Sub ComplNFValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComplNF)
End Sub

Private Sub ComplNFValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComplNF.objControle = ComplNFValor
    lErro = Grid_Campo_Libera_Foco(objGridComplNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'GridItensNF
        If objGridInt.objGrid.Name = GridItensNF.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 184629

                Case iGrid_DescricaoItem_Col

                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ItemNFValorAduaneiro_Col

                    lErro = Saida_Celula_ItemNFValorAduaneiro(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ItemNFAdicao_Col

                    lErro = Saida_Celula_ItemNFAdicao(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ItemNFItemAdicao_Col

                    lErro = Saida_Celula_ItemNFItemAdicao(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ItemNFValorII_Col

                    lErro = Saida_Celula_ItemNFValorII(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ItemNFValorTotal_Col

                    lErro = Saida_Celula_ItemNFValorTotal(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 184630

                Case iGrid_ValorUnitario_Col
                    lErro = Saida_Celula_ValorUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
                Case iGrid_IPIAliquota_Col
                    lErro = Saida_Celula_IPIAliquota(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
                Case iGrid_IPIValorUnitario_Col
                    lErro = Saida_Celula_IPIValorUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
                Case iGrid_ICMSAliquota_Col
                    lErro = Saida_Celula_ICMSAliquota(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
                Case iGrid_ICMSPercRedBase_Col
                    lErro = Saida_Celula_ICMSPercRedBase(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
                Case iGrid_DespImpValorRateado_Col
                    lErro = Saida_Celula_DespImpValorRateado(objGridInt)
                    If lErro <> SUCESSO Then gError 184630
            
            End Select
                    
        End If

        'GridComplNF
        If objGridInt.objGrid.Name = GridComplNF.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_ComplNFTipo_Col

                    lErro = Saida_Celula_ComplNFTipo(objGridInt)
                    If lErro <> SUCESSO Then gError 184625
                
                Case iGrid_ComplNFDescricao_Col

                    lErro = Saida_Celula_ComplNFDescricao(objGridInt)
                    If lErro <> SUCESSO Then gError 184625
                
                Case iGrid_ComplNFValor_Col

                    lErro = Saida_Celula_ComplNFValor(objGridInt)
                    If lErro <> SUCESSO Then gError 184625
                
            End Select
                    
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 184641

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 184625 To 184640

        Case 184641
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184642)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
        
        Case Produto.Name, DescricaoItem.Name, UnidadeMed.Name, ItemNFValorTotal.Name, _
            ItemNFValorAduaneiro.Name, ItemNFValorII.Name, IPIAliquota.Name, IPIValorUnitario.Name, ICMSAliquota.Name, ICMSPercRedBase.Name, _
            ICMSAliquota.Name, ItemNFAdicao.Name, ItemNFItemAdicao.Name, ValorUnitario.Name, DespImpValorRateado.Name
            objControl.Enabled = False
        
        Case Quantidade.Name
            objControl.Enabled = True
            
        Case Else
            objControl.Enabled = True

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184643)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemNFValorAduaneiro(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFValorAduaneiro

    Set objGridInt.objControle = ItemNFValorAduaneiro

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ItemNFValorAduaneiro = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFValorAduaneiro:

    Saida_Celula_ItemNFValorAduaneiro = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemNFAdicao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFAdicao

    Set objGridInt.objControle = ItemNFAdicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ItemNFAdicao = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFAdicao:

    Saida_Celula_ItemNFAdicao = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemNFItemAdicao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFItemAdicao

    Set objGridInt.objControle = ItemNFItemAdicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ItemNFItemAdicao = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFItemAdicao:

    Saida_Celula_ItemNFItemAdicao = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemNFValorII(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFValorII

    Set objGridInt.objControle = ItemNFValorII

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ItemNFValorII = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFValorII:

    Saida_Celula_ItemNFValorII = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemNFValorTotal(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFValorTotal

    Set objGridInt.objControle = ItemNFValorTotal

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ItemNFValorTotal = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFValorTotal:

    Saida_Celula_ItemNFValorTotal = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ComplNFTipo(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComplNFTipo

    Set objGridInt.objControle = ComplNFTipo

    If (GridComplNF.Row - GridComplNF.FixedRows) = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ComplNFTipo = SUCESSO

    Exit Function

Erro_Saida_Celula_ComplNFTipo:

    Saida_Celula_ComplNFTipo = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ComplNFDescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComplNFDescricao

    Set objGridInt.objControle = ComplNFDescricao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ComplNFDescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_ComplNFDescricao:

    Saida_Celula_ComplNFDescricao = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ComplNFValor(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComplNFValor

    Set objGridInt.objControle = ComplNFValor

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ComplNFValor = SUCESSO

    Exit Function

Erro_Saida_Celula_ComplNFValor:

    Saida_Celula_ComplNFValor = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Sub UnidadeMed_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UnidadeMed_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ValorUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub objEventoDI_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDIInfo As ClassDIInfo

On Error GoTo Erro_objEventoDI_evSelecao

    Set objDIInfo = obj1

    'Mostra os dados do DIInfo na tela
    lErro = Traz_DIInfo_Tela(objDIInfo)
    If lErro <> SUCESSO Then gError 184590

    Me.Show

    Exit Sub

Erro_objEventoDI_evSelecao:

    Select Case gErr

        Case 184590

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184591)

    End Select

    Exit Sub

End Sub

Private Function Traz_DIInfo_Tela(objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, objImportCompl As ClassImportCompl, dValorAduaneiroDI As Double, dDespImpAdicaoDI As Double
Dim objImportComplAux As ClassImportCompl, objNFImportacao As ClassNFImportacao
Dim objAdicaoDI As ClassAdicaoDI, objItemAdicaoDI As ClassItemAdicaoDI, objTipoDespesa As New ClassTipoImportCompl
Dim objAdicaoDINF As ClassAdicaoDINF, objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF
Dim iSeqComplAdicao As Integer, iSeqComplNF As Integer, iSeqItemAdicao As Integer
Dim dValorRatear As Double, dValorFalta As Double

On Error GoTo Erro_Traz_DIInfo_Tela

    Call Limpa_Tela_DIInfo
    
    'Lê o DIInfo que está sendo Passado
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184546

    If lErro = SUCESSO Then

        Set objNFImportacao = New ClassNFImportacao
        Set objNFImportacao.objDIInfo = objDIInfo
        
        dValorAduaneiroDI = Arredonda_Moeda(objDIInfo.dValorMercadoriaEmReal + objDIInfo.dValorFreteInternacEmReal + objDIInfo.dValorSeguroInternacEmReal)
        
        If objDIInfo.dIIValor <> 0 Then
        
            iSeqComplNF = iSeqComplNF + 1
            Set objImportCompl = New ClassImportCompl
                    
            With objImportCompl
                
                .iSeq = iSeqComplNF
                .iTipo = IMPORTCOMPL_TIPO_II
                .sDescricao = "IMPOSTO DE IMPORTAÇÂO"
                .dValor = objDIInfo.dIIValor
                
                .iManual = 0
            
            End With
            
            objNFImportacao.colComplNF.Add objImportCompl
            
        End If

        If objDIInfo.dPISValor <> 0 Then
        
            iSeqComplNF = iSeqComplNF + 1
            Set objImportCompl = New ClassImportCompl
                    
            With objImportCompl
                
                .iSeq = iSeqComplNF
                .iTipo = IMPORTCOMPL_TIPO_PIS
                .sDescricao = "PIS IMPORTAÇÂO"
                .dValor = objDIInfo.dPISValor
                
                .iManual = 0
            
            End With
            
            objNFImportacao.colComplNF.Add objImportCompl
            
        End If

        If objDIInfo.dCOFINSValor <> 0 Then
        
            iSeqComplNF = iSeqComplNF + 1
            Set objImportCompl = New ClassImportCompl
                    
            With objImportCompl
                
                .iSeq = iSeqComplNF
                .iTipo = IMPORTCOMPL_TIPO_COFINS
                .sDescricao = "COFINS IMPORTAÇÂO"
                .dValor = objDIInfo.dCOFINSValor
                
                .iManual = 0
            
            End With
            
            objNFImportacao.colComplNF.Add objImportCompl
        
        End If

        For Each objImportComplAux In objDIInfo.colDespesasDI
        
            iSeqComplNF = iSeqComplNF + 1
            Set objImportCompl = New ClassImportCompl
                    
            With objImportCompl
                
                .iSeq = iSeqComplNF
                .iTipo = objImportComplAux.iTipo
                .sDescricao = objImportComplAux.sDescricao
                .dValor = objImportComplAux.dValor
                
                .iManual = 0
            
            End With
            
            objTipoDespesa.iCodigo = objImportComplAux.iTipo
            lErro = CF("TiposImportCompl_Le", objTipoDespesa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184890
            If lErro <> SUCESSO Then gError 184891
            
            If objTipoDespesa.iIncluiBaseICMS <> 0 Then
                objNFImportacao.colComplNF.Add objImportCompl
            End If
            
            If objTipoDespesa.iIncluiNoValorAduaneiro <> 0 Then
                dValorAduaneiroDI = Arredonda_Moeda(dValorAduaneiroDI + objImportCompl.dValor)
            End If
            
        Next

        For Each objAdicaoDI In objDIInfo.colAdicoesDI
        
            If FiltrarFabricante.Value = vbUnchecked Or Fabricante.Text = objAdicaoDI.sCodFabricante Then
            
                Set objAdicaoDINF = New ClassAdicaoDINF
                
                objAdicaoDINF.iSeq = objAdicaoDI.iSeq
            
                iSeqComplAdicao = 0
                
                If objAdicaoDI.dIIValor <> 0 Then
                
                    iSeqComplAdicao = iSeqComplAdicao + 1
                    Set objImportCompl = New ClassImportCompl
                            
                    With objImportCompl
                        
                        .iSeq = iSeqComplAdicao
                        .iTipo = IMPORTCOMPL_TIPO_II
                        .sDescricao = "IMPOSTO DE IMPORTAÇÂO"
                        .dValorBase = objAdicaoDI.dValorAduaneiro
                        .dAliquota = objAdicaoDI.dIIAliquota
                        .dValor = objAdicaoDI.dIIValor
                        
                        .iManual = 0
                    
                    End With
                    
                    objAdicaoDINF.colComplNF.Add objImportCompl
                    
                End If
        
                If objAdicaoDI.dPISValor <> 0 Then
                
                    iSeqComplAdicao = iSeqComplAdicao + 1
                    Set objImportCompl = New ClassImportCompl
                            
                    With objImportCompl
                        
                        .iSeq = iSeqComplAdicao
                        .iTipo = IMPORTCOMPL_TIPO_PIS
                        .sDescricao = "PIS IMPORTAÇÂO"
                        .dValorBase = objAdicaoDI.dPISBase
                        .dAliquota = objAdicaoDI.dPISAliquota
                        .dValor = objAdicaoDI.dPISValor
                        
                        .iManual = 0
                    
                    End With
                    
                    objAdicaoDINF.colComplNF.Add objImportCompl
                    
                End If
        
                If objAdicaoDI.dCOFINSValor <> 0 Then
                
                    iSeqComplAdicao = iSeqComplAdicao + 1
                    Set objImportCompl = New ClassImportCompl
                            
                    With objImportCompl
                        
                        .iSeq = iSeqComplAdicao
                        .iTipo = IMPORTCOMPL_TIPO_COFINS
                        .sDescricao = "COFINS IMPORTAÇÂO"
                        .dValorBase = objAdicaoDI.dCOFINSBase
                        .dAliquota = objAdicaoDI.dCOFINSAliquota
                        .dValor = objAdicaoDI.dCOFINSValor
                        
                        .iManual = 0
                    
                    End With
                    
                    objAdicaoDINF.colComplNF.Add objImportCompl
                
                End If
                
                dDespImpAdicaoDI = Arredonda_Moeda(objAdicaoDI.dPISValor + objAdicaoDI.dCOFINSValor + objAdicaoDI.dDespesaAduaneira + objAdicaoDI.dTaxaSiscomex)
                
                For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
                    
                    Set objItemAdicaoDIItemNF = New ClassItemAdicaoDIItemNF
                    
                    With objItemAdicaoDIItemNF
                    
                        .iAdicao = objAdicaoDI.iSeq
                        .iItemAdicao = objItemAdicaoDI.iSeq
                        
                        .sProduto = objItemAdicaoDI.sProduto
                        .sDescricao = objItemAdicaoDI.sDescricao
                        .sUM = objItemAdicaoDI.sUM
                        .dQuantidade = objItemAdicaoDI.dQuantidade
                        
                        .dValorAduaneiro = objItemAdicaoDI.dValorTotalCIFEmReal
                        .dValorII = Arredonda_Moeda((.dValorAduaneiro / objAdicaoDI.dValorAduaneiro) * objAdicaoDI.dIIValor)
    
                        If gobjCRFAT.iNFImportacaoTribFlag12 = 0 Then
                            .dPrecoUnitario = Arredonda_Moeda(.dValorTotal / .dQuantidade, 6)
                        Else
                            .dPrecoUnitario = Arredonda_Moeda((.dValorTotal - .dValorII) / .dQuantidade, 6)
                        End If
                        
                        .dIPIAliquotaAdicaoDI = objAdicaoDI.dIPIAliquota
                        .dIPIUnidadePadraoValor = objItemAdicaoDI.dIPIUnidadePadraoValor
                        .dICMSAliquotaAdicaoDI = objAdicaoDI.dICMSAliquota
                        .dICMSPercRedBaseAdicaoDI = objAdicaoDI.dICMSPercRedBase
                        .dDespImpValorRateado = Arredonda_Moeda((.dValorAduaneiro / objAdicaoDI.dValorAduaneiro) * dDespImpAdicaoDI)
                        
                    End With
                    
                    objNFImportacao.ColItensNF.Add objItemAdicaoDIItemNF
                
                Next
                
                objNFImportacao.colAdicoesNF.Add objAdicaoDINF
            
            End If
            
        Next
            
        Set gobjNFiscal.objNFImportacao = objNFImportacao
        Call Traz_NFImportacao_Tela(objNFImportacao)
    
    End If

    Traz_DIInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_DIInfo_Tela:

    Traz_DIInfo_Tela = gErr

    Select Case gErr

        Case 184546, 184890
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184547)

    End Select

    Exit Function

End Function

Private Sub Traz_NFImportacao_Tela(ByVal objNFImportacao As ClassNFImportacao)

Dim objDIInfo As ClassDIInfo
Dim lErro As Long, objImportCompl As ClassImportCompl
Dim objAdicaoDINF As ClassAdicaoDINF, iIndice As Integer

    If Len(Trim(objNFImportacao.objDIInfo.sNumero)) <> 0 Then
    
        Set objDIInfo = objNFImportacao.objDIInfo
    
        DI.Text = objDIInfo.sNumero
    
        If objDIInfo.dtData <> DATA_NULA Then DataDI.Caption = Format(objDIInfo.dtData, "dd/mm/yy")
        
        'preencher 1o. tab
        For Each objImportCompl In objNFImportacao.colComplNF
        
            iIndice = iIndice + 1
            
            GridComplNF.TextMatrix(iIndice, iGrid_ComplNFTipo_Col) = CStr(objImportCompl.iTipo)
            GridComplNF.TextMatrix(iIndice, iGrid_ComplNFDescricao_Col) = objImportCompl.sDescricao
            GridComplNF.TextMatrix(iIndice, iGrid_ComplNFValor_Col) = Arredonda_Moeda(objImportCompl.dValor)
        
        Next
        
        'Atualiza o número de linhas existentes
        objGridComplNF.iLinhasExistentes = iIndice
        
        'se o grid de itens for independente da adicao selecionada
        Call Traz_Itens_Tela
    
    End If
    
End Sub

Private Sub Limpa_Tela_DIInfo()

    'Limpar Grids
    Call Grid_Limpa(objGridComplNF)
    Call Grid_Limpa(objGridItensNF)

End Sub

Private Sub Traz_Itens_Tela()

Dim lErro As Long, iIndice As Integer
Dim objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF, sProdutoEnxuto As String

    Call Grid_Limpa(objGridItensNF)
    
    For Each objItemAdicaoDIItemNF In gobjNFiscal.objNFImportacao.ColItensNF
    
        iIndice = iIndice + 1
        
        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemAdicaoDIItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 35524

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        GridItensNF.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItensNF.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objItemAdicaoDIItemNF.sDescricao
        GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorAduaneiro_Col) = Arredonda_Moeda(objItemAdicaoDIItemNF.dValorAduaneiro)
        GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorII_Col) = Arredonda_Moeda(objItemAdicaoDIItemNF.dValorII)
        GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorTotal_Col) = Arredonda_Moeda(objItemAdicaoDIItemNF.dValorTotal)
        GridItensNF.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemAdicaoDIItemNF.sUM
        GridItensNF.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemAdicaoDIItemNF.dQuantidade)
        GridItensNF.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemAdicaoDIItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItensNF.TextMatrix(iIndice, iGrid_ItemNFAdicao_Col) = CStr(objItemAdicaoDIItemNF.iAdicao)
        GridItensNF.TextMatrix(iIndice, iGrid_ItemNFItemAdicao_Col) = CStr(objItemAdicaoDIItemNF.iItemAdicao)
        GridItensNF.TextMatrix(iIndice, iGrid_IPIAliquota_Col) = Format(objItemAdicaoDIItemNF.dIPIAliquotaAdicaoDI, "Percent")
        GridItensNF.TextMatrix(iIndice, iGrid_IPIValorUnitario_Col) = Format(objItemAdicaoDIItemNF.dIPIUnidadePadraoValor, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItensNF.TextMatrix(iIndice, iGrid_ICMSAliquota_Col) = Format(objItemAdicaoDIItemNF.dICMSAliquotaAdicaoDI, "Percent")
        GridItensNF.TextMatrix(iIndice, iGrid_ICMSPercRedBase_Col) = Format(objItemAdicaoDIItemNF.dICMSPercRedBaseAdicaoDI, "Percent")
        GridItensNF.TextMatrix(iIndice, iGrid_DespImpValorRateado_Col) = Format(objItemAdicaoDIItemNF.dDespImpValorRateado, FORMATO_PRECO_UNITARIO_EXTERNO)
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridItensNF.iLinhasExistentes = iIndice

End Sub

Sub Move_Tela_Memoria()

Dim iIndice As Integer, objNFImportacao As ClassNFImportacao, objImportCompl As ClassImportCompl
Dim objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF, objAdicaoDINF As ClassAdicaoDINF
Dim iAdicao As Integer, iItemAdicao As Integer, objTipoImportCompl As ClassTipoImportCompl
Dim sProdutoFormatado As String, lErro As Long
Dim iProdutoPreenchido As Integer
Dim objAdicaoDI As ClassAdicaoDI, objItemAdicaoDI As ClassItemAdicaoDI

    Set objNFImportacao = gobjNFiscal.objNFImportacao
    
    'copia o tab de resumo
    Set objNFImportacao.colComplNF = New Collection
    For iIndice = 1 To objGridComplNF.iLinhasExistentes
    
        Set objImportCompl = New ClassImportCompl
            
        objImportCompl.iSeq = iIndice
        objImportCompl.iTipo = StrParaInt(GridComplNF.TextMatrix(iIndice, iGrid_ComplNFTipo_Col))
        objImportCompl.sDescricao = GridComplNF.TextMatrix(iIndice, iGrid_ComplNFDescricao_Col)
        objImportCompl.dValor = StrParaDbl(GridComplNF.TextMatrix(iIndice, iGrid_ComplNFValor_Col))
    
        Set objTipoImportCompl = New ClassTipoImportCompl
    
        objTipoImportCompl.iCodigo = objImportCompl.iTipo
        
        lErro = CF("TiposImportCompl_Le", objTipoImportCompl)
        If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError 184703
        If lErro <> SUCESSO Then gError 184704
        
        Set objImportCompl.objTipoImportCompl = objTipoImportCompl
        
        objNFImportacao.colComplNF.Add objImportCompl
    
    Next

    'copia o tab de itens
    Set objNFImportacao.ColItensNF = New Collection
    For iIndice = 1 To objGridItensNF.iLinhasExistentes
                
        Set objItemAdicaoDIItemNF = New ClassItemAdicaoDIItemNF
        
        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItensNF.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 35576
        
        'Armazena produto
        objItemAdicaoDIItemNF.sProduto = sProdutoFormatado
            
        With objItemAdicaoDIItemNF
        
            .iItemNF = iIndice
            
            .sDescricao = GridItensNF.TextMatrix(iIndice, iGrid_DescricaoItem_Col)
            .dValorAduaneiro = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorAduaneiro_Col))
            .dValorII = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorII_Col))
            '.dValorTotal = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ItemNFValorTotal_Col))
            .sUM = GridItensNF.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            .dQuantidade = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_Quantidade_Col))
            .dPrecoUnitario = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
            .iAdicao = StrParaInt(GridItensNF.TextMatrix(iIndice, iGrid_ItemNFAdicao_Col))
            .iItemAdicao = StrParaInt(GridItensNF.TextMatrix(iIndice, iGrid_ItemNFItemAdicao_Col))
            .dIPIAliquotaAdicaoDI = PercentParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_IPIAliquota_Col))
            .dIPIUnidadePadraoValor = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_IPIValorUnitario_Col))
            .dICMSAliquotaAdicaoDI = PercentParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ICMSAliquota_Col))
            .dICMSPercRedBaseAdicaoDI = PercentParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_ICMSPercRedBase_Col))
            .dDespImpValorRateado = StrParaDbl(GridItensNF.TextMatrix(iIndice, iGrid_DespImpValorRateado_Col))
        
        End With
        
        For Each objAdicaoDI In objNFImportacao.objDIInfo.colAdicoesDI
                        
            For Each objItemAdicaoDI In objAdicaoDI.colItensAdicaoDI
            
                If objItemAdicaoDI.iAdicao = objItemAdicaoDIItemNF.iAdicao And objItemAdicaoDI.iSeq = objItemAdicaoDIItemNF.iItemAdicao Then
                    objItemAdicaoDIItemNF.lNumIntItemAdicaoDI = objItemAdicaoDI.lNumIntDoc
                    objItemAdicaoDIItemNF.dPISAliquotaAdicaoDI = objAdicaoDI.dPISAliquota
                    objItemAdicaoDIItemNF.dCOFINSAliquotaAdicaoDI = objAdicaoDI.dCOFINSAliquota
                    Exit For
                End If
                
            Next
        
            If objItemAdicaoDIItemNF.lNumIntItemAdicaoDI <> 0 Then Exit For
            
        Next
        
        If objItemAdicaoDIItemNF.lNumIntItemAdicaoDI = 0 Then gError 184711
        
        objNFImportacao.ColItensNF.Add objItemAdicaoDIItemNF
        
    Next
    
    gobjNFiscal.dValorProdutos = StrParaDbl(Valor.Text)
    
End Sub

Private Function Saida_Celula_IPIAliquota(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_IPIAliquota

    Set objGridInt.objControle = IPIAliquota

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_IPIAliquota = SUCESSO

    Exit Function

Erro_Saida_Celula_IPIAliquota:

    Saida_Celula_IPIAliquota = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ICMSAliquota(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ICMSAliquota

    Set objGridInt.objControle = ICMSAliquota

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ICMSAliquota = SUCESSO

    Exit Function

Erro_Saida_Celula_ICMSAliquota:

    Saida_Celula_ICMSAliquota = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ICMSPercRedBase(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ICMSPercRedBase

    Set objGridInt.objControle = ICMSPercRedBase

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_ICMSPercRedBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ICMSPercRedBase:

    Saida_Celula_ICMSPercRedBase = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DespImpValorRateado(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DespImpValorRateado

    Set objGridInt.objControle = DespImpValorRateado

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_DespImpValorRateado = SUCESSO

    Exit Function

Erro_Saida_Celula_DespImpValorRateado:

    Saida_Celula_DespImpValorRateado = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_IPIValorUnitario(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_IPIValorUnitario

    Set objGridInt.objControle = IPIValorUnitario

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 184599

    Saida_Celula_IPIValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_IPIValorUnitario:

    Saida_Celula_IPIValorUnitario = gErr

    Select Case gErr

        Case 184599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184600)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Sub IPIAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub IPIAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub IPIAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = IPIAliquota
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IPIValorUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPIValorUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub IPIValorUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub IPIValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = IPIValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSAliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSAliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ICMSAliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ICMSAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ICMSAliquota
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ICMSPercRedBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSPercRedBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub ICMSPercRedBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub ICMSPercRedBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = ICMSPercRedBase
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DespImpValorRateado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DespImpValorRateado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItensNF)
End Sub

Private Sub DespImpValorRateado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensNF)
End Sub

Private Sub DespImpValorRateado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensNF.objControle = DespImpValorRateado
    lErro = Grid_Campo_Libera_Foco(objGridItensNF)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Recalcular()

    'percorre o grid de itens e compara as qtdes com as di
        'se a qtde está diferente entao
            'a di nao está toda na nf
            'pegar o valor aduaneiro do item como um todo, o valor aduaneiro da di como um todo e proporcionalizar as despesas da di
            'recalcular o ii, pis e cofins para este item e colocar este valor
            
    'se toda a di está na nf, entao
        Call BotaoTrazerDados_Click
        Exit Sub
   
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155763)

    End Select

    Exit Sub

End Sub
