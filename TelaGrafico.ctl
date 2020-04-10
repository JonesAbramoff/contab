VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TelaGrafico 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ControlContainer=   -1  'True
   DrawMode        =   1  'Blackness
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Seta 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   5760
      Picture         =   "TelaGrafico.ctx":0000
      ScaleHeight     =   67.5
      ScaleMode       =   0  'User
      ScaleWidth      =   28.846
      TabIndex        =   32
      Top             =   4980
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Line3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   6090
      ScaleHeight     =   0
      ScaleWidth      =   165
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox Line2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   5565
      ScaleHeight     =   0
      ScaleWidth      =   165
      TabIndex        =   30
      Top             =   5010
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox Line1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   0
      Left            =   5340
      ScaleHeight     =   165
      ScaleWidth      =   0
      TabIndex        =   29
      Top             =   4935
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Frame Icone 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   23
      Top             =   10
      Width           =   180
      Begin VB.Image ImageIniFim 
         Height          =   180
         Index           =   0
         Left            =   0
         Picture         =   "TelaGrafico.ctx":048A
         Top             =   -15
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image ImageFim 
         Height          =   180
         Index           =   0
         Left            =   0
         Picture         =   "TelaGrafico.ctx":095C
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image ImageIni 
         Height          =   180
         Index           =   0
         Left            =   0
         Picture         =   "TelaGrafico.ctx":0E2E
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.PictureBox PictureDin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   2490
      Picture         =   "TelaGrafico.ctx":1300
      ScaleHeight     =   67.5
      ScaleMode       =   0  'User
      ScaleWidth      =   40.385
      TabIndex        =   28
      Top             =   4950
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox PictureDin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   1995
      Picture         =   "TelaGrafico.ctx":178A
      ScaleHeight     =   67.5
      ScaleMode       =   0  'User
      ScaleWidth      =   40.385
      TabIndex        =   27
      Top             =   4905
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox PictureDin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   1725
      Picture         =   "TelaGrafico.ctx":1C14
      ScaleHeight     =   67.5
      ScaleMode       =   0  'User
      ScaleWidth      =   52.5
      TabIndex        =   26
      Top             =   4860
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.ComboBox ZOOM 
      Height          =   315
      ItemData        =   "TelaGrafico.ctx":209E
      Left            =   5820
      List            =   "TelaGrafico.ctx":20AB
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   165
      Width           =   645
   End
   Begin VB.TextBox Item 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   8220
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   270
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox TamDia 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   8
      Left            =   8220
      TabIndex        =   15
      Top             =   5250
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   2
      Left            =   1264
      TabIndex        =   9
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   2423
      TabIndex        =   10
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   4
      Left            =   3582
      TabIndex        =   11
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   5
      Left            =   4741
      TabIndex        =   12
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   6
      Left            =   5900
      TabIndex        =   13
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Botao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   7
      Left            =   7059
      TabIndex        =   14
      Top             =   5265
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox TextoExibicao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1575
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton BotaoGerar 
      Caption         =   "Gerar Gráfico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6705
      TabIndex        =   4
      Top             =   150
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8085
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   1155
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   105
         Picture         =   "TelaGrafico.ctx":20BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir direto"
         Top             =   60
         Width           =   405
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "TelaGrafico.ctx":21BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   405
      End
   End
   Begin MSMask.MaskEdBox NumDias 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   165
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   4950
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   195
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   3675
      TabIndex        =   1
      Top             =   165
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridAux 
      Height          =   705
      Left            =   75
      TabIndex        =   20
      Top             =   4095
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1244
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   4155
      Left            =   15
      TabIndex        =   7
      Top             =   690
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.Label Label4 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   25
      Top             =   225
      Width           =   570
   End
   Begin VB.Label Label3 
      Caption         =   "Zoom:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      TabIndex        =   24
      Top             =   210
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Data de Início:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2370
      TabIndex        =   18
      Top             =   225
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Núm. dias Gráfico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   17
      Top             =   210
      Width           =   2415
   End
End
Attribute VB_Name = "TelaGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjTelaGrafico As ClassTelaGrafico
Public iTamanhoItemLargura As Integer
Public iTamanhoItemAltura As Integer
Dim iNumLinhasExibidas As Integer
Dim iNumMinColunasVisiveis As Integer
Dim iNumControles As Integer
Dim iToleranciaErros As Integer
Dim dZOOM As Double
Dim iTamanhoIcone As Integer
Dim iTamanhoFonte As Integer

Const DEFAULT_TAMANHO_ITEM_LARGURA = 540
Const DEFAULT_TAMANHO_ITEM_ALTURA = 360
Const DEFAULT_NUMERO_LINHAS_EXIBIDAS = 10
Const DEFAULT_TAMANHO_ICONE = 180
Const DEFAULT_TAMANHO_FONTE = 7

Const MINIMO_TAMANHO_ITEM_ALTURA = 240
Const ESPESSURA_LINE = 40
Const COMPRIMENTO_LINE3 = 140

Const POSICAO_GRID_TOP = 660
Const POSICAO_GRID_LEFT = 105

Const NUMERO_MAX_CONTROLES = 500
Const NUM_MAXIMO_BOTOES = 8
Const NUM_MAXIMO_LINHAS = 200

Const TAMANHO_MINIMO_WIDTH = 75
Const TAMANHO_MINIMO_HEIGHT = 165
Const TAMANHO_MINIMO_GRID = 8100
Const TAMANHO_SCROLL = 340

Const AJUSTE_LARGURA_INICIAL = 45
Const AJUSTE_ALTURA_INICIAL = 60
Const AJUSTE_ALTURA = 15
Const AJUSTE_LARGURA = 15
Const AJUSTE_GRIDAUX_TOP = 40

Const PORCENTAGEM_ERRO = 0.2

Const INDICE_IMPRESSAO_ITEM = 1000

Dim bDesabilitaCmdGridAux As Boolean

'Grid de Itens
Dim objGridItens As AdmGrid
Dim objGridAux As AdmGrid

Private Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    gobjTelaGrafico.iNumDias = StrParaInt(NumDias.Text)
    gobjTelaGrafico.dtDataInicio = StrParaDate(Data.Text)
    gobjTelaGrafico.iZOOM = ZOOM.ItemData(ZOOM.ListIndex)

    lErro = Trata_Parametros(gobjTelaGrafico)
    If lErro <> SUCESSO Then gError 138230
    
    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr

        Case 138230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174604)

    End Select

    Exit Sub

End Sub

Private Sub Default()

Dim lErro As Long
Dim iIndice As Integer

    GridItens.Top = POSICAO_GRID_TOP
    GridItens.Left = POSICAO_GRID_LEFT
    
    GridItens.Font.Name = "Arial"
    GridItens.Font.Size = iTamanhoFonte + 1
    
    GridItens.CellFontName = "Arial"
    GridItens.CellFontSize = iTamanhoFonte + 1
    
    TamDia.Width = iTamanhoItemLargura
    TamDia.Height = iTamanhoItemAltura
    
    For iIndice = 1 To iNumControles
        Item(iIndice).Height = iTamanhoItemAltura
        Icone(iIndice).Height = iTamanhoIcone
        Icone(iIndice).Width = iTamanhoIcone
        ImageIni(iIndice).Stretch = True
        ImageIni(iIndice).Height = iTamanhoIcone
        ImageIni(iIndice).Width = iTamanhoIcone
        ImageFim(iIndice).Stretch = True
        ImageFim(iIndice).Height = iTamanhoIcone
        ImageFim(iIndice).Width = iTamanhoIcone
        ImageIniFim(iIndice).Stretch = True
        ImageIniFim(iIndice).Height = iTamanhoIcone
        ImageIniFim(iIndice).Width = iTamanhoIcone
        Item(iIndice).FontSize = iTamanhoFonte
    Next
    
End Sub

Private Sub Botao_Click(Index As Integer)

Dim lErro As Long
Dim objTelaGraficoBotao As ClassTelaGraficoBotao

On Error GoTo Erro_Botao_Click

    Set objTelaGraficoBotao = gobjTelaGrafico.colBotoes.Item(Index)
    
    Select Case objTelaGraficoBotao.colParametros.Count
    
        Case 0
            lErro = CallByName(gobjTelaGrafico.objTela, objTelaGraficoBotao.sNomeFuncao, VbMethod)
        
        Case 1
            lErro = CallByName(gobjTelaGrafico.objTela, objTelaGraficoBotao.sNomeFuncao, VbMethod, objTelaGraficoBotao.colParametros.Item(1))
        
        Case 2
            lErro = CallByName(gobjTelaGrafico.objTela, objTelaGraficoBotao.sNomeFuncao, VbMethod, objTelaGraficoBotao.colParametros.Item(1), objTelaGraficoBotao.colParametros.Item(2))
        
        Case 3
            lErro = CallByName(gobjTelaGrafico.objTela, objTelaGraficoBotao.sNomeFuncao, VbMethod, objTelaGraficoBotao.colParametros.Item(1), objTelaGraficoBotao.colParametros.Item(2), objTelaGraficoBotao.colParametros.Item(3))
        
        Case 4
            lErro = CallByName(gobjTelaGrafico.objTela, objTelaGraficoBotao.sNomeFuncao, VbMethod, objTelaGraficoBotao.colParametros.Item(1), objTelaGraficoBotao.colParametros.Item(2), objTelaGraficoBotao.colParametros.Item(3), objTelaGraficoBotao.colParametros.Item(4))
        
        Case Else
    
    End Select
    
    If lErro <> SUCESSO Then gError 138236
    
    If objTelaGraficoBotao.iAtualizaRetornoClick = MARCADO Then

        lErro = Trata_Parametros(gobjTelaGrafico)
        If lErro <> SUCESSO Then gError 138248

    End If
    
    Exit Sub

Erro_Botao_Click:

    Select Case gErr
    
        Case 138236, 138248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174605)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174606)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Gráfico"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "Gráfico"
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

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    bDesabilitaCmdGridAux = False
               
    Set gobjTelaGrafico = New ClassTelaGrafico
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174607)

    End Select

    Exit Sub

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing
    Set objGridAux = Nothing

    Set gobjTelaGrafico = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174608)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTelaGrafico As ClassTelaGrafico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    bDesabilitaCmdGridAux = False

    If Not (objTelaGrafico Is Nothing) Then
    
        Set gobjTelaGrafico = objTelaGrafico
        
        lErro = Traz_Default_Tela(objTelaGrafico)
        If lErro <> SUCESSO Then gError 138231
        
        lErro = Traz_Grafico_Tela(objTelaGrafico)
        If lErro <> SUCESSO Then gError 138232
    
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 138231, 138232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174609)

    End Select
    
    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
'    MsgBox "Top: " & GridItens.CellTop & ", Left: " & GridItens.CellLeft & ", Height: " & GridItens.CellHeight & ", Width: " & GridItens.CellWidth

    
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhaAnterior As Integer
Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridItens_KeyDown

    'guarda as linhas do grid antes de apagar
    iLinhaAnterior = GridItens.Row
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
       
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174610)
    
    End Select

    Exit Sub
        
End Sub

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    If GridAux.Visible Then
        GridAux.LeftCol = GridItens.LeftCol
        
        'No caso da última linha o scroll interfere por isso não pode passar de
        'onde o grid auxiliar vai
        GridItens.LeftCol = GridAux.LeftCol
    End If

    Call Grid_Scroll(objGridItens)

    Call Traz_Grafico_Tela(gobjTelaGrafico)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case Else
                    'Não há tratamento específico na saída de célula,
                    'uma vez que os campos não serão editados
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 138233

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 138233
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174611)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    
    For iIndice = 1 To StrParaInt(NumDias.Text)
        objGrid.colColuna.Add (CStr(iIndice))
    Next

    'Controles que participam do Grid
    For iIndice = 1 To StrParaInt(NumDias.Text)
        objGrid.colCampo.Add (TamDia.Name)
    Next

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_LINHAS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    objGrid.iLinhasVisiveis = iNumLinhasExibidas

    'Largura da primeira coluna
    GridItens.ColWidth(0) = iTamanhoItemLargura

    GridItens.RowHeight(0) = iTamanhoItemAltura
        
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
                
    objControl.Enabled = False
                
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174612)

    End Select

    Exit Sub

End Sub

Private Sub TamDia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub TamDia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub TamDia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TamDia
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Traz_Grafico_Tela(ByVal objTelaGrafico As ClassTelaGrafico, Optional ByVal bImpressao As Boolean = False) As Long

Dim lErro As Long
Dim objTelaGraficoItem As ClassTelaGraficoItens
Dim lIndice As Long
Dim lIndice2 As Long
Dim bComecou As Boolean
Dim iDiaDif As Integer
Dim iDiaDifAux As Integer
Dim lTop As Long
Dim lLeft As Long
Dim lLeftAux As Long
Dim lAux As Long
Dim lTamanho As Long
Dim iDiasAux As Integer
Dim objIcone As Object
Dim objTelaGraficoImp As New ClassTelaGraficoImpressao
Dim objTelaGraficoImpItens As ClassTelaGraficoImpItens
Dim objPicture As Object
Dim lLeftIcone As Long
Dim lTopIcone As Long
Dim lNumLinhas As Long
Dim lNumLinhasAux As Long

On Error GoTo Erro_Traz_Grafico_Tela

    If Not bImpressao Then

        Call Remonta_Grid
        Call Limpa_Grafico
        
    Else
    
        objTelaGraficoImp.sNome = objTelaGrafico.sNomeTela
        objTelaGraficoImp.sTexto = objTelaGrafico.sTextoImpressao
        objTelaGraficoImp.sTexto2 = "Data de início: " & Format(Data.Text, "dd/mm/yyyy")
        objTelaGraficoImp.sNomeArqFigura = objTelaGrafico.sNomeArqFigura
        
        Set objTelaGraficoImp.colItens = New Collection
        
    End If

    lIndice = 0

    'Para cada item
    For Each objTelaGraficoItem In objTelaGrafico.colItens
            
        'Acha a quantidade de dias => repercute no tamanho do controle
        iDiaDif = DateDiff("d", StrParaDate(Data.Text), objTelaGraficoItem.dtDataInicio) + 1
        
        lIndice = lIndice + 1
        
        If iDiaDif <= StrParaInt(NumDias.Text) And iDiaDif + objTelaGraficoItem.iQtdDias > 0 Then
               
            'Obtém indicadores de posição e tamanho
            lTop = POSICAO_GRID_TOP + GridItens.RowPos(1) + AJUSTE_ALTURA_INICIAL
            
            If iDiaDif >= 0 Then
                lLeft = POSICAO_GRID_LEFT + GridItens.ColPos(iDiaDif) + AJUSTE_LARGURA_INICIAL
            Else
                lLeft = POSICAO_GRID_LEFT + (iDiaDif * ((iTamanhoItemLargura + AJUSTE_LARGURA))) + AJUSTE_LARGURA_INICIAL
            End If
            
            lTamanho = (CLng(iTamanhoItemLargura) + AJUSTE_LARGURA) * objTelaGraficoItem.iQtdDias
            
            lIndice2 = 0
            
            'Procura para ver se já exite controle na posição,
            'se existir leva esse controle para baixo até não mais existir
            'conflito
            lNumLinhasAux = 0
            For lIndice2 = 1 To iNumControles
            
                If lIndice2 = lIndice Then
                    Exit For
                End If
                
                If objTelaGrafico.colItens.Item(lIndice2).dtDataInicio <> DATA_NULA Then
                
                    'Acha a quantidade de dias => repercute no tamanho do controle
                    iDiaDifAux = DateDiff("d", StrParaDate(Data.Text), objTelaGrafico.colItens.Item(lIndice2).dtDataInicio) + 1
                    iDiasAux = objTelaGrafico.colItens.Item(lIndice2).iQtdDias
                    
                    'Se o ínício ou fim do controle atual está entre o início de algum
                    'outro controle anterior, tem que colocar o controle mais para baixo.
                    If (iDiaDif >= iDiaDifAux And iDiaDif < iDiaDifAux + iDiasAux) Or _
                    (iDiaDif + objTelaGraficoItem.iQtdDias > iDiaDifAux And iDiaDif + objTelaGraficoItem.iQtdDias <= iDiaDifAux + iDiasAux) Or _
                    (iDiaDif <= iDiaDifAux And iDiaDif + objTelaGraficoItem.iQtdDias >= iDiaDifAux + iDiasAux) Then
                        If Item(lIndice2).Top >= lTop Then
                            lTop = Item(lIndice2).Top + iTamanhoItemAltura + AJUSTE_ALTURA
                            lNumLinhasAux = lNumLinhasAux + 1
                        End If
                    End If
                    
                End If
            Next
            
            If objTelaGrafico.iCadaEtapaUmaLinha = MARCADO Then

                 'Obtém indicadores de posição e tamanho
                lTop = POSICAO_GRID_TOP + GridItens.RowPos(lIndice) + AJUSTE_ALTURA_INICIAL
                lNumLinhasAux = lIndice
            End If
            
            If lNumLinhas < lNumLinhasAux Then lNumLinhas = lNumLinhasAux
                 
            'Se parte do controle passou do início do grid
            If (lLeft <= GridItens.Left + iTamanhoItemLargura + AJUSTE_LARGURA - iToleranciaErros) And Not bImpressao Then
            
                'Se ainda tem controle para ser exibido, ou seja, se a parte final dele estiver após
                'o início do grid, então acerta o tamanho e a posição, senão tira da tela.
                If lLeft + lTamanho >= GridItens.ColPos(0) + iToleranciaErros + iTamanhoItemLargura Then
                    lLeftAux = lLeft
                    lLeft = POSICAO_GRID_LEFT + iTamanhoItemLargura + AJUSTE_LARGURA_INICIAL
                    lAux = Round((lLeft - lLeftAux) / (iTamanhoItemLargura + AJUSTE_LARGURA))
                    lTamanho = (iTamanhoItemLargura + AJUSTE_LARGURA) * (objTelaGraficoItem.iQtdDias - lAux)
                Else
                    lLeft = POSICAO_FORA_TELA
                End If
            
            End If
            
            lIndice2 = 0
            bComecou = False
            
            'Obtém o último controle
            For lIndice2 = 1 To StrParaInt(NumDias.Text)
                If (GridItens.ColPos(lIndice2) + iToleranciaErros) > (GridItens.Left + GridItens.Width) Then
                    If bComecou Then
                        Exit For
                    End If
                Else
                    bComecou = True
                End If
            Next
            lIndice2 = lIndice2 - 1
                       
            'Se parte do controle passou do final do grid
            'If lLeft + lTamanho >= GridItens.ColPos(lIndice2) + iToleranciaErros Then
            If (lLeft + lTamanho >= GridItens.Left + GridItens.Width - TAMANHO_SCROLL - iToleranciaErros) And Not bImpressao Then
            
                'Se ainda tem controle para ser exibido, ou seja, se a parte inicial estiver antes do
                'fim do grid, então acerta o tamanho e a posição, senão tira da tela.
                'If lLeft <= GridItens.ColPos(lIndice2) + iToleranciaErros Then
                If lLeft <= GridItens.Left + GridItens.Width - TAMANHO_SCROLL - iToleranciaErros Then
                    'Se não existe não espaço entre o final do grid e a última coluna
                    If (GridItens.ColPos(lIndice2) - iToleranciaErros + iTamanhoItemLargura) < (GridItens.Left + GridItens.Width - TAMANHO_SCROLL) Then
                        lTamanho = GridItens.ColPos(lIndice2) + iTamanhoItemLargura + AJUSTE_LARGURA_INICIAL - lLeft
                    Else
                        lTamanho = GridItens.Left + GridItens.Width - lLeft - TAMANHO_SCROLL
                    End If
                Else
                    lLeft = POSICAO_FORA_TELA
                End If
            
            End If
                        
            If (lTamanho > GridItens.Width) And Not bImpressao Then
                lTamanho = GridItens.Width
            End If
            
            If lTamanho > 0 Then
            
                If Not bImpressao Then
            
                    Item(lIndice).MousePointer = vbArrow
                    Item(lIndice).TabStop = False
                    Item(lIndice).BackColor = objTelaGraficoItem.lCor
                    
                    'Incluido por Jorge Specian
                    '---------------------------------------
                    Item(lIndice).ForeColor = LetraCor(objTelaGraficoItem.lCor)
                    '---------------------------------------
                    
                    Item(lIndice).Text = objTelaGraficoItem.sNome
                    Item(lIndice).Width = lTamanho
                    Item(lIndice).Top = lTop
                    Item(lIndice).Left = lLeft
                    
                Else
                
                    Set objTelaGraficoImpItens = New ClassTelaGraficoImpItens
                    
                    objTelaGraficoImpItens.sText = objTelaGraficoItem.sNome
                    objTelaGraficoImpItens.lWidth = lTamanho
                    objTelaGraficoImpItens.lTop = lTop
                    objTelaGraficoImpItens.lLeft = lLeft
                    objTelaGraficoImpItens.lForeColor = LetraCor(objTelaGraficoItem.lCor)
                    objTelaGraficoImpItens.lBackColor = objTelaGraficoItem.lCor
                    objTelaGraficoImpItens.lHeight = Item(lIndice).Height
                    objTelaGraficoImpItens.iBorderStyle = Item(lIndice).BorderStyle
                    objTelaGraficoImpItens.lFontSize = Item(lIndice).FontSize
                    objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT
                    objTelaGraficoImpItens.sFontName = Item(lIndice).FontName
                    objTelaGraficoImpItens.iLegenda = MARCADO
                    objTelaGraficoImpItens.sDescricao = objTelaGraficoItem.sDescricao
                    
                    objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
                
                End If
                
                'Se não passou do final em altura
                If (Item(lIndice).Top + Item(lIndice).Height < GridItens.Top + GridItens.Height + iToleranciaErros - TAMANHO_SCROLL) Or bImpressao Then
                    
                    'Se não passou do início em altura
                    If (Item(lIndice).Top > GridItens.Top - iToleranciaErros + iTamanhoItemAltura) Or bImpressao Then
                    
                        If Not bImpressao Then
                            Item(lIndice).Visible = True
                        End If
                        
                        'Se tem ìcone coloca na tela
                        If objTelaGraficoItem.iIcone <> 0 Then
                        
                            Select Case objTelaGraficoItem.iIcone
                            
                                Case TELA_GRAFICO_ICONE_INICIO
                                
                                    If Not bImpressao Then
                                        ImageIni(lIndice).Visible = True
                                        Icone(lIndice).Left = Item(lIndice).Left
                                    Else
                                        Set objPicture = ImageIni(lIndice)
                                        lLeftIcone = Item(lIndice).Left
                                    End If
                                    
                                Case TELA_GRAFICO_ICONE_INICIO_E_FIM
                                    If Not bImpressao Then
                                        ImageIniFim(lIndice).Visible = True
                                        Icone(lIndice).Left = Item(lIndice).Left + (Item(lIndice).Width - Icone(lIndice).Width) / 2
                                    Else
                                        Set objPicture = ImageIniFim(lIndice)
                                        lLeftIcone = Item(lIndice).Left + (Item(lIndice).Width - Icone(lIndice).Width) / 2
                                    End If
                                Case TELA_GRAFICO_ICONE_FIM
                                    If Not bImpressao Then
                                        ImageFim(lIndice).Visible = True
                                        Icone(lIndice).Left = Item(lIndice).Left + Item(lIndice).Width - Icone(lIndice).Width
                                    Else
                                        Set objPicture = ImageFim(lIndice)
                                        lLeftIcone = Item(lIndice).Left + Item(lIndice).Width - Icone(lIndice).Width
                                    End If
                                Case Else
                            
                            End Select
                            
                            If (Icone(lIndice).Width <= Item(lIndice).Width) Then
                                If Not bImpressao Then
                                    Icone(lIndice).Top = Item(lIndice).Top + Item(lIndice).Height - Icone(lIndice).Height
                                    Icone(lIndice).Visible = True
                                Else
                                    lTopIcone = Item(lIndice).Top + Item(lIndice).Height - Icone(lIndice).Height
                                End If
                            End If
                        
                            If Not bImpressao Then

                                Set objTelaGraficoImpItens = New ClassTelaGraficoImpItens
                                     
                                 objTelaGraficoImpItens.lWidth = Icone(lIndice).Width
                                 objTelaGraficoImpItens.lTop = lTopIcone
                                 objTelaGraficoImpItens.lLeft = lLeftIcone
                                 objTelaGraficoImpItens.lBackColor = Icone(lIndice).BackColor
                                 objTelaGraficoImpItens.lHeight = Icone(lIndice).Height
                                 objTelaGraficoImpItens.iBorderStyle = Icone(lIndice).BorderStyle
                                 objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_LINE
                                 'Set objTelaGraficoImpItens.objPicture = objPicture
                                 
                                 objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
                                 
                            End If
                        
                        End If
                         
                    End If
                
                End If
            
            End If
            
            If Not bImpressao Then
                Set objTelaGraficoItem.objControle = Item(lIndice)
            Else
                'Inclui na tela um novo Controle para essa coluna
                Load Item(lIndice + INDICE_IMPRESSAO_ITEM)
                'Traz o controle recem desenhado para a frente
                Item(lIndice + INDICE_IMPRESSAO_ITEM).ZOrder
                'Torna o controle visível
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Visible = False
                
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Text = objTelaGraficoItem.sNome
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Width = lTamanho
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Top = lTop
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Left = lLeft
                Item(lIndice + INDICE_IMPRESSAO_ITEM).Height = Item(lIndice).Height

                Set objTelaGraficoItem.objControle = Item(lIndice + INDICE_IMPRESSAO_ITEM)

            End If
                    
        End If
    
    Next
    
    lErro = Traz_Relacoes_Tela(objTelaGrafico, objTelaGraficoImp, bImpressao)
    If lErro <> SUCESSO Then gError 185738
    
    If bImpressao Then
    
        For lIndice = INDICE_IMPRESSAO_ITEM To Item.UBound
            Call Unload_Object(Item, lIndice)
        Next
    
        For lIndice = 1 To GridItens.Cols - 1
        
            Set objTelaGraficoImpItens = New ClassTelaGraficoImpItens
                 
            objTelaGraficoImpItens.lWidth = GridItens.ColWidth(lIndice)
            objTelaGraficoImpItens.lTop = GridItens.Top
            objTelaGraficoImpItens.lLeft = POSICAO_GRID_LEFT + GridItens.ColPos(lIndice)
            objTelaGraficoImpItens.lBackColor = GridItens.BackColorFixed
            objTelaGraficoImpItens.lHeight = GridItens.RowHeight(1)
            objTelaGraficoImpItens.iBorderStyle = 1
            objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_LINHA
            objTelaGraficoImpItens.sText = GridItens.TextMatrix(0, lIndice)
            objTelaGraficoImpItens.sFontName = GridItens.Font.Name
            objTelaGraficoImpItens.lForeColor = GridItens.ForeColor
            objTelaGraficoImpItens.lFontSize = GridItens.Font.Size
             
            objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
        
        Next
        
        For lIndice = 1 To lNumLinhasAux
        
            Set objTelaGraficoImpItens = New ClassTelaGraficoImpItens
                 
            objTelaGraficoImpItens.lWidth = GridItens.ColWidth(lIndice)
            objTelaGraficoImpItens.lTop = GridItens.Top + GridItens.RowPos(lIndice)
            objTelaGraficoImpItens.lLeft = GridItens.Left
            objTelaGraficoImpItens.lBackColor = GridItens.BackColorFixed
            objTelaGraficoImpItens.lHeight = GridItens.RowHeight(1)
            objTelaGraficoImpItens.iBorderStyle = 1
            objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_COLUNA
            objTelaGraficoImpItens.sText = GridItens.TextMatrix(lIndice, 0)
            objTelaGraficoImpItens.sFontName = GridItens.Font.Name
            objTelaGraficoImpItens.lForeColor = GridItens.ForeColor
            objTelaGraficoImpItens.lFontSize = GridItens.Font.Size
             
            objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
        
        Next
    
        Call Imprimir.Imprimir_Layout(objTelaGraficoImp)
        
        objTelaGrafico.iNumFiguras = objTelaGraficoImp.iNumFiguras
    
    End If
     
    objGridItens.iLinhasExistentes = NUM_MAXIMO_LINHAS

    Call Trata_Grid_Aux(TELAGRAFICO_FUNCAO_PREENCHE_GRIDAUX)
    
    

    Traz_Grafico_Tela = SUCESSO

    Exit Function

Erro_Traz_Grafico_Tela:

    Traz_Grafico_Tela = gErr

    Select Case gErr
    
        Case 185738
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174613)

    End Select

    Exit Function

End Function

Private Function Traz_Default_Tela(ByVal objTelaGrafico As ClassTelaGrafico) As Long

Dim lErro As Long
Dim objTelaGraficoItem As ClassTelaGraficoItens
Dim objTelaGraficoBotao As ClassTelaGraficoBotao
Dim dtDataMin As Date
Dim dtDataMax As Date
Dim iIndice As Integer
Dim lNum As Long
Dim iNumLinhasPadrao As Integer

On Error GoTo Erro_Traz_Default_Tela
        
    If objTelaGrafico.iZOOM <> 0 Then
    
        Call Combo_Seleciona_ItemData(ZOOM, objTelaGrafico.iZOOM)
    
        Select Case objTelaGrafico.iZOOM
        
            Case ZOOM_100
                dZOOM = ZOOM_100_PORCENT
            Case ZOOM_50
                dZOOM = ZOOM_50_PORCENT
            Case ZOOM_25
                dZOOM = ZOOM_25_PORCENT
        
        End Select
        
    Else
    
        Call Combo_Seleciona_ItemData(ZOOM, ZOOM_100)

        dZOOM = ZOOM_100_PORCENT
        
    End If
        
    'Se foi passado o tamanho que o dia deve ter colocana tela
    'senão usa o default da tela
    If objTelaGrafico.iTamanhoDia <> 0 Then
        iTamanhoItemLargura = objTelaGrafico.iTamanhoDia / dZOOM
    Else
        iTamanhoItemLargura = DEFAULT_TAMANHO_ITEM_LARGURA / dZOOM
    End If
    
    If objTelaGrafico.iAlturaDia <> 0 Then
        iTamanhoItemAltura = objTelaGrafico.iAlturaDia / dZOOM
    Else
        iTamanhoItemAltura = DEFAULT_TAMANHO_ITEM_ALTURA / dZOOM
    End If
    
    If iTamanhoItemAltura < MINIMO_TAMANHO_ITEM_ALTURA Then
        iTamanhoItemAltura = MINIMO_TAMANHO_ITEM_ALTURA
    End If
    
    If objTelaGrafico.iAlteraQtdLinhasPara <> 0 Then
        iNumLinhasPadrao = objTelaGrafico.iAlteraQtdLinhasPara
    Else
        iNumLinhasPadrao = DEFAULT_NUMERO_LINHAS_EXIBIDAS
    End If
    
    If iTamanhoItemAltura < TAMANHO_MINIMO_HEIGHT Then iTamanhoItemAltura = TAMANHO_MINIMO_HEIGHT
    If iTamanhoItemLargura < TAMANHO_MINIMO_WIDTH Then iTamanhoItemLargura = TAMANHO_MINIMO_WIDTH
    
    iNumMinColunasVisiveis = Fix(TAMANHO_MINIMO_GRID / iTamanhoItemLargura) + 1
    iNumLinhasExibidas = Fix(iNumLinhasPadrao * (DEFAULT_TAMANHO_ITEM_ALTURA / iTamanhoItemAltura))
    
    iTamanhoIcone = DEFAULT_TAMANHO_ICONE / dZOOM
    iTamanhoFonte = Fix(1 + DEFAULT_TAMANHO_FONTE - dZOOM)
    
    'Obtém as datas iniciais para começar o algoritmo
    If objTelaGrafico.colItens.Count > 0 Then
        dtDataMin = objTelaGrafico.colItens.Item(1).dtDataInicio
        dtDataMax = objTelaGrafico.colItens.Item(1).dtDataFim
    Else
        dtDataMin = gdtDataAtual
        dtDataMax = gdtDataAtual + 30
    End If
    
    iIndice = 0

    'Para cada item que vai ser colocado na tela
    For Each objTelaGraficoItem In objTelaGrafico.colItens
    
        iIndice = iIndice + 1
    
        'Obtém a menor data
        If dtDataMin > objTelaGraficoItem.dtDataInicio Then dtDataMin = objTelaGraficoItem.dtDataInicio
        
        'Obtém a maior data
        If dtDataMax < objTelaGraficoItem.dtDataFim Then dtDataMax = objTelaGraficoItem.dtDataFim
    
        'Obtém a diferença de dias entre as datas do item
        objTelaGraficoItem.iQtdDias = DateDiff("d", objTelaGraficoItem.dtDataInicio, objTelaGraficoItem.dtDataFim) + 1
                
        'Se a cor não tiver sido passada
        If objTelaGraficoItem.lCor = 0 Then
        
            'Se não foi passado um indice para agrupamento de cor
            If objTelaGraficoItem.iIndiceCor = 0 Then
                objTelaGraficoItem.lCor = Cor(iIndice)
            Else
                objTelaGraficoItem.lCor = Cor(objTelaGraficoItem.iIndiceCor)
            End If
        End If
    
    Next
    
    Data.PromptInclude = False
    If objTelaGrafico.dtDataInicio = DATA_NULA Then
        Data.Text = Format(dtDataMin, "dd/mm/yy")
    Else
        Data.Text = Format(objTelaGrafico.dtDataInicio, "dd/mm/yy")
    End If
    Data.PromptInclude = True
    
    If objTelaGrafico.sNomeArqFigura = "" Then
        Caption = objTelaGrafico.sNomeTela
    End If
    
    iIndice = 0
    
    'Deixa os botões invisíveis
    For iIndice = 1 To NUM_MAXIMO_BOTOES
        Botao(iIndice).Visible = False
    Next
    
    iIndice = 0
    
    'Para cada botão passado coloca na tela
    For Each objTelaGraficoBotao In objTelaGrafico.colBotoes
    
        iIndice = iIndice + 1
        
        If iIndice > NUM_MAXIMO_BOTOES Then gError 138235
        
        Botao(iIndice).Caption = objTelaGraficoBotao.sNome
        Botao(iIndice).ToolTipText = objTelaGraficoBotao.sTextoExibicao
        
        Botao(iIndice).Visible = True
    
    Next
    
    'Obtém a maior diferença de datas
    If objTelaGrafico.iNumDias <> 0 Then
        lNum = objTelaGrafico.iNumDias
    Else
        lNum = DateDiff("d", dtDataMin, dtDataMax) + 1
    End If
    
    'Coloca um número mínibo de dias a serem exibidos
    If lNum < iNumMinColunasVisiveis Then lNum = iNumMinColunasVisiveis
    
    iNumControles = objTelaGrafico.colItens.Count
    
    If iNumControles > NUMERO_MAX_CONTROLES Then gError 138248
    
    For iIndice = 1 To iNumControles
    
        If Item.UBound < iIndice Then
            'Inclui na tela um novo Controle para essa coluna
            Load Item(iIndice)
            'Traz o controle recem desenhado para a frente
            Item(iIndice).ZOrder
            'Torna o controle visível
            Item(iIndice).Visible = True
        End If
        
        If Icone.UBound < iIndice Then
            'Inclui na tela um novo Controle para essa coluna
            Load Icone(iIndice)
            'Traz o controle recem desenhado para a frente
            Icone(iIndice).ZOrder
            'Torna o controle visível
            Icone(iIndice).Visible = True
         
        End If
        
        If ImageIniFim.UBound < iIndice Then
            'Inclui na tela um novo Controle para essa coluna
            Load ImageIniFim(iIndice)
            Set ImageIniFim(iIndice).Container = Icone(iIndice)
            'Traz o controle recem desenhado para a frente
            ImageIniFim(iIndice).ZOrder
            'Torna o controle visível
            ImageIniFim(iIndice).Visible = True
            
        End If

        If ImageIni.UBound < iIndice Then
            'Inclui na tela um novo Controle para essa coluna
            Load ImageIni(iIndice)
            Set ImageIni(iIndice).Container = Icone(iIndice)
            'Traz o controle recem desenhado para a frente
            ImageIni(iIndice).ZOrder
            'Torna o controle visível
            ImageIni(iIndice).Visible = True
        End If

        If ImageFim.UBound < iIndice Then
            'Inclui na tela um novo Controle para essa coluna
            Load ImageFim(iIndice)
            Set ImageFim(iIndice).Container = Icone(iIndice)
            'Traz o controle recem desenhado para a frente
            ImageFim(iIndice).ZOrder
            'Torna o controle visível
            ImageFim(iIndice).Visible = True
        End If
    
    Next
        
    iToleranciaErros = iTamanhoItemLargura * PORCENTAGEM_ERRO
        
    NumDias.Text = lNum
    
    Call Default
    
    Traz_Default_Tela = SUCESSO

    Exit Function

Erro_Traz_Default_Tela:

    Traz_Default_Tela = gErr

    Select Case gErr
    
        Case 138235
            Call Rotina_Erro(vbOKOnly, "ERRO_TELAGRAFICO_LIMITE_BOTOES", gErr, NUM_MAXIMO_BOTOES)
    
        Case 138248
            Call Rotina_Erro(vbOKOnly, "ERRO_TELAGRAFICO_LIMITE_CONTROLES", gErr, NUMERO_MAX_CONTROLES)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174614)

    End Select

    Exit Function

End Function

Private Function Cor(ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim iValor As Integer

On Error GoTo Erro_Cor

    iValor = iIndice Mod 20

    Select Case iValor
    
        Case 0
             Cor = vbRed
        Case 1
             Cor = vbBlue
        Case 2
             Cor = vbGreen
        Case 3
             Cor = vbYellow
        Case 4
             Cor = vbMagenta
        Case 5
             Cor = vbCyan
        Case 6
             Cor = 8421631 ' Vermelho Claro
        Case 7
             Cor = 12632064 ' Petróleo Claro
        Case 8
             Cor = 12648447 ' Amarelo Claro
        Case 9
             Cor = 14737632 ' Cinza Claro
        Case 10
             Cor = 16761087 ' Rosa
        Case 11
             Cor = 4210752 ' Cinza Escuro
        Case 12
             Cor = 64 ' Marrom
        Case 13
             Cor = 12648384 ' Verde Claro
        Case 14
             Cor = 16448 ' Verde Musgo
        Case 15
             Cor = 4210688 ' Petróleo
        Case 16
             Cor = 16777152 ' Azul Claro
        Case 17
             Cor = 16384 'Verde Escuro
        Case 18
             Cor = 33023 ' Laranja
        Case 19
             Cor = 12640511 ' Laranja Claro
        Case Else
             Cor = vbWhite
        
    End Select

    Exit Function

Erro_Cor:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174615)

    End Select

    Exit Function

End Function

Private Sub Item_Click(Index As Integer)
'Ao clicar em um item abre a tela passando os parametros

Dim objTelaGraficoItem As ClassTelaGraficoItens
Dim lErro As Long

On Error GoTo Erro_Item_Click

'    MsgBox "Top: " & Item(Index).Top & ", Left: " & Item(Index).Left & ", Height: " & Item(Index).Height & ", Width: " & Item(Index).Width

    Set objTelaGraficoItem = gobjTelaGrafico.colItens.Item(Index)

    'Se foi passada alguma tela para o Click
    If Len(Trim(objTelaGraficoItem.sNomeTela)) <> 0 Then

        'Se não for para chamar Modal
        If gobjTelaGrafico.iModal = DESMARCADO Then

            Select Case objTelaGraficoItem.colobj.Count

                Case 0
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela)

                Case 1
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1))

                Case 2
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2))

                Case 3
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3))

                Case 4
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3), objTelaGraficoItem.colobj.Item(4))

                Case Else
                    lErro = Chama_Tela(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3), objTelaGraficoItem.colobj.Item(4), objTelaGraficoItem.colobj.Item(5))

            End Select

        Else

            Select Case objTelaGraficoItem.colobj.Count

                Case 0
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela)

                Case 1
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1))

                Case 2
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2))

                Case 3
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3))

                Case 4
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3), objTelaGraficoItem.colobj.Item(4))

                Case Else
                    lErro = Chama_Tela_Modal(objTelaGraficoItem.sNomeTela, objTelaGraficoItem.colobj.Item(1), objTelaGraficoItem.colobj.Item(2), objTelaGraficoItem.colobj.Item(3), objTelaGraficoItem.colobj.Item(4), objTelaGraficoItem.colobj.Item(5))

            End Select

        End If

        If lErro <> SUCESSO Then gError 138234

    End If

    If gobjTelaGrafico.iAtualizaRetornoClick = MARCADO Then
    
        Select Case gobjTelaGrafico.colParametros.Count
        
            Case 0
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico)
            
            Case 1
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico, gobjTelaGrafico.colParametros.Item(1))
            
            Case 2
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico, gobjTelaGrafico.colParametros.Item(1), gobjTelaGrafico.colParametros.Item(2))
            
            Case 3
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico, gobjTelaGrafico.colParametros.Item(1), gobjTelaGrafico.colParametros.Item(2), gobjTelaGrafico.colParametros.Item(3))
            
            Case 4
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico, gobjTelaGrafico.colParametros.Item(1), gobjTelaGrafico.colParametros.Item(2), gobjTelaGrafico.colParametros.Item(3), gobjTelaGrafico.colParametros.Item(4))
            
            Case Else
                lErro = CallByName(gobjTelaGrafico.objTela, gobjTelaGrafico.sNomeFuncAtualiza, VbMethod, gobjTelaGrafico, gobjTelaGrafico.colParametros.Item(1), gobjTelaGrafico.colParametros.Item(2), gobjTelaGrafico.colParametros.Item(3), gobjTelaGrafico.colParametros.Item(4), gobjTelaGrafico.colParametros.Item(5))
            
        End Select
        If lErro <> SUCESSO Then gError 138247

        lErro = Trata_Parametros(gobjTelaGrafico)
        If lErro <> SUCESSO Then gError 138248

    End If
    
    Exit Sub

Erro_Item_Click:
    
    Select Case gErr
    
        Case 138234, 138247, 138248
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174616)

    End Select

    Exit Sub

End Sub

Private Sub NumDias_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumDias_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(NumDias.Text)) <> 0 Then

        'Critica a Codigo
        lErro = Inteiro_Critica(NumDias.Text)
        If lErro <> SUCESSO Then gError 138239
        
        If StrParaInt(NumDias.Text) < iNumMinColunasVisiveis Then gError 138240

    End If

    Exit Sub

Erro_NumDias_Validate:

    Cancel = True

    Select Case gErr

        Case 138239

        Case 138240
            Call Rotina_Erro(vbOKOnly, "ERRO_TELAGRAFICO_MINIMO_DIAS", gErr, iNumMinColunasVisiveis)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174617)

    End Select

    Exit Sub
    
End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 138238

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 138238

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174618)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 138237

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 138237

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174619)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 138236

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 138236

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174620)

    End Select

    Exit Sub

End Sub

Private Sub Item_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Simula a propriedade ToolTipText sendo MultLine

    TextoExibicao.Text = gobjTelaGrafico.colItens.Item(Index).sTextoExibicao
    
    If (GridItens.Left + Item(Index).Left) + TextoExibicao.Width > UserControl.Width Then
        TextoExibicao.Left = UserControl.Width - TextoExibicao.Width
    Else
        TextoExibicao.Left = Item(Index).Left
    End If

    If (GridItens.Top + Item(Index).Top + Item(Index).Height) + TextoExibicao.Height > UserControl.Height Then
        TextoExibicao.Top = Item(Index).Top - TextoExibicao.Height
    Else
        TextoExibicao.Top = Item(Index).Top + Item(Index).Height
    End If
    
    TextoExibicao.ZOrder (0)
    
    TextoExibicao.Visible = True
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TextoExibicao.Visible = False
End Sub

Private Sub GridItens_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim iIndice As Integer
Dim iColuna As Integer

On Error GoTo Erro_GridItens_MouseMove

    TextoExibicao.Visible = False
    
    For iIndice = 1 To StrParaInt(NumDias.Text)
        If GridItens.ColPos(iIndice) <= X And GridItens.ColPos(iIndice) + iTamanhoItemLargura >= X Then
            iColuna = iIndice
            Exit For
        End If
    Next
    
    If iColuna <> 0 Then
        If IsDate(Data.Text) Then
            GridItens.ToolTipText = Format(DateAdd("d", iColuna - 1, StrParaDate(Data.Text)), "dd/mm/yyyy")
        Else
            GridItens.ToolTipText = ""
        End If
    Else
        GridItens.ToolTipText = ""
    End If

    Exit Sub

Erro_GridItens_MouseMove:

    Select Case gErr
    
        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211713)

    End Select

    Exit Sub

End Sub

Private Sub Remonta_Grid()

Dim colParametros As New Collection

    'Grid Itens
    Set objGridItens = New AdmGrid
    Set objGridAux = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    Set objGridAux.objForm = Me
    
    colParametros.Add objGridAux

    'Monta o Grid
    Call Inicializa_GridItens(objGridItens)
    
    Call Trata_Grid_Aux(TELAGRAFICO_FUNCAO_INICIALIZA_GRIDAUX, colParametros)
    
End Sub

Private Sub Limpa_Grafico()
    
Dim iIndice As Integer
    
    For iIndice = 1 To iNumControles
        Item(iIndice).Visible = False
        Item(iIndice).Left = POSICAO_FORA_TELA
        ImageIni(iIndice).Visible = False
        ImageIni(iIndice).ToolTipText = "Início"
        ImageFim(iIndice).Visible = False
        ImageFim(iIndice).ToolTipText = "Fim"
        ImageIniFim(iIndice).Visible = False
        ImageIniFim(iIndice).ToolTipText = "Início e Fim"
        Icone(iIndice).Visible = False
    Next
    
    For iIndice = 1 To Line1.UBound
        Line1(iIndice).Visible = False
        Line2(iIndice).Visible = False
        Line3(iIndice).Visible = False
        Seta(iIndice).Visible = False
    Next

End Sub

Private Function LetraCor(ByVal lCor As Long) As Long
'Incluido por Jorge Specian

Dim lErro As Long

On Error GoTo Erro_LetraCor

    Select Case lCor
             
        Case 4210752, 64, 16448, 4210688, 16384, vbBlue
            LetraCor = vbWhite
        Case Else
            LetraCor = vbBlack
        
    End Select

    Exit Function

Erro_LetraCor:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174621)

    End Select

    Exit Function

End Function

Private Function Trata_Grid_Aux(ByVal sNomeFuncao As String, Optional ByVal colParametros As Collection = Nothing) As Long
'Ao clicar em um item abre a tela passando os parametros

Dim lErro As Long
Dim objTelaGraficoParam As ClassTelaGraficoParam

On Error GoTo Erro_Trata_Grid_Aux

    If colParametros Is Nothing Then Set colParametros = New Collection
    
    For Each objTelaGraficoParam In gobjTelaGrafico.colParametrosTrataGrid
        If objTelaGraficoParam.sNomeFuncao = sNomeFuncao Then
            colParametros.Add objTelaGraficoParam.vParam
        End If
    Next

    If gobjTelaGrafico.iExibirGridAux = MARCADO Then
    
        lErro = CallByName(Me, gobjTelaGrafico.sNomeFuncTrataGrid, VbMethod, sNomeFuncao, colParametros)
        If lErro <> SUCESSO Then gError 185604

    End If
    
    Trata_Grid_Aux = SUCESSO
    
    Exit Function

Erro_Trata_Grid_Aux:

    Trata_Grid_Aux = gErr
    
    Select Case gErr
    
        Case 185604
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185605)

    End Select

    Exit Function

End Function

Private Sub GridAux_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAux, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAux, iAlterado)
    End If

End Sub

Private Sub GridAux_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridAux)
End Sub

Private Sub GridAux_GotFocus()
    Call Grid_Recebe_Foco(objGridAux)
End Sub

Private Sub GridAux_EnterCell()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_Entrada_Celula(objGridAux, iAlterado)
    End If
End Sub

Private Sub GridAux_LeaveCell()
    If Not bDesabilitaCmdGridAux Then
        Call Saida_Celula(objGridAux)
    End If
End Sub

Private Sub GridAux_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAux, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAux, iAlterado)
    End If

End Sub

Private Sub GridAux_RowColChange()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_RowColChange(objGridAux)
    End If
End Sub

Private Sub GridAux_Scroll()
    GridItens.LeftCol = GridAux.LeftCol

    Call Grid_Scroll(objGridAux)
End Sub

'Expecífica para tela de projetos
Public Function Cronograma_Trata_GridAux_PRJ(ByVal sNomeFuncao As String, ByVal colParametros As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_Cronograma_Trata_GridAux_PRJ

    Select Case sNomeFuncao

        Case TELAGRAFICO_FUNCAO_INICIALIZA_GRIDAUX
            Call Cronograma_Inicializa_GridAux_PRJ(colParametros.Item(1))
            
        Case TELAGRAFICO_FUNCAO_PREENCHE_GRIDAUX
            Call Cronograma_Preenche_GridAux_PRJ(colParametros.Item(1), colParametros.Item(2))

    End Select
    
    Cronograma_Trata_GridAux_PRJ = SUCESSO

    Exit Function

Erro_Cronograma_Trata_GridAux_PRJ:

    Cronograma_Trata_GridAux_PRJ = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185607)

    End Select

    Exit Function

End Function

'Expecífica para tela de projetos
Public Function Cronograma_Inicializa_GridAux_PRJ(ByVal objGrid As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceAux As Integer

On Error GoTo Erro_Cronograma_Inicializa_GridAux_PRJ

    GridAux.Left = GridItens.Left
    GridAux.Top = GridItens.Top + GridItens.Height - TAMANHO_SCROLL + AJUSTE_GRIDAUX_TOP
    GridAux.Visible = True

    TamDia.Font.Name = "Arial"
    TamDia.Font.Size = iTamanhoFonte

    'titulos do grid
    objGrid.colColuna.Add ("")
    
    For iIndice = 1 To StrParaInt(NumDias.Text)
        objGrid.colColuna.Add (CStr(iIndice))
    Next

    'Controles que participam do Grid
    For iIndice = 1 To StrParaInt(NumDias.Text)
        objGrid.colCampo.Add (TamDia.Name)
    Next

    objGrid.objGrid = GridAux

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 4

    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    objGrid.iLinhasVisiveis = 3
    
    'Largura da primeira coluna
    GridAux.ColWidth(0) = iTamanhoItemLargura

    GridAux.RowHeight(0) = iTamanhoItemAltura
        
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    GridAux.Font.Name = "Arial"
    GridAux.Font.Size = iTamanhoFonte

    bDesabilitaCmdGridAux = True
    For iIndice = 1 To GridAux.Cols - 1
        GridAux.Col = iIndice
        For iIndiceAux = 0 To 3
            GridAux.Row = iIndiceAux
            GridAux.CellFontName = "Arial"
            GridAux.CellFontSize = iTamanhoFonte
        Next
    Next
    bDesabilitaCmdGridAux = False
    
    Call Grid_Inicializa(objGrid)
    
    For iIndice = 1 To StrParaInt(NumDias.Text)
        GridAux.TextMatrix(0, iIndice) = Format(DateAdd("d", iIndice - 1, Data.Text), "dd/mm")
    Next
    
    'Posiciona figura de receitas\despesas\saldo
    For iIndice = 1 To 3
        GridAux.TextMatrix(iIndice, 0) = ""
        
        PictureDin.Item(iIndice).Visible = True
        
        PictureDin.Item(iIndice).Top = GridAux.Top + GridAux.RowPos(iIndice) + PictureDin.Item(iIndice).Height / 2
        PictureDin.Item(iIndice).Left = GridAux.Left + (iTamanhoItemLargura / 2) - (PictureDin.Item(iIndice).Width / 2) + 30
    Next
    
    'Zera o conteúdo em valor
    For iIndice = 0 To StrParaInt(NumDias.Text)
        GridAux.TextMatrix(1, iIndice) = ""
        GridAux.TextMatrix(2, iIndice) = ""
        GridAux.TextMatrix(3, iIndice) = ""
    Next
    
    Cronograma_Inicializa_GridAux_PRJ = SUCESSO

    Exit Function

Erro_Cronograma_Inicializa_GridAux_PRJ:

    bDesabilitaCmdGridAux = False

    Cronograma_Inicializa_GridAux_PRJ = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185606)

    End Select

    Exit Function

End Function

'Expecífica para tela de projetos
Public Function Cronograma_Preenche_GridAux_PRJ(ByVal objProjeto As ClassProjetos, ByVal objTela As Object) As Long

Dim lErro As Long
Dim colRecebPagto As New Collection
Dim objRecebPagto As ClassPRJRecebPagto
Dim objRecebPagtoRegras As ClassPRJRecebPagtoRegras
Dim dtData As Date
Dim dValorP As Double
Dim dValorR As Double
Dim dSaldoAnt As Double
Dim iIndice As Integer
'Dim objCondicaoPagto As ClassCondicaoPagto
Dim objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_Cronograma_Preenche_GridAux_PRJ

    lErro = CF("PRJRecebPagto_Le_Fluxo", objProjeto, colRecebPagto)
    If lErro <> SUCESSO Then gError 187454
    
'    For iIndice = objProjeto.colRecebPagto.Count To 1 Step -1
'        objProjeto.colRecebPagto.Remove iIndice
'    Next
'
'    For Each objRecebPagto In colRecebPagto
'
'        objProjeto.colRecebPagto.Add objRecebPagto
'
'        lErro = CF("RecebPagto_Calcula_Regras", objProjeto, objRecebPagto)
'        If lErro <> SUCESSO Then gError 187455
'
'        For Each objRecebPagtoRegras In objRecebPagto.colRegras
'
'            Set objCondicaoPagto = New ClassCondicaoPagto
'
'            objCondicaoPagto.iCodigo = objRecebPagtoRegras.iCondPagto
'
'            'Lê a condição de pagamento
'            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'            If lErro <> SUCESSO And lErro <> 19205 Then gError 189376
'
'            'Calcula os valores das Parcelas
'            objCondicaoPagto.dValorTotal = objRecebPagto.dValor * objRecebPagtoRegras.dPercentual
'            objCondicaoPagto.dtDataRef = objRecebPagtoRegras.dtRegraValor
'
'            'Calcula os valores das Parcelas
'            lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
'            If lErro <> SUCESSO Then gError 189377
'
'            Set objRecebPagtoRegras.colParcelas = objCondicaoPagto.colParcelas
'
'        Next
'
'    Next

    For iIndice = 1 To StrParaInt(NumDias.Text)

        dtData = DateAdd("d", iIndice - 1, StrParaDate(Data.Text))
        dValorR = 0
        dValorP = 0
        dSaldoAnt = 0
        For Each objRecebPagto In colRecebPagto
            For Each objRecebPagtoRegras In objRecebPagto.colRegras
                For Each objParc In objRecebPagtoRegras.colParcelas
                    If objParc.dtVencimento = dtData Then
                        If objRecebPagto.iTipo = PRJ_TIPO_PAGTO Then
                            dValorP = dValorP + objParc.dValor
                        Else
                            dValorR = dValorR + objParc.dValor
                        End If
                    End If
                    If objParc.dtVencimento < dtData Then
                        If objRecebPagto.iTipo = PRJ_TIPO_PAGTO Then
                            dSaldoAnt = dSaldoAnt - objParc.dValor
                        Else
                            dSaldoAnt = dSaldoAnt + objParc.dValor
                        End If
                    End If
                Next
            Next
        Next
        
        GridAux.Col = iIndice
        If dValorP <> 0 Then
            bDesabilitaCmdGridAux = True
            GridAux.Row = 1
            GridAux.CellForeColor = vbRed
            bDesabilitaCmdGridAux = False
            GridAux.TextMatrix(1, iIndice) = Format(dValorP, "STANDARD")
        End If
        If dValorR <> 0 Then
            bDesabilitaCmdGridAux = True
            GridAux.Row = 2
            GridAux.CellForeColor = vbBlue
            bDesabilitaCmdGridAux = False
            GridAux.TextMatrix(2, iIndice) = Format(dValorR, "STANDARD")
        End If
        If Abs(dSaldoAnt - dValorP + dValorR) > QTDE_ESTOQUE_DELTA Then
            bDesabilitaCmdGridAux = True
            GridAux.Row = 3
            GridAux.CellForeColor = &H8000&
            bDesabilitaCmdGridAux = False
            GridAux.TextMatrix(3, iIndice) = Format(dSaldoAnt - dValorP + dValorR, "STANDARD")
        End If
    Next
    
    Cronograma_Preenche_GridAux_PRJ = SUCESSO

    Exit Function

Erro_Cronograma_Preenche_GridAux_PRJ:

    bDesabilitaCmdGridAux = False

    Cronograma_Preenche_GridAux_PRJ = gErr

    Select Case gErr
    
        Case 187454, 187455, 189376, 189377

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185608)

    End Select

    Exit Function

End Function

Private Function Traz_Relacoes_Tela(ByVal objTelaGrafico As ClassTelaGrafico, ByVal objTelaGraficoImp As ClassTelaGraficoImpressao, ByVal bImpressao As Boolean) As Long

Dim lErro As Long
Dim objTelaGraficoItem As ClassTelaGraficoItens
Dim objTelaGraficoItemPre As ClassTelaGraficoItens
Dim lIndice As Long
Dim lPosInicial As Long
Dim lPosFinal As Long
Dim lX1Pre As Long
Dim lX1 As Long
Dim lY1Pre As Long
Dim lY1 As Long
Dim lX2Pre As Long
Dim lX2 As Long
Dim lY2Pre As Long
Dim lY2 As Long
Dim lA As Long
Dim lC As Long
Dim lAPre As Long
Dim lCPre As Long
Dim dtDataColLeft As Date
Dim lPosInicialAlt As Long
Dim lPosFinalAlt As Long
Dim lTopLinha2 As Long
Dim lTopLinha3 As Long
Dim bPreForaTela As Boolean
Dim bForaTela As Boolean
Dim objTelaGraficoImpItens As ClassTelaGraficoImpItens

On Error GoTo Erro_Traz_Relacoes_Tela

    If Not bImpressao Then
        For lIndice = 1 To Line1.UBound
            Line1(lIndice).Visible = False
            Line2(lIndice).Visible = False
            Line3(lIndice).Visible = False
            Seta(lIndice).Visible = False
        Next
    End If

    lIndice = 0
    
    lPosInicialAlt = GridItens.Top + iTamanhoItemAltura + AJUSTE_ALTURA_INICIAL
    lPosFinalAlt = GridItens.Top + GridItens.Height - TAMANHO_SCROLL
    lPosInicial = POSICAO_GRID_LEFT + iTamanhoItemLargura + AJUSTE_LARGURA_INICIAL
    lPosFinal = GridItens.Left + GridItens.Width - TAMANHO_SCROLL
    dtDataColLeft = DateAdd("d", GridItens.LeftCol - 1, StrParaDate(Data.Text))

    'Para cada item
    For Each objTelaGraficoItem In objTelaGrafico.colItens
    
        lX1 = objTelaGraficoItem.objControle.Left
        lX2 = lX1 + objTelaGraficoItem.objControle.Width
        lY1 = objTelaGraficoItem.objControle.Top
        lY2 = objTelaGraficoItem.objControle.Top + objTelaGraficoItem.objControle.Height
        lA = objTelaGraficoItem.objControle.Height
        lC = objTelaGraficoItem.objControle.Width
        lTopLinha3 = lY1 + lA / 2 - ESPESSURA_LINE / 2
        
        If (lTopLinha3 < lPosInicialAlt Or lTopLinha3 > lPosFinalAlt) And Not bImpressao Then
            bForaTela = True
        Else
            bForaTela = False
        End If
        
        For Each objTelaGraficoItemPre In objTelaGraficoItem.colPredecessores
        
            lX1Pre = objTelaGraficoItemPre.objControle.Left
            lX2Pre = lX1Pre + objTelaGraficoItemPre.objControle.Width
            lY1Pre = objTelaGraficoItemPre.objControle.Top
            lY2Pre = objTelaGraficoItemPre.objControle.Top + objTelaGraficoItemPre.objControle.Height
            lAPre = objTelaGraficoItemPre.objControle.Height
            lCPre = objTelaGraficoItemPre.objControle.Width
            lTopLinha2 = lY1Pre + 2 * lAPre / 3
            
            If (lTopLinha2 < lPosInicialAlt Or lTopLinha2 > lPosFinalAlt) And Not bImpressao Then
                bPreForaTela = True
            Else
                bPreForaTela = False
            End If
        
            lIndice = lIndice + 1
                
            If Line1.UBound < lIndice Then
                'Inclui na tela um novo Controle para essa coluna
                Load Line1(lIndice)
                'Traz o controle recem desenhado para a frente
                Line1(lIndice).ZOrder
                'Inclui na tela um novo Controle para essa coluna
                Load Line2(lIndice)
                'Traz o controle recem desenhado para a frente
                Line2(lIndice).ZOrder
                'Inclui na tela um novo Controle para essa coluna
                Load Line3(lIndice)
                'Traz o controle recem desenhado para a frente
                Line3(lIndice).ZOrder
                'Inclui na tela um novo Controle para essa coluna
                Load Seta(lIndice)
                'Traz o controle recem desenhado para a frente
                Seta(lIndice).ZOrder
            End If
            
            Line1(lIndice).ToolTipText = objTelaGraficoItemPre.sNome & " precede " & objTelaGraficoItem.sNome
            Line2(lIndice).ToolTipText = objTelaGraficoItemPre.sNome & " precede " & objTelaGraficoItem.sNome
            Line3(lIndice).ToolTipText = objTelaGraficoItemPre.sNome & " precede " & objTelaGraficoItem.sNome
            
            'Se a etapa está na tela
            If (lX1 <> POSICAO_FORA_TELA And (Not bForaTela)) Or bImpressao Then
                'Se não precisa diminuir a linha da etapa
                If (lX1 - COMPRIMENTO_LINE3 > lPosInicial) Or bImpressao Then
                    'Coloca a linha a partir da esqueda da etapa com tamanho Default
                    Call Posiciona(objTelaGraficoImp, bImpressao, Line3(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha3, lX1, lTopLinha3)
                
                    If (lTopLinha2 >= lPosInicialAlt) Or bImpressao Then
                        If (lTopLinha2 <= lPosFinalAlt) Or bImpressao Then
                            'Coloca a linha vertical na mesma posição X
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha3, lX1 - COMPRIMENTO_LINE3, lTopLinha2)
                        Else
                            'Coloca a linha vertical na mesma posição X
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha3, lX1 - COMPRIMENTO_LINE3, lPosFinalAlt)
                        End If
                    Else
                        'Coloca a linha vertical na mesma posição X
                        Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha3, lX1 - COMPRIMENTO_LINE3, lPosInicialAlt)
                    End If
                Else
                    
                    'Se na posição atual cabe pelo menos a espessura de uma linha
                    If lPosInicial - lX1 >= ESPESSURA_LINE Then
                        Call Posiciona(objTelaGraficoImp, bImpressao, Line3(lIndice), lPosInicial, lTopLinha3, lX1, lTopLinha3)
                    End If

                End If
            
                If (lX1 - Seta(lIndice).Width > lPosInicial) Or bImpressao Then
                
                    If Not bImpressao Then
                        Seta(lIndice).Visible = True
                        Seta(lIndice).Left = lX1 - Seta(lIndice).Width
                        Seta(lIndice).Top = lY1 + (lA / 2) - (Seta(lIndice).Height / 2)
                    Else
            
                        Set objTelaGraficoImpItens = New ClassTelaGraficoImpItens

                        objTelaGraficoImpItens.lWidth = Seta(lIndice).Width
                        objTelaGraficoImpItens.lTop = lY1 + (lA / 2) - (Seta(lIndice).Height / 2)
                        objTelaGraficoImpItens.lLeft = lX1 - Seta(lIndice).Width
                        objTelaGraficoImpItens.lHeight = Seta(lIndice).Height
                        objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_SETA

                        objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
                    
                    End If
                    
                End If
                
            Else
                
                'Se a etapa está fora para cima
                If lTopLinha3 < lPosInicialAlt Then
                
                    'Se a pre não estiver fora para cima também
                    If lTopLinha2 >= lPosInicialAlt Then
                        
                        'Se passou do final
                        If lTopLinha2 > lPosFinalAlt Then
                            'Coloca a linha vertical na mesma posição X
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lPosFinalAlt, lX1 - COMPRIMENTO_LINE3, lPosInicialAlt)
                        Else
                            'Coloca a linha vertical na mesma posição X
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha2, lX1 - COMPRIMENTO_LINE3, lPosInicialAlt)
                        End If
                    
                    End If
                    
                Else
                'Está fora para baixo
                
                    'Se a pre não estiver fora para baixo também
                    If lTopLinha2 <= lPosFinalAlt Then
                        
                        'Se não passou do início
                        If lTopLinha2 <= lPosInicialAlt Then
                            If (lX1 - COMPRIMENTO_LINE3 >= lPosInicial) Then
                                'Coloca a linha vertical na mesma posição X
                                Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lPosInicialAlt, lX1 - COMPRIMENTO_LINE3, lPosFinalAlt)
                            End If
                        Else
                            If (lX1 - COMPRIMENTO_LINE3 >= lPosInicial) Then
                            'Coloca a linha vertical na mesma posição X
                                Call Posiciona(objTelaGraficoImp, bImpressao, Line1(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha2, lX1 - COMPRIMENTO_LINE3, lPosFinalAlt)
                            End If
                        End If
                    
                    End If
                
                End If
                
            End If
            
            If (Not bPreForaTela) Or bImpressao Then
                'Se o predecessor não está fora da tela
                If (lX1Pre <> POSICAO_FORA_TELA) Or bImpressao Then
                    'Se a etapa também não estiver
                    If (lX1 <> POSICAO_FORA_TELA) Or bImpressao Then
                        If (lX1 - COMPRIMENTO_LINE3 < lPosInicial) And Not bImpressao Then
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lX2Pre, lTopLinha2, lPosInicial, lTopLinha2)
                        Else
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lX2Pre, lTopLinha2, lX1 - COMPRIMENTO_LINE3, lTopLinha2)
                        End If
                    Else
                        'Se a etapa está fora para direita
                        If dtDataColLeft < objTelaGraficoItem.dtDataInicio Then
                            'Joga relação para o final do grid
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lX2Pre, lTopLinha2, lPosFinal, lTopLinha2)
                        Else
                            'Joga relação para o início do grid
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lX2Pre, lTopLinha2, lPosInicial, lTopLinha2)
                        End If
                    End If
                Else
                'Se o predecessor está fora da tela
                    'Se a etapa também
                    If lX1 = POSICAO_FORA_TELA Then
                                    
                        'Se as etapas relacionadas estão em cantos opostos do grid, será atravessada uma linha pela tela
                        If ((dtDataColLeft < objTelaGraficoItemPre.dtDataFim And dtDataColLeft > objTelaGraficoItem.dtDataInicio) Or (dtDataColLeft > objTelaGraficoItemPre.dtDataFim And dtDataColLeft < objTelaGraficoItem.dtDataInicio)) Then
                        
                            Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lPosInicial, lTopLinha2, lPosFinal, lTopLinha2)
                        End If
                    Else
                        'Se o predecessor está antes
                        If objTelaGraficoItemPre.dtDataFim <= objTelaGraficoItem.dtDataInicio Then
                            If lX1 - COMPRIMENTO_LINE3 > lPosInicial Then
                                Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lPosInicial, lTopLinha2, lX1 - COMPRIMENTO_LINE3, lTopLinha2)
                            End If
                        Else
                            If lX1 - COMPRIMENTO_LINE3 > lPosInicial Then
                                Call Posiciona(objTelaGraficoImp, bImpressao, Line2(lIndice), lX1 - COMPRIMENTO_LINE3, lTopLinha2, lPosFinal, lTopLinha2)
                            End If
                        End If
                    End If
                End If
            End If
        Next

    Next

    Traz_Relacoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Relacoes_Tela:

    Traz_Relacoes_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185739)

    End Select

    Exit Function

End Function

Private Function Posiciona(ByVal objTelaGraficoImp As ClassTelaGraficoImpressao, ByVal bImpressao As Boolean, ByVal objControle As Object, ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long) As Long

Dim lErro As Long
Dim objTelaGraficoImpItens As New ClassTelaGraficoImpItens

On Error GoTo Erro_Posiciona

    If Not bImpressao Then
        objControle.Visible = True
    Else
        objTelaGraficoImpItens.lBackColor = objControle.BackColor
        objTelaGraficoImpItens.iBorderStyle = objControle.BorderStyle
        objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_LINE
        'Set objTelaGraficoImpItens.objPicture = objControle
                                 
        objTelaGraficoImp.colItens.Add objTelaGraficoImpItens
    End If

    If lX1 < lX2 Then
        If Not bImpressao Then
            objControle.Left = lX1
            objControle.Width = lX2 - lX1
        Else
            objTelaGraficoImpItens.lLeft = lX1
            objTelaGraficoImpItens.lWidth = lX2 - lX1
        End If
    Else
        If Not bImpressao Then
            objControle.Left = lX2
            objControle.Width = lX1 - lX2
        Else
            objTelaGraficoImpItens.lLeft = lX2
            objTelaGraficoImpItens.lWidth = lX1 - lX2
        End If
    End If
    If Not bImpressao Then
        If objControle.Width < ESPESSURA_LINE Then objControle.Width = ESPESSURA_LINE
    End If
    If lY1 < lY2 Then
        If Not bImpressao Then
            objControle.Top = lY1
            objControle.Height = lY2 - lY1 + ESPESSURA_LINE
        Else
            objTelaGraficoImpItens.lTop = lY1
            objTelaGraficoImpItens.lHeight = lY2 - lY1
        End If
    Else
        If Not bImpressao Then
            objControle.Top = lY2
            objControle.Height = lY1 - lY2 + ESPESSURA_LINE
        Else
            objTelaGraficoImpItens.lTop = lY2
            objTelaGraficoImpItens.lHeight = lY1 - lY2
        End If
    End If

    Posiciona = SUCESSO

    Exit Function

Erro_Posiciona:

    Posiciona = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185741)

    End Select

    Exit Function

End Function

Public Sub BotaoImprimir_Click()
    GL_objMDIForm.MousePointer = vbHourglass
    Call Traz_Grafico_Tela(gobjTelaGrafico, True)
    GL_objMDIForm.MousePointer = vbDefault
End Sub

Private Function Unload_Object(ByVal objControle As Object, ByVal lIndice As Long) As Long

On Error GoTo Erro_Unload_Object

    VB.Unload objControle.Item(lIndice)

    Unload_Object = SUCESSO

    Exit Function

Erro_Unload_Object:

    Unload_Object = gErr

End Function
