VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RateioOnOcx 
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   9435
   Begin VB.CommandButton BotaoConta 
      Caption         =   "Plano de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7245
      TabIndex        =   33
      Top             =   3255
      Width           =   1605
   End
   Begin VB.CommandButton BotaoCcl 
      Caption         =   "Centros de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7245
      TabIndex        =   32
      Top             =   3870
      Width           =   1605
   End
   Begin VB.CommandButton BotaoHist 
      Caption         =   "Históricos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7230
      TabIndex        =   31
      Top             =   4485
      Width           =   1605
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1875
      Picture         =   "RateioOnOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   135
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7185
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RateioOnOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RateioOnOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RateioOnOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RateioOnOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListHistoricos 
      Height          =   3570
      Left            =   6570
      TabIndex        =   9
      Top             =   1350
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4650
      MaxLength       =   150
      TabIndex        =   7
      Top             =   1920
      Width           =   3405
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrição do Elemento Selecionado"
      Height          =   1050
      Left            =   165
      TabIndex        =   17
      Top             =   3885
      Width           =   6315
      Begin VB.Label CclDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   18
         Top             =   645
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label ContaDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1845
         TabIndex        =   19
         Top             =   285
         Width           =   3720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
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
         Left            =   1125
         TabIndex        =   20
         Top             =   300
         Width           =   570
      End
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         TabIndex        =   21
         Top             =   660
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin MSMask.MaskEdBox Debito 
      Height          =   225
      Left            =   3450
      TabIndex        =   6
      Top             =   1890
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0%"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Credito 
      Height          =   225
      Left            =   2280
      TabIndex        =   5
      Top             =   1860
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0%"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Top             =   1860
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   225
      Left            =   195
      TabIndex        =   3
      Top             =   1845
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Rateio 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1245
      TabIndex        =   2
      Top             =   540
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
      Height          =   1860
      Left            =   135
      TabIndex        =   8
      Top             =   1200
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3570
      Left            =   6570
      TabIndex        =   11
      Top             =   1350
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   6297
      _Version        =   393217
      Indentation     =   511
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   3570
      Left            =   6570
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   6297
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabelCcl 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo / Lucro"
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
      Left            =   6600
      TabIndex        =   22
      Top             =   1095
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label LabelHistoricos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Históricos"
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
      Left            =   6600
      TabIndex        =   23
      Top             =   1095
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelContas 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   6600
      TabIndex        =   24
      Top             =   1095
      Width           =   1410
   End
   Begin VB.Label LabelTotais 
      AutoSize        =   -1  'True
      Caption         =   "Totais:"
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
      Left            =   1635
      TabIndex        =   25
      Top             =   3300
      Width           =   600
   End
   Begin VB.Label TotalDebito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2445
      TabIndex        =   26
      Top             =   3285
      Width           =   1155
   End
   Begin VB.Label TotalCredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3750
      TabIndex        =   27
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Descrição:"
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
      Left            =   255
      TabIndex        =   28
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   555
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   135
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rateios"
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
      Left            =   165
      TabIndex        =   30
      Top             =   960
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -525
      X2              =   9570
      Y1              =   930
      Y2              =   930
   End
End
Attribute VB_Name = "RateioOnOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Fernando, favor subir essa constante
Const DELTA_CONTABIL = 0.000000000000001

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Private WithEvents objEventoRateioOn As AdmEvento
Attribute objEventoRateioOn.VB_VarHelpID = -1

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoHist As AdmEvento
Attribute objEventoHist.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'mostra código do proximo Rateio disponível
    lErro = CF("RateioOn_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57518

    Rateio.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57518
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166170)
    
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Form_Load
  
    Set objEventoRateioOn = New AdmEvento
    
    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoHist = New AdmEvento
    
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    lErro = Inicializa_Grid_Lancamentos(objGrid1)
    If lErro <> SUCESSO Then Error 11121


    TvwContas.Visible = False
    LabelContas.Visible = False
    TvwCcls.Visible = False
    LabelCcl.Visible = False
    ListHistoricos.Visible = False
    LabelHistoricos.Visible = False


'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 11122
'
'    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
'
'        'Inicializa a Lista de Centros de Custo
'        lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
'        If lErro <> SUCESSO Then Error 11123
'
'    End If
'
'    'Inicializa a Lista de Historicos
'    lErro = Carga_Lista_Historico()
'    If lErro <> SUCESSO Then Error 11124

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        CclLabel.Visible = True
        CclDescricao.Visible = True
    End If
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 11121, 11122, 11123, 11124
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166171)
    
    End Select
        
    iAlterado = 0
        
    Exit Sub

End Sub

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 10511
    
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 10512

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 10513
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "S" & sCclPai
            
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
            'se o centro de custo/lucro não possui centro de custo/lucro "pai"
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 10511
        
        Case 10512
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 10513
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166172)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_Historico
    
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 11130
    
    For Each objHistPadrao In colHistPadrao
        
        ListHistoricos.AddItem CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
        
    Next
    
    Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Carga_Lista_Historico:

    Carga_Lista_Historico = Err

    Select Case Err

        Case 11130

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166173)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Lancamentos(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Lancamentos
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("Histórico")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Debito.Name)
    objGridInt.colCampo.Add (Credito.Name)
    objGridInt.colCampo.Add (Historico.Name)
    
    'indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Debito_Col = 3
        iGrid_Credito_Col = 4
        iGrid_Historico_Col = 5
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Debito_Col = 2
        iGrid_Credito_Col = 3
        iGrid_Historico_Col = 4
        Ccl.Visible = False
    End If
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 11133
    
    objGridInt.objGrid = GridLancamentos
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridLancamentos.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalDebito.Top = GridLancamentos.Top + GridLancamentos.Height
    TotalDebito.Left = GridLancamentos.Left
    For iIndice = 0 To iGrid_Debito_Col - 1
        TotalDebito.Left = TotalDebito.Left + GridLancamentos.ColWidth(iIndice) + GridLancamentos.GridLineWidth + 20
    Next
    
    TotalDebito.Width = GridLancamentos.ColWidth(iGrid_Debito_Col)
    
    TotalCredito.Top = TotalDebito.Top
    TotalCredito.Left = TotalDebito.Left + TotalDebito.Width + GridLancamentos.GridLineWidth
    TotalCredito.Width = GridLancamentos.ColWidth(iGrid_Credito_Col)
    
    LabelTotais.Top = TotalCredito.Top + (TotalDebito.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalDebito.Left - LabelTotais.Width

    Inicializa_Grid_Lancamentos = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Lancamentos:

    Inicializa_Grid_Lancamentos = Err
    
    Select Case Err
    
        Case 11133
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166174)
        
    End Select

    Exit Function
        
End Function

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras
   
    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 11134
    
    Conta.Mask = sMascaraConta
    
    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 11135

        Ccl.Mask = sMascaraCcl
        
    End If
    
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 11134, 11135
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166175)
        
    End Select

    Exit Function
    
End Function

Function Trata_Parametros(Optional objRateioOn As ClassRateioOn) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Rateio passado como parametro, exibir seus dados
    If Not (objRateioOn Is Nothing) Then
    
        lErro = Traz_Doc_Tela(objRateioOn)
        If lErro <> SUCESSO And lErro <> 11243 Then Error 11131
    
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
    
    Else
                
        iAlterado = 0
        
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 11131
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166176)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Traz_Doc_Tela(objRateioOn As ClassRateioOn) As Long
'traz os dados do rateio do banco de dados para a tela

Dim lErro As Long
Dim colRateioOns As New Collection
Dim objRateioOn1 As ClassRateioOn
Dim iLinha As Integer
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim iIndice As Integer
Dim dValor As Double
Dim sDescricao As String
Dim dTotalCredito As Double
Dim dTotalDebito As Double

On Error GoTo Erro_Traz_Doc_Tela

    lErro = CF("RateioOn_Le_Doc", objRateioOn, colRateioOns)
    If lErro <> SUCESSO And lErro <> 11136 Then Error 11240
    
    Call Limpa_Tela_RateioOn
    
    'move os dados para a tela
    
    Rateio.Text = CStr(objRateioOn.iCodigo)
        
    If lErro = SUCESSO Then
        
        dTotalCredito = 0
        dTotalCredito = 0
                
        For Each objRateioOn1 In colRateioOns
        
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objRateioOn1.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 11241
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
            
            'coloca a conta na tela
            GridLancamentos.TextMatrix(objRateioOn1.iSeq, iGrid_Conta_Col) = Conta.Text
            
            If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
                
                If objRateioOn1.sCcl <> "" Then
                
                    lErro = Mascara_MascararCcl(objRateioOn1.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then Error 11242
                
                Else
                   sCclMascarado = ""
                   
                End If
                
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objRateioOn1.iSeq, iGrid_Ccl_Col) = sCclMascarado
                
            End If
            
            sDescricao = objRateioOn1.sDescricao
                    
            'coloca o percentual na tela
            If objRateioOn1.dPercentual > 0 Then
                GridLancamentos.TextMatrix(objRateioOn1.iSeq, iGrid_Credito_Col) = Format(objRateioOn1.dPercentual, "Percent")
                dTotalCredito = dTotalCredito + objRateioOn1.dPercentual
            Else
                GridLancamentos.TextMatrix(objRateioOn1.iSeq, iGrid_Debito_Col) = Format(-objRateioOn1.dPercentual, "Percent")
                dTotalDebito = dTotalDebito - objRateioOn1.dPercentual
            End If
                
            'coloca o historico na tela
            GridLancamentos.TextMatrix(objRateioOn1.iSeq, iGrid_Historico_Col) = objRateioOn1.sHistorico
            
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
                
        Next
        
        TotalCredito.Caption = Format(dTotalCredito, "Percent")
        TotalDebito.Caption = Format(dTotalDebito, "Percent")
        
        Descricao = sDescricao
        
    End If
            
    iAlterado = 0
    
    Traz_Doc_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err
    
        Case 11240, 11242
        
        Case 11241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOn1.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166177)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Credito_Change()
   iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Debito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoRateioOn = Nothing

    Set objEventoConta = Nothing
    Set objEventoCcl = Nothing
    Set objEventoHist = Nothing

    Set objGrid1 = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GridLancamentos_LeaveCell()
    
    Call Saida_Celula(objGrid1)

End Sub

Private Sub GridLancamentos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Private Sub GridLancamentos_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Private Sub GridLancamentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim dColunaSoma As Double

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dColunaSoma, "Percent")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dColunaSoma, "Percent")
    
End Sub

Private Sub GridLancamentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridLancamentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridLancamentos.Col
    
            Case iGrid_Conta_Col
            
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then Error 11234
                
            Case iGrid_Ccl_Col
            
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then Error 11235
                
            Case iGrid_Credito_Col
            
                lErro = Saida_Celula_Credito(objGridInt)
                If lErro <> SUCESSO Then Error 11236
                
            Case iGrid_Debito_Col
            
                lErro = Saida_Celula_Debito(objGridInt)
                If lErro <> SUCESSO Then Error 11237

            Case iGrid_Historico_Col
            
                lErro = Saida_Celula_Historico(objGridInt)
                If lErro <> SUCESSO Then Error 11238
                
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 11239
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
            
        Case 11234, 11235, 11236, 11237, 11238
    
        Case 11239
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166178)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula ccl do grid que está deixando de ser a corrente

Dim sCclFormatada As String
Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sConta As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
                
    'critica o formato do ccl, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 11228
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 11232
                
    If Len(Ccl.ClipText) > 0 Then
    
        If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)) > 0 Then
    
            sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 11230
        
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
        
            
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 11231
        
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 11233
        
'                If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
'                    objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'                End If
        
        End If
                
        CclDescricao.Caption = objCcl.sDescCcl
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11229

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 11228, 11229, 11230, 11231
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 11232
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                objCcl.sCcl = sCclFormatada
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 11233
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                objContaCcl.sConta = sContaFormatada
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
    
        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166179)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim sContaFormatada As String
Dim sContaMascarada As String
Dim sCclFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sCcl As String
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", Conta.Text, Conta.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 19369
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 19370
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then
    
        'testa a conta no seu formato normal
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11221
                    
        'conta não cadastrada
        If lErro = 5700 Then Error 11226
    
    End If
               
    If Len(Conta.ClipText) > 0 Then
    
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)) > 0 Then
            
                sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
        
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 11223
        
                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada
        
                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then Error 11222
        
                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then Error 11225
                
            End If
            
        End If
                
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
                
        ContaDescricao.Caption = objPlanoConta.sDescConta
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11224

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
    
        Case 11221, 11222, 11223, 11224, 19369
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 11225
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, sCcl)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("ContaCcl", objPlanoConta)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case 11226
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
                    
        Case 19370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166180)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Debito(objGridInt As AdmGrid) As Long
'faz a critica da celula debito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dValor As Double

On Error GoTo Erro_Saida_Celula_Debito

    Set objGridInt.objControle = Debito
    
    If Len(Trim(Debito.Text)) > 0 Then
        
        lErro = Porcentagem_Critica(Debito.Text)
        If lErro <> SUCESSO Then Error 20321
        
    End If
               
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11220
                
    If Len(Trim(Debito.Text)) > 0 Then
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Credito_Col) = ""
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
   End If
    
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dColunaSoma, "Percent")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dColunaSoma, "Percent")

    Saida_Celula_Debito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Debito:

    Saida_Celula_Debito = Err
    
    Select Case Err

        Case 11220, 20321
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166181)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Historico

    Set objHistPadrao = New ClassHistPadrao
    
    Set objGridInt.objControle = Historico
                
    If Left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
    
        sValor = Trim(Mid(Historico.Text, 2))
        
        lErro = Valor_Inteiro_Critica(sValor)
        If lErro <> SUCESSO Then Error 11215
        
        objHistPadrao.iHistPadrao = CInt(sValor)
                
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 11216
        
        If lErro = 5446 Then Error 11217

        Historico.Text = objHistPadrao.sDescHistPadrao

    End If
    
    If Len(Trim(Historico.Text)) > 0 Then
            If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
    End If
    
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11218

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = Err
    
    Select Case Err
    
        Case 11215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 11216, 11218
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 11217
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166182)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Credito(objGridInt As AdmGrid) As Long
'faz a critica da celula credito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dValor As Double

On Error GoTo Erro_Saida_Celula_Credito

    Set objGridInt.objControle = Credito
    
    If Len(Trim(Credito.Text)) > 0 Then

        lErro = Porcentagem_Critica(Credito.Text)
        If lErro <> SUCESSO Then Error 20320
    
    End If
                  
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11214
              
    If Len(Trim(Credito.Text)) > 0 Then
            GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Debito_Col) = ""
            If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
    End If
                    
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dColunaSoma, "Percent")
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dColunaSoma, "Percent")

    Saida_Celula_Credito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Credito:

    Saida_Celula_Credito = Err
    
    Select Case Err
    
        Case 11214, 20320
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166183)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    For iLinha = 1 To objGrid1.iLinhasExistentes
        If Len(GridLancamentos.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(Format(GridLancamentos.TextMatrix(iLinha, iColuna), "General Number"))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Private Sub Label3_Click()

Dim objRateioOn As New ClassRateioOn
Dim colSelecao As Collection

    If Len(Rateio.Text) = 0 Then
        objRateioOn.iCodigo = 0
    Else
        objRateioOn.iCodigo = CInt(Rateio.ClipText)
    End If

    objRateioOn.iSeq = 0

    Call Chama_Tela("RateioOnLista", colSelecao, objRateioOn, objEventoRateioOn)
    
End Sub

Private Sub ListHistoricos_DblClick()

Dim lPosicaoSeparador As Long
    
    If GridLancamentos.Col = iGrid_Historico_Col Then
    
        lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
        Historico.Text = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
            
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
    
    End If
    
End Sub

Private Sub objEventoRateioOn_evSelecao(obj1 As Object)

Dim objRateioOn As ClassRateioOn
Dim lErro As Long
    
On Error GoTo Erro_objEventoRateioOn_evSelecao
    
    Set objRateioOn = obj1
    
    lErro = Traz_Doc_Tela(objRateioOn)
    If lErro <> SUCESSO Then Error 11212
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoRateioOn_evSelecao:

    Select Case Err
    
        Case 11212
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166184)
            
        End Select
        
    Exit Sub
        
End Sub

Private Sub Rateio_Change()

    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Rateio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Rateio, iAlterado)

End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String
    
On Error GoTo Erro_TvwCcls_NodeClick
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then
    
        sCaracterInicial = Left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = Right(Node.Key, Len(Node.Key) - 1)
              
            sCclEnxuta = String(STRING_CCL, 0)
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 10514
            
            Ccl.PromptInclude = False
            Ccl.Text = sCclEnxuta
            Ccl.PromptInclude = True
              
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            End If
        
            'Preenche a Descricao do centro de custo/lucro
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            CclDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    
        End If
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err
    
        Case 10514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166185)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40826

    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40826
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166186)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_TvwContas_NodeClick
    
    If GridLancamentos.Col = iGrid_Conta_Col Then
    
        sCaracterInicial = Left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sConta = Right(Node.Key, Len(Node.Key) - 1)
            
            sContaEnxuta = String(STRING_CONTA, 0)
            
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 11210
            
            Conta.PromptInclude = False
            Conta.Text = sContaEnxuta
            Conta.PromptInclude = True
        
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
                objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
            End If
        
            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            ContaDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
        End If
        
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 11210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166187)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 14730
    
    Call Limpa_Tela_RateioOn

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
                
        Case 14730
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166188)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
            
Dim lErro As Long
Dim iCodigo As Integer
Dim sDescricao As String
Dim colRateioOns As New Collection
Dim objRateioOn As New ClassRateioOn
Dim iIndice1 As Integer
Dim dSomaCredito As Double
Dim dSomaDebito As Double
        
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Campo Rateio esta preenchido
    If Len(Rateio.ClipText) = 0 Then Error 11204
        
    'Verifica se pelo menos uma linha do Grid está preenchida
    If objGrid1.iLinhasExistentes = 0 Then Error 11205
    
    'Move-se os dados da tela para as variáveis
    iCodigo = CInt(Rateio.ClipText)

    If Len(Trim(Descricao.ClipText)) = 0 Then
       
       sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    Else
       sDescricao = Descricao.Text
    End If
    
    'Preenche a colRateioOns com as informacoes contidas no Grid
    lErro = Grid_RateioOn(colRateioOns)
    If lErro <> SUCESSO Then Error 11206
       
    dSomaCredito = GridColuna_Soma(iGrid_Credito_Col)
    dSomaDebito = GridColuna_Soma(iGrid_Debito_Col)
    
    If Abs(1 - dSomaCredito) > DELTA_CONTABIL And dSomaCredito <> 0 Then Error 11144
    
    If Abs(1 - dSomaDebito) > DELTA_CONTABIL And dSomaDebito <> 0 Then Error 11143
        
    lErro = Trata_Alteracao(objRateioOn, iCodigo)
    If lErro <> SUCESSO Then Error 32303
        
        
    lErro = CF("RateioOn_Grava", iCodigo, sDescricao, colRateioOns)
    If lErro <> SUCESSO Then Error 11207
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 11143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_NAO_VALIDA", Err)
                       
        Case 11144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_NAO_VALIDA", Err)
            
        Case 11204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_RATEIO_NAO_PREENCHIDO", Err)
            Rateio.SetFocus
        
        Case 11205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_RATEIOON_GRAVAR", Err)
            
        Case 11206, 11207, 32303

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166189)
            
    End Select
    
    Exit Function
    
End Function

Function Grid_RateioOn(colRateioOns As Collection) As Long

Dim iIndice1 As Integer
Dim objRateioOn As ClassRateioOn
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim dPercentualCredito As Double
Dim dPercentualDebito As Double
Dim lErro As Long

On Error GoTo Erro_Grid_RateioOn

    For iIndice1 = 1 To objGrid1.iLinhasExistentes
        
        Set objRateioOn = New ClassRateioOn
            
        objRateioOn.iSeq = iIndice1
  
        sConta = GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col)
        
        If Len(Trim(sConta)) = 0 Then Error 11245
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11201
            
        objRateioOn.sConta = sContaFormatada
    
        'Testa para ver se houve crédito ou débito
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col)) > 0 Then
            dPercentualCredito = CDbl(Format(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col), "General Number"))
        Else
            dPercentualCredito = 0
        End If
            
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col)) > 0 Then
            dPercentualDebito = CDbl(Format(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col), "General Number"))
        Else
            dPercentualDebito = 0
        End If

        'Armazena débito ou crédito
        If dPercentualDebito = 0 And dPercentualCredito = 0 Then Error 11203
            
        objRateioOn.dPercentual = dPercentualCredito - dPercentualDebito
    
        'Armazena Histórico e Ccl
        objRateioOn.sHistorico = GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col)
            
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
                
            sCcl = GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 11202
                           
            If iCclPreenchida = CCL_PREENCHIDA Then
                objRateioOn.sCcl = sCclFormatada
            Else
                objRateioOn.sCcl = ""
            End If
                
        End If
                
        'Armazena o objeto objRateioOn na coleção colRateioOns
        colRateioOns.Add objRateioOn
        
    Next
    
    Grid_RateioOn = SUCESSO

    Exit Function

Erro_Grid_RateioOn:

    Grid_RateioOn = Err

    Select Case Err
    
        Case 11201, 11202
        
        Case 11203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO", Err)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Debito_Col
            GridLancamentos.SetFocus
                
        Case 11245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166190)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Exclui os lançamentos relativos ao Rateio digitado na tela
    
Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o campo Rateio está preenchido
    If Len(Rateio.ClipText) = 0 Then Error 11198
     
    'Envia Mensagem pedindo confirmação da Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RATEIO")
    
    If vbMsgRes = vbYes Then
    
        iCodigo = CInt(Rateio.ClipText)
        
        'Exclui todos os lancamentos daquele Rateio automático
        lErro = CF("RateioOn_Exclui", iCodigo)
        If lErro <> SUCESSO Then Error 11199
    
        Call Limpa_Tela_RateioOn
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
                
        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
       Case 11198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_RATEIO_NAO_PREENCHIDO", Err)
            Rateio.SetFocus
        
       Case 11199
    
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166191)
        
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim objRateioOn As New ClassRateioOn

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 11194

    Call Limpa_Tela_RateioOn

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 11194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166192)

    End Select

    Exit Sub
    
End Sub

Private Sub Conta_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Conta_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCcl.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
    
    If Len(sConta) > 0 Then

   lErro = Conta_Exibe_Descricao(sConta)
   If lErro <> SUCESSO Then Error 11085
        
    Else

        ContaDescricao = ""

    End If
    
    Exit Sub
    
Erro_Conta_GotFocus:

    Select Case Err
    
        Case 11085
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166193)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Ccl_GotFocus()

Dim sCcl As String
Dim lErro As Long

On Error GoTo Erro_Ccl_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwCcls.Visible = True
'    LabelCcl.Visible = True
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
    
    'Coloca descricao de Ccl no panel
    If Len(sCcl) > 0 Then

        lErro = Ccl_Exibe_Descricao(sCcl)
        If lErro <> SUCESSO Then Error 11086

    Else
    
        CclDescricao = ""
        
    End If
    
    Exit Sub
    
Erro_Ccl_GotFocus:

    Select Case Err
    
        Case 11086
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166194)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Credito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Credito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Credito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Credito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Debito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Debito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Debito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Debito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Historico_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwCcls.Visible = False
'    LabelCcl.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListHistoricos.Visible = True
'    LabelHistoricos.Visible = True
    
End Sub

Private Sub Historico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Historico
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridLancamentos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Function Limpa_Tela_RateioOn() As Long

Dim lErro As Long

    Call Grid_Limpa(objGrid1)
    
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""
    Descricao.Text = ""
    Rateio.Text = ""

End Function

Function Conta_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Conta_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 11189
    
    lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 11190

    If lErro = 6030 Then Error 11191
    
    ContaDescricao.Caption = objPlanoConta.sDescConta
    
    Conta_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Conta_Exibe_Descricao:

    Conta_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11189, 11190
            ContaDescricao = ""
            
        Case 11191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            ContaDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166195)
            
    End Select
        
    Exit Function

End Function

Function Ccl_Exibe_Descricao(sCcl As String) As Long
'exibe a descrição do centro de custo/lucro no campo CclDescricao. O ccl passado como parametro deve estar mascarado

Dim sCclFormatada As String
Dim sCclArvore As String
Dim objNode As Node
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Exibe_Descricao

    'Retorna Ccl formatada como no BD
    lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 11186
    
    objCcl.sCcl = sCclFormatada

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 11187
    
    If lErro = 5599 Then Error 11188
    
    CclDescricao.Caption = objCcl.sDescCcl
    
    Ccl_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Ccl_Exibe_Descricao:

    Ccl_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11186, 11187
            CclDescricao = ""
            
        Case 11188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            CclDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166196)
            
    End Select
        
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim colRateioOns As New Collection
Dim objRateioOn As New ClassRateioOn

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RateiosOnLine"
    
    If Len(Trim(Rateio.ClipText)) > 0 Then
        objRateioOn.iCodigo = CInt(Rateio.ClipText)
    Else
        objRateioOn.iCodigo = 0
    End If

    If Len(Trim(Descricao.ClipText)) = 0 Then
       
       objRateioOn.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
    
    Else
       objRateioOn.sDescricao = Descricao.Text
       
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objRateioOn.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objRateioOn.sDescricao, STRING_RATEIO_DESCRICAO, "Descricao"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166197)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objRateioOn As New ClassRateioOn

On Error GoTo Erro_Tela_Preenche

    objRateioOn.iCodigo = colCampoValor.Item("Codigo").vValor

    If objRateioOn.iCodigo <> 0 Then
    
        objRateioOn.sDescricao = colCampoValor.Item("Descricao").vValor
            
        lErro = Traz_Doc_Tela(objRateioOn)
        If lErro <> SUCESSO Then Error 24158

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 24158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166198)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RATEIO_ON_LINE
    Set Form_Load_Ocx = Me
    Caption = "Rateio On-Line"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RateioOn"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Rateio Then
            Call Label3_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Historico Then
            Call BotaoHist_Click
        End If
    
    End If

End Sub

Private Sub CclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDescricao, Source, X, Y)
End Sub

Private Sub CclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDescricao, Button, Shift, X, Y)
End Sub

Private Sub ContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaDescricao, Source, X, Y)
End Sub

Private Sub ContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub TotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalDebito, Source, X, Y)
End Sub

Private Sub TotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalDebito, Button, Shift, X, Y)
End Sub

Private Sub TotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCredito, Source, X, Y)
End Sub

Private Sub TotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCredito, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub BotaoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Conta.Text) > 0 Then objPlanoConta.sConta = Conta.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_objEventoConta_evSelecao
    
    If GridLancamentos.Col = iGrid_Conta_Col Then

        Set objPlanoConta = obj1
        
        sConta = objPlanoConta.sConta
        
        'le a conta
        lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 6030 Then gError 197949
        
        If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197950
        
        If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197951
        
        sContaEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 197952

        Conta.PromptInclude = False
        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

        ContaDescricao.Caption = objPlanoConta.sDescConta
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            objHistPadrao.iHistPadrao = objPlanoConta.iHistPadrao
                        
            'le os dados do historico
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then gError 197953
                                    
            If lErro = SUCESSO Then
            
                GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = objHistPadrao.sDescHistPadrao
                
            End If
                        
        End If

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197949, 197953
    
        Case 197950
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197951
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case 197952
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197954)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Ccl.Text) > 0 Then objCcl.sCcl = Ccl.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sConta As String
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then

        Set objCcl = obj1

        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 197955

        If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197956
        
        If objCcl.iAtivo = 0 Then gError 197957
        
        sCclEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 197958

        Ccl.PromptInclude = False
        Ccl.Text = sCclEnxuta
        Ccl.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

        CclDescricao.Caption = objCcl.sDescCcl

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197955

        Case 197956
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197957
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197958
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197959)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoHist_Click()

Dim colSelecao As Collection
Dim objHistPadrao As New ClassHistPadrao

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHist)

End Sub

Private Sub objEventoHist_evSelecao(obj1 As Object)


Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_objEventoHist_evSelecao

    If GridLancamentos.Col = iGrid_Historico_Col Then

        Set objHistPadrao = obj1

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = objHistPadrao.sDescHistPadrao
        Historico.Text = objHistPadrao.sDescHistPadrao

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.iLinhasExistentes Then
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoHist_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197960)

    End Select

    Exit Sub

End Sub

