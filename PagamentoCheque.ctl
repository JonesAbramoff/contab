VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoCheque 
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ForeColor       =   &H00000080&
   KeyPreview      =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9390
   Begin VB.CommandButton BotaoValidar 
      Caption         =   "(F7)  Validar"
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
      Left            =   7935
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1350
   End
   Begin VB.CheckBox CheckFixar 
      Caption         =   "Fixar"
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
      Left            =   3420
      TabIndex        =   31
      Top             =   1170
      Width           =   795
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
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
      Left            =   6000
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4470
      Width           =   1725
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
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
      Left            =   3870
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4470
      Width           =   1725
   End
   Begin VB.CommandButton BotaoImprimir 
      Caption         =   "(F4)   Imprimir"
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
      Left            =   1590
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4470
      Width           =   1725
   End
   Begin VB.CommandButton BotaoLe 
      Caption         =   "(F2)  Ler"
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
      Left            =   7935
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   1350
   End
   Begin VB.Frame FrameCheque 
      Caption         =   "Cheques"
      Height          =   2625
      Left            =   165
      TabIndex        =   22
      Top             =   1560
      Width           =   9105
      Begin MSMask.MaskEdBox StatusGrid 
         Height          =   255
         Left            =   7350
         TabIndex        =   32
         Top             =   255
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   255
         Left            =   5880
         TabIndex        =   28
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteGrid 
         Height          =   255
         Left            =   4530
         TabIndex        =   24
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox DataBomParaGrid 
         Height          =   255
         Left            =   3495
         TabIndex        =   23
         Top             =   240
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaGrid 
         Height          =   255
         Left            =   1830
         TabIndex        =   13
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox NumeroGrid 
         Height          =   255
         Left            =   2655
         TabIndex        =   14
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox AgenciaGrid 
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox BancoGrid 
         Height          =   255
         Left            =   435
         TabIndex        =   11
         Top             =   240
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid GridCheques 
         Height          =   1980
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3493
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
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
      Begin VB.Label TotalCheque 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5880
         TabIndex        =   30
         Top             =   2250
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Cheques: "
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
         Left            =   4515
         TabIndex        =   29
         Top             =   2295
         Width           =   1365
      End
   End
   Begin VB.CommandButton BotaoIncluir 
      Caption         =   "(F6)  Incluir"
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
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Width           =   1350
   End
   Begin MSMask.MaskEdBox Agencia 
      Height          =   300
      Left            =   3405
      TabIndex        =   1
      Top             =   195
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   300
      Left            =   5985
      TabIndex        =   2
      Top             =   195
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   3405
      TabIndex        =   4
      Top             =   675
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataBomPara 
      Height          =   300
      Left            =   5985
      TabIndex        =   5
      Top             =   645
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown SpinData 
      Height          =   315
      Left            =   7155
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   630
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   1110
      TabIndex        =   7
      ToolTipText     =   "CGC/CPF do Cliente"
      Top             =   1125
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      Mask            =   "##############"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   1110
      TabIndex        =   3
      Top             =   660
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Banco 
      Height          =   300
      Left            =   1110
      TabIndex        =   0
      Top             =   195
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5985
      TabIndex        =   35
      Top             =   1125
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Index           =   7
      Left            =   5340
      TabIndex        =   34
      Top             =   1170
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
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
      Index           =   0
      Left            =   435
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
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
      Index           =   3
      Left            =   2595
      TabIndex        =   20
      Top             =   225
      Width           =   765
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   5385
      TabIndex        =   19
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   315
      TabIndex        =   18
      Top             =   705
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Index           =   4
      Left            =   2835
      TabIndex        =   17
      Top             =   690
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
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
      Index           =   2
      Left            =   5070
      TabIndex        =   16
      Top             =   690
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CNPJ:"
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
      Index           =   6
      Left            =   90
      TabIndex        =   15
      Top             =   1170
      Width           =   975
   End
End
Attribute VB_Name = "PagamentoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Public iAlterado As Integer
Dim giTipo As Integer

'Variável que guarda as características do grid da tela
Dim objGridCheques As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_Conta_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataBomPara_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Status_Col As Integer

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objCheque As ClassChequePre
Dim iIndice As Integer

    Set gobjVenda = objVenda
    
    giTipo = MOVIMENTOCAIXA_RECEB_CHEQUE
    
'    'Se o projeto <> SGEECF
'    If gsNomePrinc <> "SGEECF" Then giTipo = MOVIMENTOCAIXA_RECEB_CARNE_CHEQUE
        
    'Joga na tela todos os Chequess
    For Each objCheque In gobjVenda.colCheques
               
        If objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
        
            objGridCheques.iLinhasExistentes = objGridCheques.iLinhasExistentes + 1
                
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Banco_Col) = objCheque.iBanco
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Agencia_Col) = objCheque.sAgencia
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Conta_Col) = objCheque.sContaCorrente
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Numero_Col) = objCheque.lNumero
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_DataBomPara_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Valor_Col) = Format(objCheque.dValor, "standard")
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Cliente_Col) = objCheque.sCPFCGC
            GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Status_Col) = IIf(objCheque.iAprovado = CHEQUE_APROVADO, STRING_CHEQUE_APROVADO, STRING_CHEQUE_NAO_APROVADO)

        End If
        
    Next
    
    Call Atualiza_Total
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
    
Dim sRetorno As String
Dim lTamanho As Long

On Error GoTo Erro_Form_Load

    Set objGridCheques = New AdmGrid
        
    Call Inicializa_Grid_Cheques(objGridCheques)
        
    lTamanho = 1
    sRetorno = String(lTamanho, 0)
        
    Call GetPrivateProfileString(APLICACAO_CAIXA, "ChequeFixar", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then CheckFixar.Value = StrParaInt(sRetorno)
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164176)

    End Select

    Exit Sub

End Sub

Function Inicializa_Grid_Cheques(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Banco")
    objGridInt.colColuna.Add ("Agencia")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Bom Para")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Status")
      
    'Controles que participam do Grid
    objGridInt.colCampo.Add (BancoGrid.Name)
    objGridInt.colCampo.Add (AgenciaGrid.Name)
    objGridInt.colCampo.Add (ContaGrid.Name)
    objGridInt.colCampo.Add (NumeroGrid.Name)
    objGridInt.colCampo.Add (DataBomParaGrid.Name)
    objGridInt.colCampo.Add (ClienteGrid.Name)
    objGridInt.colCampo.Add (ValorGrid.Name)
    objGridInt.colCampo.Add (StatusGrid.Name)
    
    'Colunas do Grid
    iGrid_Banco_Col = 1
    iGrid_Agencia_Col = 2
    iGrid_Conta_Col = 3
    iGrid_Numero_Col = 4
    iGrid_DataBomPara_Col = 5
    iGrid_Cliente_Col = 6
    iGrid_Valor_Col = 7
    iGrid_Status_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridCheques

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_CHEQUES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridCheques.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Cheques = SUCESSO

    Exit Function

End Function

Private Sub Agencia_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub Banco_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

Private Sub BotaoImprimir_Click()
        
Dim sMsg As String
        
    If giImpressoraCheque = IMPRESSORA_PRESENTE Then
        Call AFRAC_ChequeImprimir(Banco.Text, Valor.Text, gsNomeEmpresa, gsCidade, DataBomPara.Text, sMsg)
    End If

End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoIncluir_Click
    
    'verificação do preenchimento dos campos
    If Len(Trim(Banco.Text)) = 0 Then gError 99790
    If Len(Trim(Agencia.Text)) = 0 Then gError 99791
    If Len(Trim(Conta.Text)) = 0 Then gError 99792
    If Len(Trim(Numero.Text)) = 0 Then gError 99793
    
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99641
    
    'Se quantidade não preenchido --> Erro.
    If Len(Trim(DataBomPara.ClipText)) = 0 Then gError 99642
    
    'verifica se o valor pago ultrapassa o valor minimo da condicao de pagto
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CHEQUE Then
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                    If StrParaDbl(Valor.Text) < objAdmMeioPagtoCondPagto.dValorMinimo Then gError 126818
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    objGridCheques.iLinhasExistentes = objGridCheques.iLinhasExistentes + 1
    
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Banco_Col) = Banco.Text
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Agencia_Col) = Agencia.Text
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Conta_Col) = Conta.Text
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Numero_Col) = Numero.Text
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_DataBomPara_Col) = Format(DataBomPara.Text, "dd/mm/yyyy")
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Cliente_Col) = Cliente.Text
    GridCheques.TextMatrix(objGridCheques.iLinhasExistentes, iGrid_Status_Col) = Status.Caption
       
    Call Atualiza_Total
    
    If CheckFixar.Value = DESMARCADO Then
    
        'Limpa os campos da tela
        Banco.Text = ""
        Agencia.Text = ""
        Conta.Text = ""
        Numero.Text = ""
        Valor.Text = ""
        Cliente.Text = ""
        
        DataBomPara.PromptInclude = False
        DataBomPara.Text = ""
        DataBomPara.PromptInclude = True
    
    End If
    
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 99641
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case 99642
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
            
        Case 99790
            Call Rotina_ErroECF(vbOKOnly, ERRO_BANCO_NAO_PREENCHIDO, gErr)
            
        Case 99791
            Call Rotina_ErroECF(vbOKOnly, ERRO_AGENCIA_NAO_PREENCHIDA, gErr)
            
        Case 99792
            Call Rotina_ErroECF(vbOKOnly, ERRO_CONTA_NAO_PREENCHIDA, gErr)
            
        Case 99793
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_PREENCHIDO, gErr)
            
        Case 99794
            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_PREENCHIDO1, gErr)
            
        Case 126818
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORMINIMO_CONDPAGTO, gErr, objAdmMeioPagtoCondPagto.dValorMinimo, Valor.Text)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164177)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objCheques As ClassChequePre
Dim iIndice As Integer
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice2 As Integer
Dim lTamanho As Long
Dim sRetorno As String

On Error GoTo Erro_BotaoOk_Click
            
    If Not gobjVenda Is Nothing Then
            
        'Exclui todos os movimentos em cheque especificados
        For iIndice = gobjVenda.colCheques.Count To 1 Step -1
            Set objCheques = gobjVenda.colCheques.Item(iIndice)
            If objCheques.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
                gobjVenda.colCheques.Remove (iIndice)
                For iIndice2 = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
                    Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice2)
                    'remove os movimentos de caixa relacionado ao cheque excluído
                    If objMovimento.iTipo = giTipo And objMovimento.lNumRefInterna = objCheques.lSequencialCaixa Then gobjVenda.colMovimentosCaixa.Remove (iIndice2)
                Next
            End If
        Next
        
        'Para cada linha do grid...
        For iIndice = 1 To objGridCheques.iLinhasExistentes
                
            Set objCheques = New ClassChequePre
        
            'Insere um novo movimento
            objCheques.dtDataDeposito = StrParaDate(GridCheques.TextMatrix(iIndice, iGrid_DataBomPara_Col))
            objCheques.dValor = StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            objCheques.iBanco = StrParaInt(GridCheques.TextMatrix(iIndice, iGrid_Banco_Col))
            objCheques.iFilialEmpresaLoja = giFilialEmpresa
            objCheques.sCPFCGC = GridCheques.TextMatrix(iIndice, iGrid_Cliente_Col)
            objCheques.lNumero = StrParaLong(GridCheques.TextMatrix(iIndice, iGrid_Numero_Col))
            objCheques.sAgencia = GridCheques.TextMatrix(iIndice, iGrid_Agencia_Col)
            objCheques.sContaCorrente = GridCheques.TextMatrix(iIndice, iGrid_Conta_Col)
            objCheques.sCPFCGC = GridCheques.TextMatrix(iIndice, iGrid_Cliente_Col)
            objCheques.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
            objCheques.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
            objCheques.iAprovado = IIf(GridCheques.TextMatrix(iIndice, iGrid_Status_Col) = STRING_CHEQUE_APROVADO, CHEQUE_APROVADO, CHEQUE_NAO_APROVADO)
            
            lTamanho = 50
            sRetorno = String(lTamanho, 0)
    
            Call GetPrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
            If sRetorno <> String(lTamanho, 0) Then objCheques.lSequencialCaixa = StrParaLong(sRetorno)
            
            If objCheques.lSequencialCaixa = 0 Then objCheques.lSequencialCaixa = 1
            
            'Atualiza o sequencial de arquivo
            lErro = WritePrivateProfileString(APLICACAO_CAIXA, "NumProxCheque", CStr(objCheques.lSequencialCaixa + 1), NOME_ARQUIVO_CAIXA)
            If lErro = 0 Then gError 105774
            
            gobjVenda.colCheques.Add objCheques
                    
            objCheques.iNaoEspecificado = CHEQUE_ESPECIFICADO
                   
            'criar movimento para cada cheque
            Set objMovimento = New ClassMovimentoCaixa
        
            'Insere um novo movimento
            objMovimento.iFilialEmpresa = giFilialEmpresa
            objMovimento.iCaixa = giCodCaixa
            objMovimento.iCodOperador = giCodOperador
            objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
            objMovimento.iTipo = giTipo
            objMovimento.iParcelamento = COD_A_VISTA
            objMovimento.dtDataMovimento = Date
            objMovimento.dValor = StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            objMovimento.dHora = CDbl(Time)
            objMovimento.lNumRefInterna = objCheques.lSequencialCaixa
            objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
            objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
            If objCheques.iAprovado = CHEQUE_APROVADO Then
                objMovimento.iTipoCartao = TIPOMEIOPAGTOLOJA_TEFCHQ
                objMovimento.iIndiceImpChq = gobjVenda.colIndiceImpCheque.Item(1)
                gobjVenda.colIndiceImpCheque.Remove (1)
            End If
            
            gobjVenda.colMovimentosCaixa.Add objMovimento
            
        Next
        
        Unload Me
    
    End If
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 105774
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_CAIXA, "NumProxCheque", NOME_ARQUIVO_CAIXA)

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164178)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalCheque.Caption = ""
    
    For iIndice = 1 To objGridCheques.iLinhasExistentes
        TotalCheque.Caption = Format(StrParaDbl(TotalCheque.Caption) + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col)), "standard")
    Next
    
End Sub

Private Sub BotaoValidar_Click()

Dim objCheque As New ClassChequePre
Dim objMsg As Object
Dim objTela As Object
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim lErro As Long

On Error GoTo Erro_BotaoValidar_Click

    If gobjVenda.objCupomFiscal.lNumero = 0 Then gError 133770

    Set objTela = Me
    Set objMsg = MsgTEF
    
    'verificação do preenchimento dos campos
    If Len(Trim(Banco.Text)) = 0 Then gError 133762
    If Len(Trim(Agencia.Text)) = 0 Then gError 133763
    If Len(Trim(Conta.Text)) = 0 Then gError 133764
    If Len(Trim(Numero.Text)) = 0 Then gError 133765
    
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 133766
    
    'Se quantidade não preenchido --> Erro.
    If Len(Trim(DataBomPara.ClipText)) = 0 Then gError 133767
    
    'verifica se o valor pago ultrapassa o valor minimo da condicao de pagto
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CHEQUE Then
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                    If StrParaDbl(Valor.Text) < objAdmMeioPagtoCondPagto.dValorMinimo Then gError 133768
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    objCheque.dtDataDeposito = StrParaDate(DataBomPara.Text)
    objCheque.dValor = StrParaDbl(Valor.Text)
    objCheque.iBanco = StrParaInt(Banco.Text)
    objCheque.sCPFCGC = Cliente.Text
    objCheque.lNumero = StrParaLong(Numero.Text)
    objCheque.sAgencia = Agencia.Text
    objCheque.sContaCorrente = Conta.Text
    
    If StrParaDbl(TotalCheque.Caption) + objCheque.dValor > gobjVenda.dFalta Then gError 126595
    
    lErro = CF_ECF("TEF_CHQ", objMsg, objTela, objCheque, gobjVenda)
    If lErro <> SUCESSO Then gError 133769
    
    Status.Caption = IIf(objCheque.iAprovado = CHEQUE_APROVADO, STRING_CHEQUE_APROVADO, STRING_CHEQUE_NAO_APROVADO)
    
    If objCheque.iAprovado = CHEQUE_APROVADO Then
    
        Call BotaoIncluir_Click
    
        If StrParaDbl(TotalCheque.Caption) = gobjVenda.dFalta Then
            Call BotaoOk_Click
        End If

    End If
    
    Exit Sub
    
Erro_BotaoValidar_Click:

    Select Case gErr

        Case 126595
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_MAIOR_FALTA, gErr, gobjVenda.dFalta - StrParaDbl(TotalCheque.Caption))

        Case 133762
            Call Rotina_ErroECF(vbOKOnly, ERRO_BANCO_NAO_PREENCHIDO, gErr)
            
        Case 133763
            Call Rotina_ErroECF(vbOKOnly, ERRO_AGENCIA_NAO_PREENCHIDA, gErr)
            
        Case 133764
            Call Rotina_ErroECF(vbOKOnly, ERRO_CONTA_NAO_PREENCHIDA, gErr)
            
        Case 133765
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_PREENCHIDO, gErr)
            
        Case 133766
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case 133767
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
            
        Case 133768
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORMINIMO_CONDPAGTO, gErr, objAdmMeioPagtoCondPagto.dValorMinimo, Valor.Text)
            
        Case 133769
            
        Case 133770
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALIDAR_CHEQUE_SO_FUNCIONA_VENDA, gErr)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164179)

    End Select

    Exit Sub

End Sub

Private Sub CheckFixar_Click()
        Call WritePrivateProfileString(APLICACAO_CAIXA, "ChequeFixar", CStr(CheckFixar.Value), NOME_ARQUIVO_CAIXA)
End Sub

Private Sub Cliente_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub Conta_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub DataBomPara_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub DataBomParaGrid_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBomParaGrid, iAlterado)
End Sub


Private Sub Numero_Change()
    Status.Caption = STRING_CHEQUE_NAO_APROVADO
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99643
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99643
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164180)

    End Select

    Exit Sub
    
End Sub

Private Sub Banco_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Banco_Validate
    
    If Len(Trim(Banco.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Banco.Text)
        If lErro <> SUCESSO Then gError 99644
        
    End If
        
    Exit Sub
    
Erro_Banco_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99644
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164181)

    End Select

    Exit Sub
    
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Numero_Validate
    
    If Len(Trim(Numero.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 99647
        
    End If
        
    Exit Sub
    
Erro_Numero_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99647
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164182)

    End Select

    Exit Sub
    
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Cliente_Validate
    
    If Len(Trim(Cliente.ClipText)) = 0 Then Exit Sub
    
    Select Case Len(Trim(Cliente.ClipText))

    Case STRING_CPF 'CPF

        'Critica CPF
        lErro = Cpf_Critica(Cliente.Text)
        If lErro <> SUCESSO Then gError 99669
        
        Cliente.Format = "000\.000\.000-00; ; ; "
        Cliente.Text = Cliente.Text

    Case STRING_CGC 'CGC

        'Critica CGC
        lErro = Cgc_Critica(Cliente.Text)
        If lErro <> SUCESSO Then gError 99670

        Cliente.Format = "00\.000\.000\/0000-00; ; ; "
        Cliente.Text = Cliente.Text

    Case Else

        gError 99671

    End Select

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 99669, 99670

        Case 99671
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_TAMANHO_CGC_CPF, gErr)

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164183)

    End Select


    Exit Sub

End Sub

Private Sub DataBomPara_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DataBomPara_Validate
    
    If Len(Trim(DataBomPara.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataBomPara.Text)
        If lErro <> SUCESSO Then gError 99649
    
        '****Verifica se data bom para é menor que data atual se for Erro
        If StrParaDate(DataBomPara.Text) < Date Then gError 111431
    
    End If
        
    Exit Sub
    
Erro_DataBomPara_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99649
        
        Case 111431
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATABOMPARA_MENOR, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164184)

    End Select

    Exit Sub
    
End Sub

Private Sub SpinData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    lErro = Data_Up_Down_Click(DataBomPara, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 99651
    
    Exit Sub

Erro_SpinData_DownClick:

    Select Case gErr

        Case 99651

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164185)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    lErro = Data_Up_Down_Click(DataBomPara, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 99652

    Exit Sub

Erro_SpinData_UpClick:

    Select Case gErr

        Case 99652

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164186)

    End Select

    Exit Sub

End Sub

Private Sub GridCheques_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCheques, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridCheques, iAlterado)
    End If

End Sub

Private Sub GridCheques_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridCheques, iAlterado)

End Sub

Private Sub GridCheques_GotFocus()

    Call Grid_Recebe_Foco(objGridCheques)

End Sub

Private Sub GridCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCheques)
    
    Call Atualiza_Total

End Sub

Private Sub GridCheques_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCheques, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheques, iAlterado)
    End If
        
End Sub

Private Sub GridCheques_LeaveCell()

    Call Saida_Celula(objGridCheques)

End Sub

Private Sub GridCheques_LostFocus()

    Call Grid_Libera_Foco(objGridCheques)

End Sub

Private Sub GridCheques_RowColChange()

    Call Grid_RowColChange(objGridCheques)

End Sub

Private Sub GridCheques_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCheques)
        
End Sub

Private Sub GridCheques_Scroll()

    Call Grid_Scroll(objGridCheques)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99650

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99650
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164187)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a referência da tela
    Set gobjVenda = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not gobjVenda Is Nothing Then
    
    Select Case KeyCode
    
        Case vbKeyF2
            'Call BotaoLer_Click
    
        Case vbKeyF4
            'Call BotaoImprimir_Click
    
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
            Call BotaoOk_Click
    
        Case vbKeyEscape
            If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
            Call BotaoCancelar_Click
    
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
            Call BotaoIncluir_Click
    
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoValidar) Then Exit Sub
            Call BotaoValidar_Click
    
        Case vbKeyF8
            GridCheques.SetFocus
    
    End Select
        
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pagamentos em Cheque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoCheque"
    
End Function

Public Function objParent() As Object

    Set objParent = Parent
    
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
Private Sub ValorGrid_Change()

End Sub
