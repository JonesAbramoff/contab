VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ProgData2 
   Caption         =   "Programação das Datas"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleMode       =   0  'User
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Pagamentos Programados"
      Height          =   4320
      Left            =   135
      TabIndex        =   11
      Top             =   780
      Width           =   10080
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   1140
         TabIndex        =   23
         Top             =   1125
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   4080
         Left            =   7380
         TabIndex        =   15
         Top             =   195
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   15
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid Grid3 
         Height          =   4080
         Left            =   4980
         TabIndex        =   14
         Top             =   195
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   15
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   4080
         Left            =   2580
         TabIndex        =   13
         Top             =   195
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   15
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4080
         Left            =   180
         TabIndex        =   12
         Top             =   195
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   15
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados utilizados no cálculo"
      Height          =   735
      Left            =   135
      TabIndex        =   2
      Top             =   45
      Width           =   10080
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Limite dia:"
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
         Height          =   225
         Index           =   4
         Left            =   7215
         TabIndex        =   22
         Top             =   465
         Width           =   1140
      End
      Begin VB.Label LimiteDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8400
         TabIndex        =   21
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Limite:"
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
         Height          =   225
         Index           =   1
         Left            =   3720
         TabIndex        =   10
         Top             =   465
         Width           =   1590
      End
      Begin VB.Label DataLimite 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5385
         TabIndex        =   9
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Início:"
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
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   450
         Width           =   1935
      End
      Begin VB.Label DataInicio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2070
         TabIndex        =   7
         Top             =   405
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto. Assistência R$:"
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
         Index           =   24
         Left            =   165
         TabIndex        =   6
         Top             =   195
         Width           =   1860
      End
      Begin VB.Label SrvTotalAssistRS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2085
         TabIndex        =   5
         Top             =   135
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto. Seguro de Responsabilidade da Travel Ace R$:"
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
         Index           =   25
         Left            =   3825
         TabIndex        =   4
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label SrvTotalSegRS 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8415
         TabIndex        =   3
         Top             =   135
         Width           =   1560
      End
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9180
      Picture         =   "ProgData2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5130
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7755
      Picture         =   "ProgData2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5130
      Width           =   1005
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   7095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5265
      Width           =   225
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   5970
      TabIndex        =   17
      Top             =   5265
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Sugerida para o Pagamento:"
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
      Height          =   225
      Index           =   2
      Left            =   -60
      TabIndex        =   20
      Top             =   5280
      Width           =   3795
   End
   Begin VB.Label DataSugerida 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3825
      TabIndex        =   19
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   5460
      TabIndex        =   18
      Top             =   5295
      Width           =   480
   End
End
Attribute VB_Name = "ProgData2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjControle As Object

Dim iAlterado As Integer

Dim objGrid1 As AdmGrid
Dim objGrid2 As AdmGrid
Dim objGrid3 As AdmGrid
Dim objGrid4 As AdmGrid
Dim gsCodigo As String
Const iGrid_Data_Col As Integer = 0
Const iGrid_Valor_Col As Integer = 1

Private Sub BotaoCancela_Click()
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbCancel
    'Fecha a tela
    Unload Me
End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
Dim dtData As Date
    
On Error GoTo Erro_BotaoOK_Click
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
        
    Call DateParaMasked(gobjControle, StrParaDate(Data.Text))
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208700)

    End Select

    Exit Sub
    
End Sub

Private Sub Calcula_Data()
    
Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataLimite As Date
Dim dtDataProg As Date
Dim dValorPagarBD As Double
Dim dValorPagarAtual As Double
Dim objOcrCaso As New ClassTRVOcrCasos
Dim dValor As Double, dtData As Date, iIndice As Integer, iGrid As Integer
Dim sNaoUtil As String, objGrid As AdmGrid, dLimiteDia As Double, bAchou As Boolean, dtLimite As Date
Dim dtDataMenorValor As Date, dMenorValor As Double, dtDataAux As Date

On Error GoTo Erro_Calcula_Data

    bAchou = False

    objOcrCaso.sCodigo = gsCodigo
    
    lErro = CF("TRVOcrCasos_Le", objOcrCaso)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    dtDataProg = objOcrCaso.dtDataProgFinanc
    If dtDataProg <> DATA_NULA Then
        dValorPagarBD = objOcrCaso.dValorAutorizadoAssistRS + IIf(objOcrCaso.iAnteciparPagtoSeguro = MARCADO, objOcrCaso.dValorAutorizadoSeguroRS, objOcrCaso.dValorAutoSegRespTrvRS)
    End If
    dValorPagarAtual = StrParaDbl(SrvTotalAssistRS.Caption) + StrParaDbl(SrvTotalSegRS.Caption)
    
    dtDataIni = StrParaDate(DataInicio.Caption)
    dtLimite = StrParaDate(DataLimite.Caption)
    dLimiteDia = StrParaDbl(LimiteDia.Caption)
    dtDataMenorValor = DATA_NULA
    dMenorValor = 0
    dtData = DateAdd("d", -1, dtDataIni)

    'Para cada grid contendo 15 dias
    For iGrid = 1 To 4
        Select Case iGrid
            Case 1
                Set objGrid = objGrid1
            Case 2
                Set objGrid = objGrid2
            Case 3
                Set objGrid = objGrid3
            Case 4
                Set objGrid = objGrid4
        End Select
        For iIndice = 1 To 15
        
            dtData = DateAdd("d", 1, dtData)
        
            objGrid.objGrid.TextMatrix(iIndice, iGrid_Data_Col) = Format(dtData, "dd/mm/yyyy")
            
            'Busca o valor já liberado para essa data
            lErro = CF("TRVAssistVlrAuto_Le", dtData, dValor, sNaoUtil)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'Se é um dia útil
            If sNaoUtil = "" Then
                If dtData = dtDataProg Then dValor = dValor - dValorPagarBD 'Retira o próprio valor se já estiver gravado para essa data
                objGrid.objGrid.TextMatrix(iIndice, iGrid_Valor_Col) = Format(dValor, "STANDARD")
                
                'Se está dentro da data limite para pagamento e ainda não achou uma data apropriada
                If dtLimite >= dtData And Not bAchou Then
                
                    'Se não tem pagamentos para o dia, ou ele somado ao atual não chega ao limite diário
                    If (dValor = 0 Or dValor + dValorPagarAtual <= dLimiteDia) And dtLimite >= dtData Then
                        bAchou = True
                        Call DateParaMasked(Data, dtData)
                    End If
                    
                    'Dentre os dias não ótimos guarda o melhor que vai ser o com menor pagamento dentro do prazo limite
                    If (dMenorValor = 0 Or dMenorValor < dValor) And dtLimite >= dtData Then
                        dtDataMenorValor = dtData
                        dMenorValor = dValor
                    End If
                    
                End If
                
            Else
                'Informa se é sábado, domingo ou feriado
                objGrid.objGrid.TextMatrix(iIndice, iGrid_Valor_Col) = sNaoUtil
            End If
        
        Next
        
    Next
    
    'Se não achou nenhuma data dentro dos valores e prazos sugere a data com menor valor
    If Not bAchou Then
        If dtDataMenorValor = DATA_NULA Then
            Call DateParaMasked(Data, StrParaDate(DataLimite.Caption))
        Else
            Call DateParaMasked(Data, dtDataMenorValor)
        End If
    End If
    
    DataSugerida.Caption = Format(StrParaDate(Data.Text), "dd/mm/yyyy")

    Exit Sub
    
Erro_Calcula_Data:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208702)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(ByVal objControle As Object, ByVal sCodigo As String, ByVal dtDataDocRec As Date, ByVal dtDataLimite As Date, ByVal dValorAssist As Double, ByVal dValorSeg As Double, ByVal iAntecSeg As Integer, ByVal dValorSegTRV As Double) As Long

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Trata_Parametros

    Set gobjControle = objControle
    gsCodigo = sCodigo
    
    SrvTotalAssistRS.Caption = Format(dValorAssist, "STANDARD")
    DataLimite.Caption = Format(dtDataLimite, "dd/mm/yyyy")
    
    If iAntecSeg = MARCADO Then
        'AntecPagto.Value = vbChecked
        'Se está antecipando a Travel paga o seguro integral
        SrvTotalSegRS.Caption = Format(dValorSeg, "STANDARD")
    Else
        'AntecPagto.Value = vbUnchecked
        'Se não antecipar paga somente o que for de sua responsabilidade
        SrvTotalSegRS.Caption = Format(dValorSegTRV, "STANDARD")
    End If
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_ASSISTENCIA_LIMITE_DIARIO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    LimiteDia.Caption = Format(StrParaDbl(sConteudo), "STANDARD")
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_ASSISTENCIA_DATA_INICIO_LIB, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Date < StrParaDate(sConteudo) Then
        DataInicio.Caption = Format(StrParaDate(sConteudo), "dd/mm/yyyy")
    Else
        DataInicio.Caption = Format(Date, "dd/mm/yyyy")
    End If

    Set objGrid1 = New AdmGrid
    Set objGrid2 = New AdmGrid
    Set objGrid3 = New AdmGrid
    Set objGrid4 = New AdmGrid
    
    objGrid1.objGrid = Grid1
    objGrid2.objGrid = Grid2
    objGrid3.objGrid = Grid3
    objGrid4.objGrid = Grid4

    lErro = Inicializa_Grid(objGrid1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid(objGrid2)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid(objGrid3)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid(objGrid4)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_Data

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208703)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set gobjControle = Nothing
    Set objGrid1 = Nothing
    Set objGrid2 = Nothing
    Set objGrid3 = Nothing
    Set objGrid4 = Nothing
End Sub

Private Function Inicializa_Grid(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Valor.Name)
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 1000

    'Linhas do grid
    objGridInt.objGrid.Rows = 15 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid = SUCESSO

    Exit Function

Erro_Inicializa_Grid:

    Inicializa_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192096)

    End Select

    Exit Function

End Function

Private Sub Grid1_DblClick()
    Call Trata_Click_Grid(objGrid1)
End Sub

Private Sub Grid2_DblClick()
    Call Trata_Click_Grid(objGrid2)
End Sub

Private Sub Grid3_DblClick()
    Call Trata_Click_Grid(objGrid3)
End Sub

Private Sub Grid4_DblClick()
    Call Trata_Click_Grid(objGrid4)
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190610)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190612)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190614)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Click_Grid(ByVal objGrid As AdmGrid)
Dim dtData As Date
    If objGrid.objGrid.Row <> 0 Then
        dtData = StrParaDate(objGrid.objGrid.TextMatrix(objGrid.objGrid.Row, iGrid_Data_Col))
        Call DateParaMasked(Data, dtData)
    End If
End Sub
