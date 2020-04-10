VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ProcessaArqRetCobr2 
   Caption         =   "Processamento dos Títulos em Cobrança"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   30
      TabIndex        =   32
      Top             =   390
      Width           =   9255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CheckBox GravaCriticas 
      Caption         =   "Grava relatório de críticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6495
      TabIndex        =   23
      Top             =   3780
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   645
      Left            =   2115
      TabIndex        =   19
      Top             =   4830
      Width           =   6840
      Begin VB.CommandButton BotaoInterromper 
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
         Height          =   585
         Left            =   2790
         Picture         =   "ProcessaArqRetCobr2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton BotaoProcessar 
         Caption         =   "Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   885
         Picture         =   "ProcessaArqRetCobr2.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton BotaoFechar 
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4695
         Picture         =   "ProcessaArqRetCobr2.frx":025C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.CommandButton BotaoConsultarTitRec 
      Caption         =   "Consultar Título"
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
      Left            =   7290
      TabIndex        =   17
      Top             =   2775
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   4410
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox RetSeuNumero 
      Height          =   225
      Left            =   390
      TabIndex        =   6
      Top             =   1635
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   1140
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   2011
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSMask.MaskEdBox RetNossoNumero 
      Height          =   225
      Left            =   1485
      TabIndex        =   9
      Top             =   1635
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetDataVenc 
      Height          =   225
      Left            =   2580
      TabIndex        =   10
      Top             =   1635
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetValor 
      Height          =   225
      Left            =   3570
      TabIndex        =   11
      Top             =   1635
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetCritica 
      Height          =   225
      Left            =   1230
      TabIndex        =   12
      Top             =   2160
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetValorRec 
      Height          =   225
      Left            =   4440
      TabIndex        =   13
      Top             =   1635
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetValorJuros 
      Height          =   225
      Left            =   5310
      TabIndex        =   14
      Top             =   1635
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetValorDesc 
      Height          =   225
      Left            =   6180
      TabIndex        =   15
      Top             =   1635
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RetValorTarifa 
      Height          =   225
      Left            =   4860
      TabIndex        =   16
      Top             =   2280
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Custas:"
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
      Left            =   6750
      TabIndex        =   31
      Top             =   4095
      Width           =   645
   End
   Begin VB.Label Custas 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7470
      TabIndex        =   30
      Top             =   4065
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Tarifas:"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   4095
      Width           =   660
   End
   Begin VB.Label Tarifa 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5400
      TabIndex        =   28
      Top             =   4065
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Baixado:"
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
      Left            =   2715
      TabIndex        =   27
      Top             =   4110
      Width           =   750
   End
   Begin VB.Label Baixado 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3480
      TabIndex        =   26
      Top             =   4065
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Recebido:"
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
      Left            =   645
      TabIndex        =   25
      Top             =   4095
      Width           =   885
   End
   Begin VB.Label Recebido 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1575
      TabIndex        =   24
      Top             =   4065
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ocorrências:"
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
      Left            =   105
      TabIndex        =   18
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Ocr 
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   75
      TabIndex        =   8
      Top             =   3150
      Width           =   9135
   End
   Begin VB.Label TotalTitulos 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1575
      TabIndex        =   5
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5400
      TabIndex        =   4
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Processamento dos Títulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2955
      TabIndex        =   3
      Top             =   120
      Width           =   2850
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Titulos Processados:"
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
      Left            =   3555
      TabIndex        =   2
      Top             =   3750
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Títulos:"
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
      Left            =   90
      TabIndex        =   1
      Top             =   3750
      Width           =   1440
   End
End
Attribute VB_Name = "ProcessaArqRetCobr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sNomeArqParam As String

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjCobrancaEletronica As ClassCobrancaEletronica

Dim giCancelaBatch As Integer
Dim giExecutando As Integer ' 0: nao está executando, 1: em andamento
Dim iAlterado As Integer

Dim objGridItens As AdmGrid
Dim iGrid_RetSeuNumero_Col As Integer
Dim iGrid_RetNossoNumero_Col As Integer
Dim iGrid_RetDataVenc_Col As Integer
Dim iGrid_RetValor_Col As Integer
Dim iGrid_RetValorRec_Col As Integer
Dim iGrid_RetValorJuros_Col As Integer
Dim iGrid_RetValorDesc_Col As Integer
Dim iGrid_RetValorTarifa_Col As Integer
Dim iGrid_RetCritica_Col As Integer

Public dValorRecebido As Double
Public dValorTarifas As Double
Public dValorCustas As Double
Public dValorBaixado As Double

Public bTeste As Boolean
Public iGravaCriticas As Integer
Public colRetCobrErros As New Collection
Public colcolTiposDetRetCobr As New Collection
Public colTiposMovRetCobr As New Collection

Private Sub BotaoConsultarTitRec_Click()
    Call Mostra_Titulo(GridItens.Row)
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    FrameBotoes.Enabled = False
    
    Call Form_Load2

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165295)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load2()

Dim lErro As Long, X As New ClassCNABCobrRet
Dim lQuantRegistros As Long
Dim objDetRetCobr As ClassDetRetCobr
Dim iLinha As Integer, sTipoCritica As String
Dim objTiposDet As ClassTiposDetRetCobr
Dim colTiposDet As Collection
Dim vValor As Variant, bAchou As Boolean
Dim colTiposMov As New Collection
Dim colcolTiposDet As New Collection
Dim iIndice As Integer
Dim objTiposMov As ClassTiposMovRetCobr

On Error GoTo Erro_Form_Load2
    
    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO
    
    Set objGridItens = New AdmGrid
    bTeste = True
    Set gobjCobrancaEletronica.objTelaAtualizacao = Me
    
    lErro = X.CobrancaEletronica_ObtemNumTitulos(gobjCobrancaEletronica, lQuantRegistros)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_Itens(objGridItens, lQuantRegistros)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Passa para a tela os dados dos títulos
    TotalTitulos.Caption = CStr(lQuantRegistros)
    TitulosProcessados.Caption = "0"

''    ProgressBar1.Min = 0
''    ProgressBar1.Max = 100
''
'''    bTeste = True
'''    Set gobjCobrancaEletronica.objTelaAtualizacao = Me
'''
'''    lErro = Inicializa_Grid_Itens(objGridItens)
'''    If lErro <> SUCESSO Then gError 81540
''
''    FrameBotoes.Enabled = False
    
    Frame1.Visible = False
    lErro = CF("Processar_ArquivoRetorno_Cobranca", gobjCobrancaEletronica)
    FrameBotoes.Enabled = True
    If lErro <> SUCESSO And lErro <> 59190 Then gError ERRO_SEM_MENSAGEM
    
    Set objTiposMov = New ClassTiposMovRetCobr
    
    objTiposMov.lBanco = gobjCobrancaEletronica.objCCI.iCodBanco
    
    lErro = CF("TiposMovRetCobr_Le", objTiposMov, colTiposMovRetCobr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Set objTiposMov = New ClassTiposMovRetCobr
    
    objTiposMov.lBanco = 0
    
    lErro = CF("TiposMovRetCobr_Le", objTiposMov, colTiposMovRetCobr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    iLinha = 0
    For Each objDetRetCobr In colRetCobrErros
        iLinha = iLinha + 1
                
        Select Case objDetRetCobr.iTipoCritica
            Case 1
                sTipoCritica = "Não Encontrou"
            Case 2
                sTipoCritica = "Encontrou + de 1"
            Case 3
                sTipoCritica = "Entrada Rejeitada"
            Case 6
                sTipoCritica = "Liquidação"
            Case 9
                sTipoCritica = "Outras Baixas"
            Case 12
                sTipoCritica = "Tarifas"
            Case 25
                sTipoCritica = "Baixa por protesto"
            Case 33
                sTipoCritica = "Custas"
            Case 100
                sTipoCritica = "Baixa Parcial"
            Case 50
                sTipoCritica = "Já baixada"
            Case 51
                sTipoCritica = "Encontrou-Dif.Vlr"
            Case 52
                sTipoCritica = "Encontrou-Dif.Venc"
            Case RETCOBR_CRITICA_PAGA_DUPLIC
                sTipoCritica = "Paga em duplicidade"
            Case Else
                sTipoCritica = "Outros"
        End Select
        
        GridItens.TextMatrix(iLinha, iGrid_RetSeuNumero_Col) = Trim(right(Trim(objDetRetCobr.sSeuNumero), 10))
        GridItens.TextMatrix(iLinha, iGrid_RetNossoNumero_Col) = objDetRetCobr.sNossoNumero
        If objDetRetCobr.dtDataVencimento <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_RetDataVenc_Col) = Format(objDetRetCobr.dtDataVencimento, "dd/mm/yyyy")
        GridItens.TextMatrix(iLinha, iGrid_RetCritica_Col) = sTipoCritica
        GridItens.TextMatrix(iLinha, iGrid_RetValor_Col) = Format(objDetRetCobr.dValorTitulo, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_RetValorDesc_Col) = Format(objDetRetCobr.dValorDesconto, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_RetValorJuros_Col) = Format(objDetRetCobr.dValorJuros, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_RetValorRec_Col) = Format(objDetRetCobr.dValorRecebido, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_RetValorTarifa_Col) = Format(objDetRetCobr.dValorTarifa, "STANDARD")
            
        Set colTiposDet = New Collection
        
        If objDetRetCobr.iCodOcorrencia <> 0 Then
            bAchou = False
            iIndice = 0
            For Each vValor In colTiposMov
                iIndice = iIndice + 1
                If objDetRetCobr.iCodOcorrencia = vValor Then
                    bAchou = True
                    Set colTiposDet = colcolTiposDet.Item(iIndice)
                    Exit For
                End If
            Next
            If Not bAchou Then
                colTiposMov.Add objDetRetCobr.iCodOcorrencia
                
                Set objTiposDet = New ClassTiposDetRetCobr
                Set colTiposDet = New Collection
                
                objTiposDet.lBanco = gobjCobrancaEletronica.objCCI.iCodBanco
                
                If objTiposDet.lBanco = 0 Then
                    If Not (gobjCobrancaEletronica.objCobrador Is Nothing) Then
                        objTiposDet.lBanco = gobjCobrancaEletronica.objCobrador.iCodBanco
                    End If
                End If
                
                objTiposDet.iCodigoMovto = objDetRetCobr.iCodOcorrencia
                
                lErro = CF("TiposDetRetCobr_Le", objTiposDet, colTiposDet)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Set objTiposDet = New ClassTiposDetRetCobr
                
                objTiposDet.lBanco = 0
                objTiposDet.iCodigoMovto = objDetRetCobr.iCodOcorrencia
                
                lErro = CF("TiposDetRetCobr_Le", objTiposDet, colTiposDet)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                colcolTiposDet.Add colTiposDet
            End If
        End If
    
        colcolTiposDetRetCobr.Add colTiposDet
    
    Next
    objGridItens.iLinhasExistentes = iLinha
    
    Tarifa.Caption = Format(dValorTarifas, "STANDARD")
    Recebido.Caption = Format(dValorRecebido, "STANDARD")
    Baixado.Caption = Format(dValorBaixado, "STANDARD")
    Custas.Caption = Format(dValorCustas, "STANDARD")
    
    Call Mostra_Ocr(1)

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load2:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
'        Case 81540
'            Call Rotina_ErrosBatch
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165295)
    
    End Select
    
''    Call Rotina_ErrosBatch2("Processamento do Retorno da Cobrança")
    
    giCancelaBatch = CANCELA_BATCH
    BotaoProcessar.Enabled = False
    
    Exit Sub

End Sub

Function Trata_Parametros(objCobrancaEletronica As ClassCobrancaEletronica) As Long

    Trata_Parametros = SUCESSO

End Function



Private Sub BotaoFechar_Click()
   
    If giExecutando = ESTADO_ANDAMENTO Then
        giCancelaBatch = CANCELA_BATCH
        BotaoFechar.Enabled = False
        Exit Sub
    End If

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoProcessar_Click()

Dim lErro As Long, sErro As String

On Error GoTo Erro_BotaoProcessar_Click

    BotaoProcessar.Enabled = False

    BotaoInterromper.Enabled = True
    
    bTeste = False
    
    If giCancelaBatch <> CANCELA_BATCH Then

        giExecutando = ESTADO_ANDAMENTO
        
        If GravaCriticas.Value = vbChecked Then
            iGravaCriticas = MARCADO
        Else
            iGravaCriticas = DESMARCADO
        End If
        
        Set gobjCobrancaEletronica.objTelaAtualizacao = Me
            
        lErro = CF("Processar_ArquivoRetorno_Cobranca", gobjCobrancaEletronica)
                
        giExecutando = ESTADO_PARADO

        BotaoInterromper.Enabled = False

        If lErro <> SUCESSO And lErro <> 59190 Then Error 51680
        If lErro = 59190 Then Error 51679 'interrompeu

        'Fecha a tela
        Unload Me
        
    End If

    Exit Sub

Erro_BotaoProcessar_Click:

'    sErro = "Houve algum tipo de erro. Verifique o arquivo de log de erros configurado em \windows\adm100.ini ."
'    Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")
    
    Select Case Err

        Case 51679
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")

        Case 51680

        Case Else
            'Call Rotina_ErrosBatch
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165296)

    End Select

'''    If giCancelaBatch <> CANCELA_BATCH Then Call Rotina_ErrosBatch2("Processamento de Retorno da Cobrança")
    
    Unload Me
    
    Exit Sub

End Sub

Public Function Mostra_Evolucao(iCancela As Integer, iNumProc As Integer) As Long
'Mostra a evolução dos borderos processados

Dim lErro As Long
Dim iEventos As Integer
Dim iProcessados As Integer
Dim iTotal As Integer

On Error GoTo Erro_Mostra_Evolucao

    iEventos = DoEvents()
    
    If Not bTeste Then
    
        If giCancelaBatch = CANCELA_BATCH Then
    
            iCancela = CANCELA_BATCH
            giExecutando = ESTADO_PARADO
    
        Else
            'atualiza dados da tela ( registros atualizados e a barra )
    
            iProcessados = CInt(TitulosProcessados.Caption)
            iTotal = CInt(TotalTitulos.Caption)
    
            iProcessados = iProcessados + iNumProc
            TitulosProcessados.Caption = CStr(iProcessados)
    
            ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)
    
            giExecutando = ESTADO_ANDAMENTO
    
        End If
        
    End If

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165297)

    End Select

    giCancelaBatch = CANCELA_BATCH

    Exit Function

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If giExecutando = ESTADO_ANDAMENTO Then
        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If


End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set gobjCobrancaEletronica.objTelaAtualizacao = Nothing
    Set gobjCobrancaEletronica = Nothing
    Set objGridItens = Nothing
    Set colRetCobrErros = Nothing
    Set colcolTiposDetRetCobr = Nothing
    Set colTiposMovRetCobr = Nothing
End Sub

Private Sub BotaoInterromper_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    
    End If
    
    'Fecha a tela
    Unload Me

End Sub

Private Sub Timer1_Timer()

Dim lErro As Long
Dim sErro As String

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

'''*** Para depurar, usando o Batch como .dll, o trecho abaixo deve estar comentado
''    lErro = Sistema_Abrir_Batch(sNomeArqParam)
''    If lErro <> SUCESSO Then gError 189875
'''***
''
''    Set gcolModulo = New AdmColModulo
''
''    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
''    If lErro <> SUCESSO Then gError 189876
''
''    lErro = CF("Retorna_ColFiliais")
''    If lErro <> SUCESSO Then gError 189877
''
''    GL_lUltimoErro = SUCESSO
    
''    Call Form_Load2
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case gErr

        Case 189875 To 189879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189880)

    End Select

''    If giCancelaBatch <> CANCELA_BATCH Then
''        Call Rotina_ErrosBatch2("Processamento de Retorno de Cobrança")
''    End If

    giCancelaBatch = CANCELA_BATCH

    Exit Sub

End Sub

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid, ByVal lQuantRegistros As Long) As Long
'Inicializa o Grid de Alocações

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Seu Num.")
    objGridInt.colColuna.Add ("Nosso Número")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Tarifa")
    objGridInt.colColuna.Add ("Crítica")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (RetSeuNumero.Name)
    objGridInt.colCampo.Add (RetNossoNumero.Name)
    objGridInt.colCampo.Add (RetDataVenc.Name)
    objGridInt.colCampo.Add (RetValor.Name)
    objGridInt.colCampo.Add (RetValorRec.Name)
    objGridInt.colCampo.Add (RetValorJuros.Name)
    objGridInt.colCampo.Add (RetValorDesc.Name)
    objGridInt.colCampo.Add (RetValorTarifa.Name)
    objGridInt.colCampo.Add (RetCritica.Name)

    'Colunas da Grid
    iGrid_RetSeuNumero_Col = 1
    iGrid_RetNossoNumero_Col = 2
    iGrid_RetDataVenc_Col = 3
    iGrid_RetValor_Col = 4
    iGrid_RetValorRec_Col = 5
    iGrid_RetValorJuros_Col = 6
    iGrid_RetValorDesc_Col = 7
    iGrid_RetValorTarifa_Col = 8
    iGrid_RetCritica_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    If lQuantRegistros > 1000 Then
        objGridInt.objGrid.Rows = lQuantRegistros + 1
    Else
        objGridInt.objGrid.Rows = 1000 + 1
    End If
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Public Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()
    'Call Saida_Celula(objGridItens)
End Sub

Public Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Public Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
    Call Mostra_Ocr(GridItens.Row)
End Sub

Public Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub Mostra_Ocr(ByVal iLinha As Integer)
'TiposDetRetCobr_Le

Dim lErro As Long
Dim sOcr As String
Dim objDetRetCobr As ClassDetRetCobr
Dim objTiposMov As ClassTiposMovRetCobr
Dim objTiposDet As ClassTiposDetRetCobr
Dim colTiposDet As Collection
Dim bAchou As Boolean

On Error GoTo Erro_Mostra_Ocr

    If Not (colRetCobrErros Is Nothing) Then
            
        If iLinha > 0 And iLinha <= colRetCobrErros.Count Then
        
            Set objDetRetCobr = colRetCobrErros.Item(iLinha)
            
            If objDetRetCobr.iCodOcorrencia <> 0 Then
            
                bAchou = False
                For Each objTiposMov In colTiposMovRetCobr
                    If objTiposMov.iCodigoMovto = objDetRetCobr.iCodOcorrencia Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    sOcr = "O movimento de código " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator." & vbNewLine
                Else
                    sOcr = objTiposMov.sDescricao & vbNewLine
                    Set colTiposDet = colcolTiposDetRetCobr.Item(iLinha)
                    
                    If objDetRetCobr.iCodOcorrencia1 <> 0 Then
                    
                        bAchou = False
                        For Each objTiposDet In colTiposDet
                            If objTiposDet.iCodigoDetalhe = objDetRetCobr.iCodOcorrencia1 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            sOcr = sOcr & "A ocorrência 1 de código " & CStr(objDetRetCobr.iCodOcorrencia1) & " e movimento " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator."
                        Else
                            sOcr = sOcr & objTiposDet.sDescricao
                        End If
                    End If
                    
                    If objDetRetCobr.iCodOcorrencia2 <> 0 Then
                    
                        bAchou = False
                        For Each objTiposDet In colTiposDet
                            If objTiposDet.iCodigoDetalhe = objDetRetCobr.iCodOcorrencia2 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            sOcr = sOcr & ";" & "A ocorrência 2 de código " & CStr(objDetRetCobr.iCodOcorrencia1) & " e movimento " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator."
                        Else
                            sOcr = sOcr & ";" & objTiposDet.sDescricao
                        End If
                    End If
                    
                    If objDetRetCobr.iCodOcorrencia3 <> 0 Then
                    
                        bAchou = False
                        For Each objTiposDet In colTiposDet
                            If objTiposDet.iCodigoDetalhe = objDetRetCobr.iCodOcorrencia3 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            sOcr = sOcr & ";" & "A ocorrência 3 de código " & CStr(objDetRetCobr.iCodOcorrencia1) & " e movimento " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator."
                        Else
                            sOcr = sOcr & ";" & objTiposDet.sDescricao
                        End If
                    End If
                    
                    If objDetRetCobr.iCodOcorrencia4 <> 0 Then
                    
                        bAchou = False
                        For Each objTiposDet In colTiposDet
                            If objTiposDet.iCodigoDetalhe = objDetRetCobr.iCodOcorrencia4 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            sOcr = sOcr & ";" & "A ocorrência 4 de código " & CStr(objDetRetCobr.iCodOcorrencia1) & " e movimento " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator."
                        Else
                            sOcr = sOcr & ";" & objTiposDet.sDescricao
                        End If
                    End If
                    
                    If objDetRetCobr.iCodOcorrencia5 <> 0 Then
                    
                        bAchou = False
                        For Each objTiposDet In colTiposDet
                            If objTiposDet.iCodigoDetalhe = objDetRetCobr.iCodOcorrencia5 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            sOcr = sOcr & ";" & "A ocorrência 5 de código " & CStr(objDetRetCobr.iCodOcorrencia1) & " e movimento " & CStr(objDetRetCobr.iCodOcorrencia) & " não tem a descrição cadastrada no Corporator."
                        Else
                            sOcr = sOcr & ";" & objTiposDet.sDescricao
                        End If
                    End If
                    
                End If
                
            End If
        
        End If
    
    End If
    
    Ocr.Caption = sOcr

    Exit Sub

Erro_Mostra_Ocr:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165297)

    End Select

    Exit Sub
    
End Sub

Private Sub Mostra_Titulo(ByVal iLinha As Integer)

Dim lErro As Long
Dim objDetRetCobr As ClassDetRetCobr
Dim objParcRec As New ClassParcelaReceber
Dim objTitRec As New ClassTituloReceber

On Error GoTo Erro_Mostra_Titulo

    If Not (colRetCobrErros Is Nothing) Then
    
        If iLinha > 0 And iLinha <= colRetCobrErros.Count Then
        
            Set objDetRetCobr = colRetCobrErros.Item(iLinha)
            
            If objDetRetCobr.lNumIntParc = 0 Then gError 206447
            
            objParcRec.lNumIntDoc = objDetRetCobr.lNumIntParc
            
            lErro = CF("ParcelaReceber_Le", objParcRec)
            If lErro <> SUCESSO And lErro = 19147 Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then
            
                lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
                If lErro <> SUCESSO And lErro = 58559 Then gError ERRO_SEM_MENSAGEM
            
            End If
            
            objTitRec.lNumIntDoc = objParcRec.lNumIntTitulo
            
            Call Chama_Tela("TituloReceber_Consulta", objTitRec)
            
        End If

    End If

    Exit Sub

Erro_Mostra_Titulo:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206447
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELASREC_NAO_EXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165297)

    End Select

    Exit Sub

End Sub
