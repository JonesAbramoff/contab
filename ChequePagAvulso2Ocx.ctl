VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequePagAvulso2Ocx 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ScaleHeight     =   4320
   ScaleWidth      =   8850
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   4635
      TabIndex        =   37
      Top             =   2190
      Width           =   1305
      _ExtentX        =   2302
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
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   3053
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3660
      Width           =   2745
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1140
         Picture         =   "ChequePagAvulso2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   120
         Picture         =   "ChequePagAvulso2Ocx.ctx":0792
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "ChequePagAvulso2Ocx.ctx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Tipo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3900
      TabIndex        =   19
      Top             =   1965
      Width           =   795
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Height          =   690
      Left            =   120
      Picture         =   "ChequePagAvulso2Ocx.ctx":106E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3525
      Width           =   1650
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionado"
      Height          =   915
      Left            =   6405
      TabIndex        =   21
      Top             =   2595
      Width           =   2280
      Begin VB.Label TotalTitulosSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   750
         TabIndex        =   23
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label QtdTitulosSelecionados 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   750
         TabIndex        =   24
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
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
         Index           =   1
         Left            =   150
         TabIndex        =   26
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Height          =   915
      Left            =   3885
      TabIndex        =   22
      Top             =   2595
      Width           =   2295
      Begin VB.Label TotalTitulos 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   810
         TabIndex        =   27
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label QtdTitulos 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   810
         TabIndex        =   28
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
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
         Left            =   150
         TabIndex        =   29
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todos"
      Height          =   585
      Left            =   120
      Picture         =   "ChequePagAvulso2Ocx.ctx":3F84
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1650
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todos"
      Height          =   585
      Left            =   1995
      Picture         =   "ChequePagAvulso2Ocx.ctx":4F9E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   1650
   End
   Begin VB.CheckBox CheckPago 
      Height          =   255
      Left            =   4995
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin MSMask.MaskEdBox TipoCobranca 
      Height          =   225
      Left            =   6270
      TabIndex        =   11
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin MSMask.MaskEdBox Portador 
      Height          =   225
      Left            =   4725
      TabIndex        =   10
      Top             =   1935
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Juros 
      Height          =   225
      Left            =   4650
      TabIndex        =   7
      Top             =   915
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Desconto 
      Height          =   225
      Left            =   4710
      TabIndex        =   9
      Top             =   1545
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Multa 
      Height          =   225
      Left            =   4695
      TabIndex        =   8
      Top             =   1170
      Width           =   720
      _ExtentX        =   1270
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
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   3225
      TabIndex        =   3
      Top             =   555
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Titulo 
      Height          =   225
      Left            =   2385
      TabIndex        =   2
      Top             =   540
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSMask.MaskEdBox Filial 
      Height          =   225
      Left            =   1695
      TabIndex        =   1
      Top             =   525
      Width           =   630
      _ExtentX        =   1111
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
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox DataVencto 
      Height          =   225
      Left            =   5715
      TabIndex        =   6
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   3975
      TabIndex        =   4
      Top             =   525
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridChequePagAvulso2 
      Height          =   1995
      Left            =   120
      TabIndex        =   12
      Top             =   525
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   3519
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
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
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   31
      Top             =   135
      Width           =   570
   End
   Begin VB.Label LabelConta 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   870
      TabIndex        =   32
      Top             =   105
      Width           =   1950
   End
   Begin VB.Label LabelFilial 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6705
      TabIndex        =   33
      Top             =   105
      Width           =   1950
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Filial:"
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
      Left            =   6165
      TabIndex        =   34
      Top             =   135
      Width           =   465
   End
   Begin VB.Label LabelFornecedor 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4050
      TabIndex        =   35
      Top             =   105
      Width           =   1950
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
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
      Left            =   2955
      TabIndex        =   36
      Top             =   135
      Width           =   1035
   End
End
Attribute VB_Name = "ChequePagAvulso2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Private gobjChequesPagAvulso As ClassChequesPagAvulso
Private gcolTiposDeCobranca As New AdmColCodigoNome
 
Dim objGrid As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Titulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Pago_Col As Integer
Dim iGrid_DataVencto_Col As Integer
Dim iGrid_Juros_Col As Integer
Dim iGrid_Multa_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_Portador_Col As Integer
Dim iGrid_TipoCobranca_Col As Integer

Private Sub BotaoFechar_Click()
    
    'Fecha a tela
    Unload Me
    
End Sub

Private Sub BotaoDesmarcar_Click()
'Desmarca todas as parcelas do Grid

Dim iLinha As Integer
Dim iNumTitulos As Integer
Dim objInfoParcPag As ClassInfoParcPag

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela a parcela em questão
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_NAO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
        
        'Desmarca no Obj a parcela em questão
        objInfoParcPag.iSeqCheque = 0
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Limpa na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(0)
    TotalTitulosSelecionados.Caption = CStr(Format(0, "Standard"))

End Sub

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objInfoParcPag As New ClassInfoParcPag
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridChequePagAvulso2.Row = 0 Then Error 60496
        
    'Se foi selecionada uma linha que está preenchida
    If GridChequePagAvulso2.Row <= objGrid.iLinhasExistentes Then
        
        Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(GridChequePagAvulso2.Row)
               
        objParcelaPagar.lNumIntDoc = objInfoParcPag.lNumIntParc
        
        'Le o NumInterno do Titulo para passar no objTituloPag
        lErro = CF("ParcelaPagar_Le", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 60479 Then Error 60497
        
        'Se não encontrou a Parcela --> ERRO
        If lErro = 60479 Then Error 60498
        
        objTituloPagar.lNumIntDoc = objParcelaPagar.lNumIntTitulo
        
        'Le os Dados do Titulo
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 60499
        
        If lErro = 18372 Then Error 61000
        
        Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case Err
    
        Case 60496
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
         
        Case 60497, 60499 'Tratado na rotina chamada
        
        Case 60498
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", Err)
        
        Case 61000
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144482)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcar_Click()
'Marca todas as parcelas do Grid

Dim iLinha As Integer
Dim dTotalTitulosSelecionados As Double
Dim iNumTitulosSelecionados As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim dTotalJuros As Double
Dim dTotalMultas As Double
Dim dTotalDescontos As Double

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela a parcela em questão
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
        
        'Marca no Obj a parcela em questão
        objInfoParcPag.iSeqCheque = 1
                
        'Faz o somatório da Qtd e do Total das parcelas selecionadas
        dTotalTitulosSelecionados = dTotalTitulosSelecionados + CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Valor_Col))
        iNumTitulosSelecionados = iNumTitulosSelecionados + 1
        dTotalJuros = dTotalJuros + CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Juros_Col))
        dTotalMultas = dTotalMultas + CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Multa_Col))
        dTotalDescontos = dTotalDescontos + CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Desconto_Col))
        
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Atualiza na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(iNumTitulosSelecionados)
    TotalTitulosSelecionados.Caption = CStr(Format(dTotalTitulosSelecionados + dTotalJuros + dTotalMultas - dTotalDescontos, "Standard"))
    
End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iPortador As Integer, lFornecedor As Long, iFilialForn As Integer, iTipoCobranca As Integer
Dim iPortadorAnterior As Integer, objInfoParcPag As ClassInfoParcPag
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoSeguir_Click

    'Ao menos uma parcela tem que estar marcada
    If CInt(QtdTitulosSelecionados.Caption) = 0 Then gError 43057
    
    gobjChequesPagAvulso.iQtdeParcelasSelecionadas = CInt(QtdTitulosSelecionados.Caption)
    
    iPortadorAnterior = -1
    iTipoCobranca = -1
    lFornecedor = -1
    iFilialForn = -1
    
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        If (GridChequePagAvulso2.TextMatrix(iIndice, iGrid_Pago_Col)) = PAGO_CHECADO Then
        
            Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag(iIndice)
            
            lErro = CF("TituloPagar_Verifica_Adiantamento", objInfoParcPag.lFornecedor, objInfoParcPag.iFilialForn)
            If lErro <> SUCESSO Then gError 59472
    
            'Verifica se outro Portador foi preenchido
            iPortador = Codigo_Extrai(GridChequePagAvulso2.TextMatrix(iIndice, iGrid_Portador_Col))
            If iPortadorAnterior = -1 Then
                iPortadorAnterior = iPortador
                iTipoCobranca = objInfoParcPag.iTipoCobranca
                lFornecedor = objInfoParcPag.lFornecedor
                iFilialForn = objInfoParcPag.iFilialForn
            Else
                'Se for portadores diferentes --> erro
                If iPortador <> iPortadorAnterior Then gError 43059
                If objInfoParcPag.iTipoCobranca <> iTipoCobranca Then gError 41591
                If (objInfoParcPag.lFornecedor <> lFornecedor Or objInfoParcPag.iFilialForn <> iFilialForn) And (objInfoParcPag.iTipoCobranca <> TIPO_COBRANCA_BANCARIA Or iPortador = 0) Then gError 41592
            End If
            
            'Titulos com pagto após o vcto nao deveriam ter desconto
            If gobjChequesPagAvulso.dtEmissao > objInfoParcPag.dtDataVencimento And objInfoParcPag.dValorDesconto > 0 Then
            
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PAGTO_ATRASO_DESC", iIndice)
                If vbMsgRes = vbYes Then gError 59077
            
            End If
            
            'Titulos com pagto até o vcto nao deveriam ter juros ou multa
            If gobjChequesPagAvulso.dtEmissao <= objInfoParcPag.dtDataVencimento And (objInfoParcPag.dValorJuros > 0 Or objInfoParcPag.dValorMulta > 0) Then
                
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PAGTO_EM_DIA_MULTA", iIndice)
                If vbMsgRes = vbYes Then gError 59078
            
            End If
        
        End If
        
    Next
    
    'Lê as parcelas selecionadas no Grid e passa os dados para o cheque a ser emitido
    lErro = CF("ChequesPagAvulso_ChequesSelecionados", gobjChequesPagAvulso)
    If lErro <> SUCESSO Then gError 15813
        
    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequePagAvulso3", gobjChequesPagAvulso)
       
    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case gErr

        Case 59077, 59078 'desistiu por que descobriu que tinha digitado valor errado
        
        Case 15813, 59472

        Case 43057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELA_SELECIONADA", Err)
            
        Case 43059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADORES_DIFERENTES", Err)
            
        Case 41591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_TIPO_COBR_DIFERENTE", Err)
        
        Case 41592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_FORN_DIF_NAO_COBRBANC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144483)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVoltar_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iPortador As Integer, lFornecedor As Long, iFilialForn As Integer, iTipoCobranca As Integer
Dim iPortadorAnterior As Integer, objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_BotaoVoltar_Click

    gobjChequesPagAvulso.iQtdeParcelasSelecionadas = CInt(QtdTitulosSelecionados.Caption)
    
    iPortadorAnterior = -1
    iTipoCobranca = -1
    lFornecedor = -1
    iFilialForn = -1
    
    For iIndice = 1 To objGrid.iLinhasExistentes
        'Verifica se outro Portador foi preenchido
        If (GridChequePagAvulso2.TextMatrix(iIndice, iGrid_Pago_Col)) = PAGO_CHECADO Then
            Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag(iIndice)
            iPortador = Codigo_Extrai(GridChequePagAvulso2.TextMatrix(iIndice, iGrid_Portador_Col))
            If iPortadorAnterior = -1 Then
                iPortadorAnterior = iPortador
                iTipoCobranca = objInfoParcPag.iTipoCobranca
                lFornecedor = objInfoParcPag.lFornecedor
                iFilialForn = objInfoParcPag.iFilialForn
            Else
                'Se for portadores diferentes --> erro
                If iPortador <> iPortadorAnterior Then Error 61127
                If objInfoParcPag.iTipoCobranca <> iTipoCobranca Then Error 61128
                If (objInfoParcPag.lFornecedor <> lFornecedor Or objInfoParcPag.iFilialForn <> iFilialForn) And (objInfoParcPag.iTipoCobranca <> TIPO_COBRANCA_BANCARIA Or iPortador = 0) Then Error 61129
            End If
        
        End If
        
    Next
    
    'Chama a tela do passo Anterior
    Call Chama_Tela("ChequePagAvulso1", gobjChequesPagAvulso)
       
    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoVoltar_Click:

    Select Case Err

        Case 61127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADORES_DIFERENTES", Err)
            
        Case 61128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_TIPO_COBR_DIFERENTE", Err)
        
        Case 61129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_FORN_DIF_NAO_COBRBANC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144484)

    End Select

    Exit Sub

End Sub

Private Sub CheckPago_Click()

Dim iLinha As Integer
Dim dTotalTitulo As Double
Dim objInfoParcPag As ClassInfoParcPag
Dim dTotalJuros As Double
Dim dTotalMultas As Double
Dim dTotalDescontos As Double

    'Passa para iLinha o número da linha em questão
    iLinha = GridChequePagAvulso2.Row
    
    'Passa a linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
    
    'Passa para o Obj se a parcela em questão foi marcada ou desmarcada
    objInfoParcPag.iSeqCheque = CInt(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col))
                    
    'Passa para as variáveis os dados da parcela que estão na tela
    dTotalTitulo = CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Valor_Col))
    dTotalJuros = CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Juros_Col))
    dTotalMultas = CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Multa_Col))
    dTotalDescontos = CDbl(GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Desconto_Col))
        
    'Se a parcela foi marcada
    If GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then
                
        'Acrescenta a nova parcela no somatório de Qtd de Títulos selecionados e Valor total de Títulos selecionados
        QtdTitulosSelecionados.Caption = CStr(CInt(QtdTitulosSelecionados.Caption) + 1)
        TotalTitulosSelecionados.Caption = CStr(Format(CDbl(TotalTitulosSelecionados.Caption) + (dTotalTitulo + dTotalJuros + dTotalMultas - dTotalDescontos), "Standard"))
        
    'Se a parcela foi desmarcada
    Else
    
        'Subtrai a parcela do somatório de Qtd de Títulos selecionados e Valor total de Títulos selecionados
        QtdTitulosSelecionados.Caption = CStr(CInt(QtdTitulosSelecionados.Caption) - 1)
        TotalTitulosSelecionados.Caption = CStr(Format(CDbl(TotalTitulosSelecionados.Caption) - (dTotalTitulo + dTotalJuros + dTotalMultas - dTotalDescontos), "Standard"))
        
    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, gcolTiposDeCobranca)
    If lErro <> SUCESSO Then Error 18999

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 18999

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144485)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente.

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Chama rotina de inicialização da saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt.objGrid Is GridChequePagAvulso2 Then

            Select Case objGridInt.objGrid.Col
            
                'Se a célula for o campo Multa
                Case iGrid_Multa_Col
                    
                    Set objGridInt.objControle = Multa
                    
                    'Chama função de tratamento de saída da célula Multa
                    lErro = Saida_Celula_Multa(objGridInt)
                    If lErro <> SUCESSO Then Error 15803
                    
                'Se a célula for o campo Juros
                Case iGrid_Juros_Col
                    
                    Set objGridInt.objControle = Juros
                    
                    'Chama função de tratamento de saída da célula Juros
                    lErro = Saida_Celula_Juros(objGridInt)
                    If lErro <> SUCESSO Then Error 15804
                
                'Se a célula for o campo Desconto
                Case iGrid_Desconto_Col
                    
                    Set objGridInt.objControle = Desconto
                    
                    'Chama função de tratamento de saída da célula Desconto
                    lErro = Saida_Celula_Desconto(objGridInt)
                    If lErro <> SUCESSO Then Error 15805
                    
            End Select

        End If

        'Chama função de finalização da saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 15792

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 15802, 15803, 15804, 15805
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15792
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144486)

    End Select

    Exit Function

End Function

Private Sub CheckPago_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CheckPago_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CheckPago_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CheckPago
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid = Nothing

    Set gobjChequesPagAvulso = Nothing
    Set gcolTiposDeCobranca = Nothing

End Sub

Private Sub Juros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Juros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Juros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Juros
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Multa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Private Sub Multa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Multa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Multa
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridChequePagAvulso2_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridChequePagAvulso2_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridChequePagAvulso2_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridChequePagAvulso2_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridChequePagAvulso2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridChequePagAvulso2_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridChequePagAvulso2_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridChequePagAvulso2_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridChequePagAvulso2_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Function Inicializa_Grid_ChequePagAvulso2(objGridInt As AdmGrid, iRegistros As Integer) As Long

    'Tela em questão
    Set objGrid.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Pagar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Tipo De Cobrança")

   'Campos de edição do grid
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Titulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (CheckPago.Name)
    objGridInt.colCampo.Add (DataVencto.Name)
    objGridInt.colCampo.Add (Juros.Name)
    objGridInt.colCampo.Add (Multa.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (Portador.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)

    iGrid_FilialEmpresa_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_Filial_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_Titulo_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_Valor_Col = 7
    iGrid_Pago_Col = 8
    iGrid_DataVencto_Col = 9
    iGrid_Juros_Col = 10
    iGrid_Multa_Col = 11
    iGrid_Desconto_Col = 12
    iGrid_Portador_Col = 13
    iGrid_TipoCobranca_Col = 14

    objGridInt.objGrid = GridChequePagAvulso2

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Todas as linhas do grid
    If objGridInt.iLinhasVisiveis >= iRegistros + 1 Then
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    Else
        objGridInt.objGrid.Rows = iRegistros + 1
    End If

    GridChequePagAvulso2.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Function Grid_Preenche(objChequesPagAvulso As ClassChequesPagAvulso)
'preenche o grid com as parcelas que podem ser pagas

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim dTotalTitulos As Double
Dim dTotalJuros As Double
Dim dTotalMultas As Double
Dim dTotalDescontos As Double
Dim iQtdTitulosSelecionados As Integer
Dim dTotalTitulosSelecionados As Double
Dim dTotalPagarTitulos As Double
Dim dTotalJurosTitulos As Double
Dim dTotalMultasTitulos As Double
Dim dTotalDescontosTitulos As Double
Dim dTotalPagarSelecionadas As Double
Dim objCodDescricao As AdmCodigoNome
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Grid_Preenche

    Set objGrid = New AdmGrid

    'Chama função de inicialização do Grid
    lErro = Inicializa_Grid_ChequePagAvulso2(objGrid, objChequesPagAvulso.colInfoParcPag.Count)
    If lErro <> SUCESSO Then gError 14250
    
    iLinha = 0

    'Percorra toda a coleção passada por parâmetro
    For Each objInfoParcPag In objChequesPagAvulso.colInfoParcPag

        iLinha = iLinha + 1

        objFilialEmpresa.iCodFilial = objInfoParcPag.iFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 82782
        
        'Passa os dados da parcela para o Grid
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objInfoParcPag.sNomeRedForn
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcPag.iFilialForn
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcPag.sSiglaDocumento
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Titulo_Col) = objInfoParcPag.lNumTitulo
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objInfoParcPag.dValor, "Standard")
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col) = objInfoParcPag.iSeqCheque
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_DataVencto_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/mm/yy")
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Juros_Col) = Format(objInfoParcPag.dValorJuros, "Standard")
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Multa_Col) = Format(objInfoParcPag.dValorMulta, "Standard")
        GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(objInfoParcPag.dValorDesconto, "Standard")
        If objInfoParcPag.iPortador <> 0 Then GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Portador_Col) = CStr(objInfoParcPag.iPortador) & SEPARADOR & objInfoParcPag.sNomeRedPortador
        
        If objInfoParcPag.iTipoCobranca <> 0 Then
            For Each objCodDescricao In gcolTiposDeCobranca
                If objCodDescricao.iCodigo = objInfoParcPag.iTipoCobranca Then GridChequePagAvulso2.TextMatrix(iLinha, iGrid_TipoCobranca_Col) = objInfoParcPag.iTipoCobranca & SEPARADOR & objCodDescricao.sNome
            Next
        End If
        
        'Faz o somatório da Qtd e Total dos Títulos
        dTotalTitulos = dTotalTitulos + objInfoParcPag.dValor
        dTotalMultasTitulos = dTotalMultasTitulos + objInfoParcPag.dValorMulta
        dTotalJurosTitulos = dTotalJurosTitulos + objInfoParcPag.dValorJuros
        dTotalDescontosTitulos = dTotalDescontosTitulos + objInfoParcPag.dValorDesconto

        'Se a CheckBox Pago estiver checada
        If GridChequePagAvulso2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then

            'Faz o somatório da Qtd e Total dos Títulos selecionados
            iQtdTitulosSelecionados = iQtdTitulosSelecionados + 1
            dTotalTitulosSelecionados = dTotalTitulosSelecionados + objInfoParcPag.dValor
            dTotalJuros = dTotalJuros + objInfoParcPag.dValorJuros
            dTotalMultas = dTotalMultas + objInfoParcPag.dValorMulta
            dTotalDescontos = dTotalDescontos + objInfoParcPag.dValorDesconto

        End If

    Next

    objGrid.iLinhasExistentes = iLinha
    
    'Passa para a tela os dados dos somatórios das parcelas
    dTotalPagarSelecionadas = dTotalTitulosSelecionados + dTotalJuros + dTotalMultas - dTotalDescontos
    dTotalPagarTitulos = dTotalTitulos + dTotalJurosTitulos + dTotalMultasTitulos - dTotalDescontosTitulos
    
    QtdTitulos.Caption = CStr(objGrid.iLinhasExistentes)
    TotalTitulos.Caption = CStr(Format(dTotalPagarTitulos, "Standard"))
    QtdTitulosSelecionados.Caption = CStr(iQtdTitulosSelecionados)
    TotalTitulosSelecionados.Caption = CStr(Format(dTotalPagarSelecionadas, "Standard"))

    'Atualiza na tela os CheckBox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    Grid_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Preenche:

    Grid_Preenche = gErr

    Select Case gErr

        Case 14250, 82782
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144487)

    End Select
    
    Exit Function

End Function

Function Trata_Parametros(Optional objChequesPagAvulso As ClassChequesPagAvulso) As Long
'Traz os dados das Parcelas a pagar para a Tela

Dim lErro As Long
Dim objCtaCorrenteInt As New ClassContasCorrentesInternas
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Trata_Parametros

    'Passa o Obj passado por parâmetro para o Obj global
    Set gobjChequesPagAvulso = objChequesPagAvulso

    objCtaCorrenteInt.iCodigo = gobjChequesPagAvulso.iCta
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objCtaCorrenteInt.iCodigo, objCtaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 43112
    
    LabelConta.Caption = CStr(objCtaCorrenteInt.iCodigo) & SEPARADOR & objCtaCorrenteInt.sNomeReduzido
    
    If gobjChequesPagAvulso.lFornecedor <> 0 Then
        
        'Prenche objFilialFornecedor
        objFornecedor.lCodigo = gobjChequesPagAvulso.lFornecedor
        'Lê o Forcecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12732 Then Error 43113
        
        LabelFornecedor.Caption = objFornecedor.sNomeReduzido
        
        If gobjChequesPagAvulso.iFilial <> 0 Then
        
            'Preenche objFilialFornecedor
            objFilialFornecedor.iCodFilial = gobjChequesPagAvulso.iFilial
            'Lê a Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 43114
            
            LabelFilial.Caption = objFilialFornecedor.sNome
    
        End If
        
    End If
    
    lErro = Grid_Preenche(objChequesPagAvulso)
    If lErro <> SUCESSO Then Error 41533
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 43112, 43113, 43114, 41533
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144488)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Calcula_Total(objInfoParcPag As ClassInfoParcPag)
'atualiza os labels com totais (geral e selecionado) dos titulos

Dim dTotalTitulos As Double
Dim dTotalJuros As Double
Dim dTotalMultas As Double
Dim dTotalDescontos As Double
Dim dTotalTitulosSelecionados As Double
Dim dTotalJurosSelecionados As Double
Dim dTotalMultasSelecionadas As Double
Dim dTotalDescontosSelecionados As Double

    'Percorre todas as linhas do Grid
    For Each objInfoParcPag In gobjChequesPagAvulso.colInfoParcPag

        'Faz o somatório do Total dos Títulos
        dTotalTitulos = dTotalTitulos + objInfoParcPag.dValor
        dTotalJuros = dTotalJuros + objInfoParcPag.dValorJuros
        dTotalMultas = dTotalMultas + objInfoParcPag.dValorMulta
        dTotalDescontos = dTotalDescontos + objInfoParcPag.dValorDesconto

        'Se a parcela em questão está checada
        If objInfoParcPag.iSeqCheque = PAGO_CHECADO Then

            'Faz o somatório do Total dos Títulos selecionados
            dTotalTitulosSelecionados = dTotalTitulosSelecionados + objInfoParcPag.dValor
            dTotalJurosSelecionados = dTotalJurosSelecionados + objInfoParcPag.dValorJuros
            dTotalMultasSelecionadas = dTotalMultasSelecionadas + objInfoParcPag.dValorMulta
            dTotalDescontosSelecionados = dTotalDescontosSelecionados + objInfoParcPag.dValorDesconto

        End If

    Next
    
    'Atualiza na tela os somatórios das parcelas
    TotalTitulos.Caption = CStr(Format(dTotalTitulos + dTotalMultas + dTotalJuros - dTotalDescontos, "Standard"))
    TotalTitulosSelecionados.Caption = CStr(Format(dTotalTitulosSelecionados + dTotalMultasSelecionadas + dTotalJurosSelecionados - dTotalDescontosSelecionados, "Standard"))
    
End Function

Private Function Saida_Celula_Multa(objGridInt As AdmGrid) As Long
'Rotina de saída da célula Multa

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Multa

    'Formata o valor da multa na tela
    Multa.Text = Format(Multa.Text, "Standard")
    
    If Multa.Text = "" Then
        Multa.Text = Format(0, "Standard")
    Else
        'Critica se o valor é positivo
        lErro = Valor_NaoNegativo_Critica(Multa.Text)
        If lErro <> SUCESSO Then Error 57843
    End If
    
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequePagAvulso2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor da multa que está na tela
    If Len(Trim(Multa.Text)) <> 0 Then
        objInfoParcPag.dValorMulta = CDbl(Multa.Text)
    Else
        objInfoParcPag.dValorMulta = 0
    End If
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15811

    Saida_Celula_Multa = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Multa:

    Saida_Celula_Multa = Err
    
    Select Case Err

        Case 15811, 57843
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144489)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Juros(objGridInt As AdmGrid) As Long
'Rotina de saída da célula Juros

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Juros

    'Formata o valor dos Juros na tela
    Juros.Text = Format(Juros.Text, "Standard")
    
    If Juros.Text = "" Then
        Juros.Text = Format(0, "Standard")
    Else
        'Critica se o valor é positivo
        lErro = Valor_NaoNegativo_Critica(Juros.Text)
        If lErro <> SUCESSO Then Error 57842
    End If
    
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequePagAvulso2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor dos Juros que está na tela
    If Len(Trim(Juros.Text)) <> 0 Then
        objInfoParcPag.dValorJuros = CDbl(Juros.Text)
    Else
        objInfoParcPag.dValorJuros = 0
    End If
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15810

    Saida_Celula_Juros = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Juros:

    Saida_Celula_Juros = Err
    
    Select Case Err

        Case 15810, 57842
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144490)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Desconto

    'Formata o valor do desconto na tela
    Desconto.Text = Format(Desconto.Text, "Standard")
    
    If Desconto.Text = "" Then
        Desconto.Text = Format(0, "Standard")
    Else
        'Critica se o valor é positivo
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 57841
    End If
    
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequePagAvulso2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPagAvulso.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor do desconto que está na tela
    If Len(Trim(Desconto.Text)) <> 0 Then
        objInfoParcPag.dValorDesconto = CDbl(Desconto.Text)
    Else
        objInfoParcPag.dValorDesconto = 0
    End If
    
    'Verifica se o Desconto não é maior que o valor Total da parcela, com Juros e Multa
    If objInfoParcPag.dValorDesconto > (objInfoParcPag.dValor + objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta) Then Error 15812
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15809

    Saida_Celula_Desconto = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = Err
    
    Select Case Err

        Case 15809, 57841
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 15812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_MENOR_DESCONTO", Err)
            Desconto.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144491)
            
    End Select
    
    Exit Function
    
End Function

Private Sub TipoCobranca_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub TipoCobranca_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub TipoCobranca_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGrid.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUE_MANUAL_P2
    Set Form_Load_Ocx = Me
    Caption = "Cheque Manual - Passo 2"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequePagAvulso2"
    
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

'***** fim do trecho a ser copiado ******

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub


Private Sub TotalTitulosSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulosSelecionados, Source, X, Y)
End Sub

Private Sub TotalTitulosSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulosSelecionados, Button, Shift, X, Y)
End Sub

Private Sub QtdTitulosSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdTitulosSelecionados, Source, X, Y)
End Sub

Private Sub QtdTitulosSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdTitulosSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub QtdTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdTitulos, Source, X, Y)
End Sub

Private Sub QtdTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub LabelConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConta, Source, X, Y)
End Sub

Private Sub LabelConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConta, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

