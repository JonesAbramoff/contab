VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl FluxoSintProjOcx 
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   3900
   ScaleWidth      =   9480
   Begin VB.CommandButton Botao_GraficoSistema 
      Caption         =   "Gráfico Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4210
      Picture         =   "FluxoSintProjOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   1590
   End
   Begin MSMask.MaskEdBox AcumuladoPercentual 
      Height          =   225
      Left            =   6870
      TabIndex        =   22
      Top             =   1800
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox AcumuladoAjustado 
      Height          =   225
      Left            =   5310
      TabIndex        =   21
      Top             =   1740
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox AcumuladoSist 
      Height          =   225
      Left            =   3930
      TabIndex        =   20
      Top             =   1770
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Botao_GraficoAjustado 
      Caption         =   "Gráfico Ajustado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6150
      Picture         =   "FluxoSintProjOcx.ctx":0432
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1590
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   8130
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoSintProjOcx.ctx":0864
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "FluxoSintProjOcx.ctx":09BE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton Botao_ImprimeFluxo 
      Caption         =   "Imprime Fluxo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2270
      Picture         =   "FluxoSintProjOcx.ctx":0B3C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1590
   End
   Begin VB.CommandButton Botao_ExibeFluxo 
      Caption         =   "Exibe Fluxo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   330
      Picture         =   "FluxoSintProjOcx.ctx":0C3E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
   Begin MSMask.MaskEdBox ValorAjustadoRec 
      Height          =   225
      Left            =   2565
      TabIndex        =   10
      Top             =   3540
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox ValorSistemaRec 
      Height          =   225
      Left            =   1590
      TabIndex        =   9
      Top             =   3495
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   225
      Left            =   600
      TabIndex        =   8
      Top             =   3510
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PercentualRec 
      Height          =   225
      Left            =   3600
      TabIndex        =   11
      Top             =   3180
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox ValorAjustadoPag 
      Height          =   225
      Left            =   5025
      TabIndex        =   12
      Top             =   3480
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox ValorSistemaPag 
      Height          =   225
      Left            =   4065
      TabIndex        =   13
      Top             =   3510
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PercentualPag 
      Height          =   225
      Left            =   6540
      TabIndex        =   14
      Top             =   3360
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox ValorAjustadoTes 
      Height          =   225
      Left            =   5415
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox ValorSistemaTes 
      Height          =   225
      Left            =   4410
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PercentualTes 
      Height          =   225
      Left            =   6330
      TabIndex        =   4
      Top             =   1245
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox SaldoAjustado 
      Height          =   225
      Left            =   7875
      TabIndex        =   6
      Top             =   1215
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox SaldoSistema 
      Height          =   225
      Left            =   6885
      TabIndex        =   5
      Top             =   1215
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SaldoPercentual 
      Height          =   225
      Left            =   8430
      TabIndex        =   7
      Top             =   1485
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSFlexGridLib.MSFlexGrid GridFCaixa 
      Height          =   2805
      Left            =   120
      TabIndex        =   15
      Top             =   810
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   11
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   0
   End
End
Attribute VB_Name = "FluxoSintProjOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim objFluxo1 As ClassFluxo
Dim iAlterado As Integer

'Geração de gráfico
Const GRAFICO_AJUSTADO = 1
Const GRAFICO_SISTEMA = 2

'Colunas do Grid
Const GRID_DATA_COL = 1
Const GRID_VALOR_SISTEMA_RECEBER_COL = 2
Const GRID_VALOR_AJUSTADO_RECEBER_COL = 3
Const GRID_VALOR_PERCENTUAL_RECEBER_COL = 4
Const GRID_VALOR_SISTEMA_PAGAMENTO_COL = 5
Const GRID_VALOR_AJUSTADO_PAGAMENTO_COL = 6
Const GRID_VALOR_PERCENTUAL_PAGAMENTO_COL = 7
Const GRID_VALOR_SISTEMA_TESOURARIA_COL = 8
Const GRID_VALOR_AJUSTADO_TESOURARIA_COL = 9
Const GRID_VALOR_PERCENTUAL_TESOURARIA_COL = 10
Const GRID_SALDO_SISTEMA_COL = 11
Const GRID_SALDO_AJUSTADO_COL = 12
Const GRID_SALDO_PERCENTUAL_COL = 13
Const GRID_SALDO_ACUMULADO_SISTEMA_COL = 14
Const GRID_SALDO_ACUMULADO_AJUSTADO_COL = 15
Const GRID_SALDO_ACUMULADO_PERCENTUAL_COL = 16

Private Sub Botao_ExibeFluxo_Click()

Dim lErro As Long

On Error GoTo Erro_Botao_ExibeFluxo_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 55927

    lErro = ExibeFluxoSint()
    If lErro <> SUCESSO Then Error 21102

    iAlterado = 0

    Exit Sub

Erro_Botao_ExibeFluxo_Click:

    Select Case Err

        Case 21102, 55927

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160463)

    End Select

    Exit Sub

End Sub

Private Sub Botao_GraficoSistema_Click()
'Dispara a geração do gráfico

Dim lErro As Long
Dim iGrafico As Integer

On Error GoTo Erro_Botao_GraficoAjustado_Click
    
    'Define que o gráfico será gerado em cima dos valores do sistema
    iGrafico = GRAFICO_SISTEMA
    
    'Gera os dados que serão utilizados na montagem do gráfico
    lErro = Gera_Grafico(iGrafico)
    If lErro <> SUCESSO Then gError 79951
    
    Exit Sub
    
Erro_Botao_GraficoAjustado_Click:

    Select Case gErr
        
        Case 79951
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160464)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Botao_GraficoAjustado_Click()
'Dispara a geração de gráfico

Dim lErro As Long
Dim iGrafico As Integer

On Error GoTo Erro_Botao_GraficoAjustado_Click
    
    'Define que o gráfico será gerado em cima dos valores ajustados
    iGrafico = GRAFICO_AJUSTADO
    
    'Gera os dados que serão utilizados na montagem do gráfico
    lErro = Gera_Grafico(iGrafico)
    If lErro <> SUCESSO Then gError 79950
    
    Exit Sub
    
Erro_Botao_GraficoAjustado_Click:

    Select Case gErr
        
        Case 79950
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160465)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Gera_Grafico(iGrafico As Integer) As Long
'Obtém os dados necessários para gerar o gráfico
'Seta configurações do gráfico

'******************** Função alterada por Luiz Gustavo de Freitas Nogueira em 04/06/2001 ********************
'Alterações
'Inclusão de tratamento de exibição de legenda
'Inclusão de tratamento de exibição de DataLabels para cada série
'************************************************************************************************************

Dim iColuna As Integer
Dim iLinha As Integer
Dim lErro As Long
Dim sPercentual As String
Dim objPlanilha As New ClassPlanilhaExcel
Dim objColunas As ClassColunasExcel
Dim objCelulas As New ClassCelulasExcel

On Error GoTo Erro_Gera_Grafico

    MousePointer = vbHourglass
    
    'Se o gráfico não é sobre os valores ajustados nem sobre os valores do sistema = > erro
    If iGrafico <> GRAFICO_AJUSTADO And iGrafico <> GRAFICO_SISTEMA Then gError 79949
    
    'Para cada coluna existente no Grid
    For iColuna = 1 To objGrid1.colColuna.Count - 1
        
        'Instancia uma nova classe para objColunas
        Set objColunas = New ClassColunasExcel
        
        'Para cada linha existente dentro da coluna
        For iLinha = 0 To objGrid1.iLinhasExistentes
        
            'Instancia uma nova classe para objCelulas
            Set objCelulas = New ClassCelulasExcel
        
            'Seleciona as colunas de acordo com o tipo de dados e aplica formatos específicos
            Select Case iColuna
                
                'Se for uma data
                Case GRID_DATA_COL
                        objCelulas.vValor = GridFCaixa.TextMatrix(iLinha, iColuna)
                    
                'Se for moeda
                Case GRID_VALOR_SISTEMA_RECEBER_COL, GRID_VALOR_AJUSTADO_RECEBER_COL, GRID_VALOR_SISTEMA_PAGAMENTO_COL, GRID_VALOR_AJUSTADO_PAGAMENTO_COL, GRID_VALOR_SISTEMA_TESOURARIA_COL, GRID_VALOR_AJUSTADO_TESOURARIA_COL, GRID_SALDO_SISTEMA_COL, GRID_SALDO_AJUSTADO_COL, GRID_SALDO_ACUMULADO_SISTEMA_COL, GRID_SALDO_ACUMULADO_AJUSTADO_COL
                    'Converte o valor para double ou string
                    If Len(Trim(GridFCaixa.TextMatrix(iLinha, iColuna))) > 0 And IsNumeric(GridFCaixa.TextMatrix(iLinha, iColuna)) Then
                        objCelulas.vValor = CDbl(GridFCaixa.TextMatrix(iLinha, iColuna))
                    Else
                        objCelulas.vValor = GridFCaixa.TextMatrix(iLinha, iColuna)
                    End If
                
                'Se for percentual
                Case GRID_VALOR_PERCENTUAL_RECEBER_COL, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL, GRID_VALOR_PERCENTUAL_TESOURARIA_COL, GRID_SALDO_PERCENTUAL_COL, GRID_SALDO_ACUMULADO_PERCENTUAL_COL
                    'Converte o valor para double ou string
                    If Len(Trim(GridFCaixa.TextMatrix(iLinha, iColuna))) > 0 Then
                    
                        sPercentual = Mid(GridFCaixa.TextMatrix(iLinha, iColuna), 1, InStr(1, GridFCaixa.TextMatrix(iLinha, iColuna), "%") - 1)
                    
                        If IsNumeric(sPercentual) Then
                            objCelulas.vValor = CDbl(sPercentual)
                        Else
                            objCelulas.vValor = GridFCaixa.TextMatrix(iLinha, iColuna)
                        End If
                    End If
                
                'Se a coluna não foi tratada
                Case Else
                    gError 79948
            
            End Select
            
            'Adiciona a célula à coleção de células
            objColunas.colCelulas.Add objCelulas
                
        Next
        
        'Seleciona as colunas que farão parte do gráfico
        Select Case iColuna
        
            'Se for a coluna de data
            Case GRID_DATA_COL
                'Faz parte do gráfico no eixo X
                objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
            
            'Se for a coluna de saldo ajustado ou saldo sistema
            Case GRID_SALDO_AJUSTADO_COL, GRID_SALDO_SISTEMA_COL
                
                'Se o gráfico é sobre os valores ajustado e a coluna é a de valores ajustados
                If iGrafico = GRAFICO_AJUSTADO And iColuna = GRID_SALDO_AJUSTADO_COL Then
                    'Informa ao excel que a coluna faz parte do gráfico no eixo Y
                    objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                    'Informa ao excel o tipo de gráfico que será utilizado para exibir os dados da série
                    objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
                    'Informa ao excel não exibir os labels
                    objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS
                End If
                
                'Se o gráfico é sobre os valores gerados pelo sistema e a coluna é a dos valores gerados pelo sistema
                If iGrafico = GRAFICO_SISTEMA And iColuna = GRID_SALDO_SISTEMA_COL Then
                    'Informa ao excel que a coluna faz parte do gráfico no eixo Y
                    objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                    'Informa ao excel o tipo de gráfico que será utilizado para exibir os dados da série
                    objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_COLUMN_CLUSTERED
                    'Informa ao excel não exibir os labels
                    objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS
                End If
                                
            'Se for a coluna de saldo acumulado ajustado ou saldo acumulado sistema
            Case GRID_SALDO_ACUMULADO_AJUSTADO_COL, GRID_SALDO_ACUMULADO_SISTEMA_COL
                
                'Se o gráfico é sobre os valores ajustado e a coluna é a de valores ajustados
                If iGrafico = GRAFICO_AJUSTADO And iColuna = GRID_SALDO_ACUMULADO_AJUSTADO_COL Then
                    'Faz parte do gráfico no eixo X
                    objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                    'Informa o tipo de gráfico que será utilizado para exibir os dados da série
                    objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_LINE_MARKERS
                    'Informa ao excel não exibir os labels
                    objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS
                End If
            
                'Se o gráfico é sobre os valores gerados pelo sistema e a coluna é a dos valores gerados pelo sistema
                If iGrafico = GRAFICO_SISTEMA And iColuna = GRID_SALDO_ACUMULADO_SISTEMA_COL Then
                    'Faz parte do gráfico no eixo X
                    objColunas.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
                    'Informa o tipo de gráfico que será utilizado para exibir os dados da série
                    objColunas.lTipoGraficoColuna = EXCEL_GRAFICO_LINE_MARKERS
                    'Informa ao excel não exibir os labels
                    objColunas.lDataLabels = EXCEL_NAO_EXIBE_LABELS
                End If
            
            'Se for uma outra coluna
            Case Else
                'Não participa do Gráfico
                objColunas.iParticipaGrafico = EXCEL_NAO_PARTICIPA_GRAFICO
            
        End Select
        
        'Adiciona a coluna à coleção de colunas
        objPlanilha.colColunas.Add objColunas
    
    Next
    
    'Se o Gráfico é Ajustado
    If iGrafico = GRAFICO_AJUSTADO Then
    
        'Informa ao excel o nome do gráfico e o nome da planilha como ajustado
        objPlanilha.sNomeGrafico = "Gráfico - Fluxo Ajustado"
        objPlanilha.sNomePlanilha = "Planilha - Fluxo Ajustado"
    
    ElseIf iGrafico = GRAFICO_SISTEMA Then
    
        'Informa ao excel o nome do gráfico e o nome da planilha como sistema
        objPlanilha.sNomeGrafico = "Gráfico - Fluxo Sistema"
        objPlanilha.sNomePlanilha = "Planilha - Fluxo Sistema"
    End If
    
    'Informa ao excel o título do gráfico
    objPlanilha.sTituloGrafico = Parent.Caption
    
    'Informa ao excel a posição da legenda
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_DIREITA
    
    'Informa ao excel a posição dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW
    
    'Informa ao excel a orientação dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD
    
    'Informa ao excel que a plotagem do dados será por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS
    
    'Monta a planilha e o gráfico com os dados passados em objPlanilha
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 79946

    MousePointer = vbDefault
    
    Gera_Grafico = SUCESSO
    
    Exit Function
    
Erro_Gera_Grafico:

    Gera_Grafico = gErr
    
    Select Case gErr
        
        Case 79946
        
        Case 79948
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORES_COLUNAS_NAO_TRATADOS_GRAFICO", gErr, iColuna)
            
        Case 79949
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAFICO_VALORES_A_EXIBIR_NAO_DEFINIDOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160466)
            
    End Select
    
    MousePointer = vbDefault
    
    Exit Function
    
End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = Inicializa_GridFCaixa()
    If lErro <> SUCESSO Then Error 21077
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21077

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160467)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFluxo As ClassFluxo) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se objFluxo não estiver preenchido ==> erro
    If (objFluxo Is Nothing) Then Error 21078

    Set objFluxo1 = objFluxo

    lErro = ExibeFluxoSint()
    If lErro <> SUCESSO Then Error 21103

    Parent.Caption = "Fluxo Sintético " & objFluxo.sFluxo & " - Projetado "
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 21078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 21103

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160468)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Inicializa_GridFCaixa() As Long

Dim iIndice As Integer

    Set objGrid1 = New AdmGrid

    'tela em questão
    Set objGrid1.objForm = Me

    objGrid1.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid1.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'titulos do grid
    objGrid1.colColuna.Add ("")
    objGrid1.colColuna.Add ("Data")
    objGrid1.colColuna.Add ("Rec Sist")
    objGrid1.colColuna.Add ("Rec Ajust")
    objGrid1.colColuna.Add ("Rec %")
    objGrid1.colColuna.Add ("Pag  Sist")
    objGrid1.colColuna.Add ("Pag Ajust")
    objGrid1.colColuna.Add ("Pag %")
    objGrid1.colColuna.Add ("Tes  Sist")
    objGrid1.colColuna.Add ("Tes Ajust")
    objGrid1.colColuna.Add ("Tes %")
    objGrid1.colColuna.Add ("Saldo Sist")
    objGrid1.colColuna.Add ("Saldo Ajust")
    objGrid1.colColuna.Add ("Saldo %")
    objGrid1.colColuna.Add ("Acumul.Sist.")
    objGrid1.colColuna.Add ("Acumul.Ajust.")
    objGrid1.colColuna.Add ("Acumul. %")

   'campos de edição do grid
    objGrid1.colCampo.Add (Data.Name)
    objGrid1.colCampo.Add (ValorSistemaRec.Name)
    objGrid1.colCampo.Add (ValorAjustadoRec.Name)
    objGrid1.colCampo.Add (PercentualRec.Name)
    objGrid1.colCampo.Add (ValorSistemaPag.Name)
    objGrid1.colCampo.Add (ValorAjustadoPag.Name)
    objGrid1.colCampo.Add (PercentualPag.Name)
    objGrid1.colCampo.Add (ValorSistemaTes.Name)
    objGrid1.colCampo.Add (ValorAjustadoTes.Name)
    objGrid1.colCampo.Add (PercentualTes.Name)
    objGrid1.colCampo.Add (SaldoSistema.Name)
    objGrid1.colCampo.Add (SaldoAjustado.Name)
    objGrid1.colCampo.Add (SaldoPercentual.Name)
    objGrid1.colCampo.Add (AcumuladoSist.Name)
    objGrid1.colCampo.Add (AcumuladoAjustado.Name)
    objGrid1.colCampo.Add (AcumuladoPercentual.Name)
    
    objGrid1.objGrid = GridFCaixa

    'linhas visiveis do grid sem contar com as linhas fixas
    objGrid1.iLinhasVisiveis = 10
    
    objGrid1.objGrid.ColWidth(0) = 300
    
    objGrid1.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid1.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid1)
    
    objGrid1.objGrid.RowHeight(0) = 500
    
    Inicializa_GridFCaixa = SUCESSO

End Function

Function Preenche_GridFCaixa(colFluxoSint As Collection, colFluxoSldIni As Collection) As Long
'preenche o grid com os dados contidos na coleção colFluxoSint

Dim lErro As Long
Dim iIndice As Integer
Dim objFluxoSint As ClassFluxoSint
Dim dColunaSomaValorAjustadoRec As Double
Dim dColunaSomaValorSistemaRec As Double
Dim dColunaSomaValorAjustadoPag As Double
Dim dColunaSomaValorSistemaPag As Double
Dim dColunaSomaValorAjustadoTes As Double
Dim dColunaSomaValorSistemaTes As Double
Dim dColunaSomaPercRec As Double
Dim dColunaSomaPercPag As Double
Dim dColunaSomaPercTes As Double
Dim dColunaSomaPercSaldo As Double
Dim dPercRec As Double, dPercPag As Double, dPercTes As Double, dPercSaldo As Double, dPercSaldoAcumulado As Double
Dim dSaldoAcumuladoSist As Double
Dim dSaldoAcumuladoAjust As Double
Dim dSaldoInicialAjust As Double
Dim dSaldoInicialSist As Double
Dim objFluxoSldIni As ClassFluxoSldIni

On Error GoTo Erro_Preenche_GridFCaixa

    GridFCaixa.Clear

    If colFluxoSint.Count < objGrid1.iLinhasVisiveis Then
        objGrid1.objGrid.Rows = objGrid1.iLinhasVisiveis + 3
    Else
        objGrid1.objGrid.Rows = colFluxoSint.Count + 3
    End If

    Call Grid_Inicializa(objGrid1)

    objGrid1.iLinhasExistentes = colFluxoSint.Count + 2
    
    For Each objFluxoSldIni In colFluxoSldIni
        dSaldoInicialAjust = dSaldoInicialAjust + objFluxoSldIni.dSaldoAjustado
        dSaldoInicialSist = dSaldoInicialSist + objFluxoSldIni.dSaldoSistema
    Next
    
    dSaldoAcumuladoAjust = dSaldoInicialAjust
    dSaldoAcumuladoSist = dSaldoInicialSist
    
    GridFCaixa.TextMatrix(1, GRID_DATA_COL) = "Saldo Inicial:"
    GridFCaixa.TextMatrix(1, GRID_SALDO_AJUSTADO_COL) = Format(dSaldoInicialAjust, "Standard")
    GridFCaixa.TextMatrix(1, GRID_SALDO_SISTEMA_COL) = Format(dSaldoInicialSist, "Standard")
    GridFCaixa.TextMatrix(1, GRID_SALDO_ACUMULADO_AJUSTADO_COL) = Format(dSaldoInicialAjust, "Standard") 'Coluna Acumul.Ajust.
    GridFCaixa.TextMatrix(1, GRID_SALDO_ACUMULADO_SISTEMA_COL) = Format(dSaldoInicialSist, "Standard") 'Coluna Acumul.Sist.
    
    'preenche o grid com os dados retornados na coleção colFluxoSint
    For iIndice = 1 To colFluxoSint.Count

        Set objFluxoSint = colFluxoSint.Item(iIndice)

        iIndice = iIndice + 1
        
        'Preenche o campo data
        GridFCaixa.TextMatrix(iIndice, GRID_DATA_COL) = Format(objFluxoSint.dtData, "dd/mm/yyyy")
        
        'Preenche os campos Rec Sist e Rec Ajust
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_RECEBER_COL) = Format(objFluxoSint.dRecValorSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_RECEBER_COL) = Format(objFluxoSint.dRecValorAjustado, "Standard")
        
        'Calcula o valor do campo Rec % e o preenche
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_RECEBER_COL) = Format(0, "Percent")
        If objFluxoSint.dRecValorSistema <> 0 Then
            dPercRec = (objFluxoSint.dRecValorAjustado / objFluxoSint.dRecValorSistema)
            GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_RECEBER_COL) = Format(dPercRec, "Percent")
        End If
        
        'Preenche os campos Pag Sist e Pag Ajust
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_PAGAMENTO_COL) = Format(objFluxoSint.dPagValorSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_PAGAMENTO_COL) = Format(objFluxoSint.dPagValorAjustado, "Standard")
        
        'Calcula o valor do campo Pag % e o preenche
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL) = Format(0, "Percent")
        If objFluxoSint.dPagValorSistema <> 0 Then
            dPercPag = (objFluxoSint.dPagValorAjustado / objFluxoSint.dPagValorSistema)
            GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL) = Format(dPercPag, "Percent")
        End If
        
        'Preenche os campos Tes Sist e Tes Ajust
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_TESOURARIA_COL) = Format(objFluxoSint.dTesValorSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_TESOURARIA_COL) = Format(objFluxoSint.dTesValorAjustado, "Standard")
        
        'Calcula o valor do campo Tes % e o preenche
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(0, "Percent")
        If objFluxoSint.dTesValorSistema <> 0 Then
            dPercTes = (objFluxoSint.dTesValorAjustado / objFluxoSint.dTesValorSistema)
            GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(dPercTes, "Percent")
        End If
        
        'Calcula e preenche os campos Saldo Sist. e Saldo Ajust
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_SISTEMA_COL) = Format(objFluxoSint.dRecValorSistema + objFluxoSint.dTesValorSistema - objFluxoSint.dPagValorSistema, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_AJUSTADO_COL) = Format(objFluxoSint.dRecValorAjustado + objFluxoSint.dTesValorAjustado - objFluxoSint.dPagValorAjustado, "Standard")
        
        'Calcula o valor do Campo Saldo % e o preenche
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(0, "Percent")
        If (objFluxoSint.dRecValorSistema + objFluxoSint.dTesValorSistema - objFluxoSint.dPagValorSistema) <> 0 Then
            dPercSaldo = (objFluxoSint.dRecValorAjustado + objFluxoSint.dTesValorAjustado - objFluxoSint.dPagValorAjustado) / (objFluxoSint.dRecValorSistema + objFluxoSint.dTesValorSistema - objFluxoSint.dPagValorSistema)
            GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(dPercSaldo, "Percent")
        End If
        
        'Calcula e preenche os campos Acumulado Sist. e Acumulado Ajust
        dSaldoAcumuladoSist = dSaldoAcumuladoSist + objFluxoSint.dRecValorSistema + objFluxoSint.dTesValorSistema - objFluxoSint.dPagValorSistema
        dSaldoAcumuladoAjust = dSaldoAcumuladoAjust + objFluxoSint.dRecValorAjustado + objFluxoSint.dTesValorAjustado - objFluxoSint.dPagValorAjustado
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_SISTEMA_COL) = Format(dSaldoAcumuladoSist, "Standard")
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_AJUSTADO_COL) = Format(dSaldoAcumuladoAjust, "Standard")
        
        'Calcula o campo Acumul. % e o preenche
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_PERCENTUAL_COL) = Format(0, "Percent")
        If dSaldoAcumuladoSist <> 0 Then
            dPercSaldoAcumulado = dSaldoAcumuladoAjust / dSaldoAcumuladoSist
            GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_PERCENTUAL_COL) = Format(dPercSaldoAcumulado, "Percent")
        End If
        
        'Acumula os valores totais que serão exibidos na última linha
        dColunaSomaValorAjustadoRec = dColunaSomaValorAjustadoRec + objFluxoSint.dRecValorAjustado
        dColunaSomaValorSistemaRec = dColunaSomaValorSistemaRec + objFluxoSint.dRecValorSistema
        dColunaSomaValorAjustadoPag = dColunaSomaValorAjustadoPag + objFluxoSint.dPagValorAjustado
        dColunaSomaValorSistemaPag = dColunaSomaValorSistemaPag + objFluxoSint.dPagValorSistema
        dColunaSomaValorAjustadoTes = dColunaSomaValorAjustadoTes + objFluxoSint.dTesValorAjustado
        dColunaSomaValorSistemaTes = dColunaSomaValorSistemaTes + objFluxoSint.dTesValorSistema
        iIndice = iIndice - 1
        
    Next
    
    iIndice = iIndice + 1
    
    'Preenche a última linha do Grid com os totais
    GridFCaixa.TextMatrix(iIndice, GRID_DATA_COL) = "Totais:" 'Coluna Data
    
    'Preenche os totais de Rec Sist e Rec Ajust
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_RECEBER_COL) = Format(dColunaSomaValorSistemaRec, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_RECEBER_COL) = Format(dColunaSomaValorAjustadoRec, "Standard")
    
    'Calcula e preenche o valor total de Rec %
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_RECEBER_COL) = Format(0, "Percent")
    If dColunaSomaValorSistemaRec <> 0 Then
        dColunaSomaPercRec = dColunaSomaValorAjustadoRec / dColunaSomaValorSistemaRec
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_RECEBER_COL) = Format(dColunaSomaPercRec, "Percent")
    End If
    
    'Preenche os totais de Pag Sist e Pag Ajust
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_PAGAMENTO_COL) = Format(dColunaSomaValorSistemaPag, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_PAGAMENTO_COL) = Format(dColunaSomaValorAjustadoPag, "Standard")
    
    'Calcula e preenche o valor total de Pag %
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL) = Format(0, "Percent")
    If dColunaSomaValorSistemaPag <> 0 Then
        dColunaSomaPercPag = dColunaSomaValorAjustadoPag / dColunaSomaValorSistemaPag
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL) = Format(dColunaSomaPercPag, "Percent")
    End If
    
    'Preenche os totais de Tes Sist e Tes Ajust
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_TESOURARIA_COL) = Format(dColunaSomaValorSistemaTes, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_TESOURARIA_COL) = Format(dColunaSomaValorAjustadoTes, "Standard")
    
    'Calcula e preenche o valor total de Tes %
    GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(0, "Percent")
    If dColunaSomaValorSistemaTes <> 0 Then
        dColunaSomaPercTes = dColunaSomaValorAjustadoTes / dColunaSomaValorSistemaTes
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(dColunaSomaPercTes, "Percent")
    End If
    
    
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_SISTEMA_COL) = Format(dSaldoInicialSist + dColunaSomaValorSistemaRec + dColunaSomaValorSistemaTes - dColunaSomaValorSistemaPag, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_AJUSTADO_COL) = Format(dSaldoInicialAjust + dColunaSomaValorAjustadoRec + dColunaSomaValorAjustadoTes - dColunaSomaValorAjustadoPag, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(0, "Percent")
    If (dSaldoInicialAjust + dColunaSomaValorAjustadoRec + dColunaSomaValorAjustadoTes - dColunaSomaValorAjustadoPag) <> 0 Then
        dColunaSomaPercSaldo = (dSaldoInicialSist + dColunaSomaValorSistemaRec + dColunaSomaValorSistemaTes - dColunaSomaValorSistemaPag) / (dSaldoInicialAjust + dColunaSomaValorAjustadoRec + dColunaSomaValorAjustadoTes - dColunaSomaValorAjustadoPag)
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(dColunaSomaPercSaldo, "Percent")
    End If
    
    
    'Preenche os totais de Acumul.Sist. e Acumul.Ajust.
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_SISTEMA_COL) = Format(dSaldoAcumuladoSist, "Standard")
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_AJUSTADO_COL) = Format(dSaldoAcumuladoAjust, "Standard")
    
    'Calcula e preenche o valor total de Acumul.%
    GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_PERCENTUAL_COL) = Format(0, "Percent")
    If dSaldoAcumuladoSist <> 0 Then
        dPercSaldoAcumulado = (dSaldoAcumuladoAjust / dSaldoAcumuladoSist)
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_ACUMULADO_PERCENTUAL_COL) = Format(dPercSaldoAcumulado, "Percent")
    End If
    
    Preenche_GridFCaixa = SUCESSO

    Exit Function

Erro_Preenche_GridFCaixa:

    Preenche_GridFCaixa = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160469)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
        
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid1 = Nothing

    Set objFluxo1 = Nothing
    
End Sub

Private Sub GridFCaixa_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid1, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid1, iAlterado)
        End If

End Sub

Private Sub GridFCaixa_DblClick()

    If GridFCaixa.Row > 1 Then

        objFluxo1.dtData = GridFCaixa.TextMatrix(GridFCaixa.Row, 1)
    
        If GridFCaixa.Col = 2 Or GridFCaixa.Col = 3 Or GridFCaixa.Col = 4 Then
    
            Call Chama_Tela("FluxoReceb", objFluxo1)
    
        ElseIf GridFCaixa.Col = 5 Or GridFCaixa.Col = 6 Or GridFCaixa.Col = 7 Then
    
            Call Chama_Tela("FluxoPag", objFluxo1)
    
        End If

    End If

End Sub

Private Sub GridFCaixa_GotFocus()
    Call Grid_Recebe_Foco(objGrid1)
End Sub

Private Sub GridFCaixa_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridFCaixa_LeaveCell()
    Call Saida_Celula(objGrid1)
End Sub

Private Sub GridFCaixa_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double

    Call Grid_Trata_Tecla2(KeyCode, objGrid1)

End Sub

Private Sub GridFCaixa_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridFCaixa_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid1)
End Sub

Private Sub GridFCaixa_RowColChange()
    Call Grid_RowColChange(objGrid1)
End Sub

Private Sub GridFCaixa_Scroll()
    Call Grid_Scroll(objGrid1)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If GridFCaixa.Col = GRID_VALOR_AJUSTADO_TESOURARIA_COL Then

            lErro = Saida_Celula_ValorAjustadoTes(objGridInt)
            If lErro <> SUCESSO Then Error 21101

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 21081

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 21081
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 21101

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160470)

    End Select

    Exit Function

End Function

Private Function ValorPercSaldo(iLinha As Integer) As Double

Dim lTamanho As Long
Dim dPercentDesc As Double

    lTamanho = Len(Trim(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_PERCENTUAL_COL)))
    
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Left(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_PERCENTUAL_COL), lTamanho - 1))
    Else
        dPercentDesc = 0
    End If

    ValorPercSaldo = dPercentDesc
    
End Function

Private Function Saida_Celula_ValorAjustadoTes(objGridInt As AdmGrid) As Long
'faz a critica da celula valor ajustado do grid que está deixando de ser a corrente

Dim iIndice As Integer
Dim lErro As Long
Dim dColunaSomaValorAjustadoTes As Double
Dim dColunaSomaSaldoAjustado As Double
Dim objFluxoSint As ClassFluxoSint
Dim dColunaSomaSaldoSistema As Double
Dim dSaldoValorSistema As Double

On Error GoTo Erro_Saida_Celula_ValorAjustadoTes

    Set objGridInt.objControle = ValorAjustadoTes

    If objGrid1.iLinhasExistentes <> GridFCaixa.Row And GridFCaixa.Row > 1 Then
    
        If Len(ValorAjustadoTes.ClipText) > 0 Then
            lErro = Valor_Critica(ValorAjustadoTes.Text)
            If lErro <> SUCESSO Then Error 21099
            
        Else
            ValorAjustadoTes.Text = "0"
        End If
    
        Set objFluxoSint = New ClassFluxoSint
    
        objFluxoSint.dTesValorAjustado = CDbl(ValorAjustadoTes.Text)
        objFluxoSint.dTesValorSistema = CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_SISTEMA_TESOURARIA_COL))
        objFluxoSint.dRecValorAjustado = CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_AJUSTADO_RECEBER_COL))
        objFluxoSint.dPagValorAjustado = CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_AJUSTADO_PAGAMENTO_COL))
        
        If objFluxoSint.dTesValorSistema <> 0 Then
            GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(objFluxoSint.dTesValorAjustado / objFluxoSint.dTesValorSistema, "Percent")
        Else
            GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(0, "Percent")
        End If
        
        GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SALDO_AJUSTADO_COL) = Format(objFluxoSint.dRecValorAjustado + objFluxoSint.dTesValorAjustado - objFluxoSint.dPagValorAjustado, "Standard")
        objFluxoSint.dSaldoValorAjustado = CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SALDO_AJUSTADO_COL))
        objFluxoSint.dSaldoValorSistema = CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SALDO_SISTEMA_COL))
        
        If objFluxoSint.dSaldoValorSistema <> 0 Then
            GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SALDO_PERCENTUAL_COL) = Format(objFluxoSint.dSaldoValorAjustado / objFluxoSint.dSaldoValorSistema, "Percent")
        Else
            GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_SALDO_PERCENTUAL_COL) = Format(0, "Percent")
        End If
        
        dColunaSomaSaldoAjustado = 0
        
        For iIndice = 1 To objGrid1.iLinhasExistentes - 1
    
            dColunaSomaSaldoAjustado = dColunaSomaSaldoAjustado + CDbl(GridFCaixa.TextMatrix(iIndice, GRID_SALDO_AJUSTADO_COL))
            
        Next
    
        dColunaSomaSaldoSistema = CDbl(GridFCaixa.TextMatrix(iIndice, GRID_SALDO_SISTEMA_COL))
    
        dColunaSomaValorAjustadoTes = CDbl(GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_TESOURARIA_COL)) - CDbl(GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_AJUSTADO_TESOURARIA_COL)) + CDbl(ValorAjustadoTes.Text)
    
        GridFCaixa.TextMatrix(iIndice, GRID_VALOR_AJUSTADO_TESOURARIA_COL) = Format(dColunaSomaValorAjustadoTes, "Standard")
        
        dSaldoValorSistema = CDbl(GridFCaixa.TextMatrix(iIndice, GRID_VALOR_SISTEMA_TESOURARIA_COL))
        
        If dSaldoValorSistema <> 0 Then
            GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(dColunaSomaValorAjustadoTes / dSaldoValorSistema, "Percent")
        Else
            GridFCaixa.TextMatrix(iIndice, GRID_VALOR_PERCENTUAL_TESOURARIA_COL) = Format(0, "Percent")
        End If
        
        GridFCaixa.TextMatrix(iIndice, GRID_SALDO_AJUSTADO_COL) = Format(dColunaSomaSaldoAjustado, "Standard")
        
        If dColunaSomaSaldoSistema <> 0 Then
            GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(dColunaSomaSaldoAjustado / dColunaSomaSaldoSistema, "Percent")
        Else
            GridFCaixa.TextMatrix(iIndice, GRID_SALDO_PERCENTUAL_COL) = Format(0, "Percent")
        End If
        
    Else
    
        ValorAjustadoTes.Text = GridFCaixa.TextMatrix(GridFCaixa.Row, GRID_VALOR_AJUSTADO_TESOURARIA_COL)
         
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 21100
        
    Saida_Celula_ValorAjustadoTes = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorAjustadoTes:

    Saida_Celula_ValorAjustadoTes = Err

    Select Case Err

        Case 21099, 21100
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160471)

    End Select

    Exit Function

End Function

Private Function ExibeFluxoSint() As Long

Dim lErro As Long
Dim colFluxoSint As New Collection
Dim colFluxoSldIni As New Collection

On Error GoTo Erro_ExibeFluxoSint

    'le os valores relativos ao fluxo
    lErro = CF("FluxoSintetico_Le", colFluxoSint, objFluxo1.lFluxoId, FLUXOSINT_PROJ)
    If lErro <> SUCESSO And lErro <> 21089 Then gError 21090

    If lErro = 21089 Then Call Rotina_Aviso(vbOK, "AVISO_NUM_FLUXOSINTETICO_ULTRAPASSOU_LIMITE", Format(Data.Text, "dd/mm/yyyy"), MAX_FLUXO)

    'Le os FluxoSaldosIniciais.
    lErro = CF("FluxoSaldosIniciais_Le", colFluxoSldIni, objFluxo1.lFluxoId)
    If lErro <> SUCESSO And lErro <> 21141 Then gError 83797

    'preenche o grid com os valores lidos
    lErro = Preenche_GridFCaixa(colFluxoSint, colFluxoSldIni)
    If lErro <> SUCESSO Then gError 21079

    ExibeFluxoSint = SUCESSO
    
    Exit Function

Erro_ExibeFluxoSint:

    ExibeFluxoSint = gErr

    Select Case gErr

        Case 21079, 21090, 83797
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160472)

    End Select

    Exit Function

End Function

Private Sub ValorAjustadoTes_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorAjustadoTes_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorAjustadoTes_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorAjustadoTes
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21092
    
    Call Grid_Limpa(objGrid1)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21092

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160473)

    End Select

    Exit Sub

End Sub

Private Sub ValorAjustadoTes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Gravar_Registro() As Long
'grava os registros da tela

Dim lErro As Long
Dim colFluxoSint As New Collection
Dim objFluxoSint As ClassFluxoSint
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    For iLinha = 2 To objGrid1.iLinhasExistentes - 1

        Set objFluxoSint = New ClassFluxoSint
        
        objFluxoSint.lFluxoId = objFluxo1.lFluxoId
        objFluxoSint.dtData = CDate(GridFCaixa.TextMatrix(iLinha, GRID_DATA_COL))
        objFluxoSint.dTesValorAjustado = CDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_AJUSTADO_TESOURARIA_COL))
        objFluxoSint.dSaldoValorAjustado = CDbl(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_AJUSTADO_COL))

        colFluxoSint.Add objFluxoSint

    Next

    lErro = CF("FluxoSintetico_Grava", colFluxoSint, objFluxoSint.lFluxoId)
    If lErro <> SUCESSO Then Error 21098

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21098

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160474)

    End Select

    Exit Function

End Function

Private Sub Botao_ImprimeFluxo_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objFluxoSint As New ClassFluxoSint
Dim colFluxoSint As New Collection
Dim dtData As Date
Dim sOrdenados As String
Dim dPercRec As Double, dPercPag As Double, dPercTes As Double, dPercSaldo As Double
Dim dSaldoAcumuladoSist As Double
Dim dSaldoAcumuladoAjust As Double
Dim dSaldoInicialAjust As Double
Dim dSaldoInicialSist As Double

On Error GoTo Erro_Botao_ImprimeFluxo_Click
    
    lErro = objRelTela.Iniciar("REL_FLUXOSINTPROJ_CPR")
    If lErro <> SUCESSO Then Error 47922
    
    If objGrid1.iLinhasExistentes <> 0 Then
        dSaldoInicialAjust = StrParaDbl(GridFCaixa.TextMatrix(1, GRID_SALDO_AJUSTADO_COL))
        dSaldoInicialSist = StrParaDbl(GridFCaixa.TextMatrix(1, GRID_SALDO_SISTEMA_COL))
    End If
    
    dSaldoAcumuladoAjust = dSaldoInicialAjust
    dSaldoAcumuladoSist = dSaldoInicialSist
    
    lErro = Grid_FCaixa_Obter(colFluxoSint)
    If lErro <> SUCESSO Then Error 47923
    
    For iIndice1 = 1 To colFluxoSint.Count
    
        Set objFluxoSint = colFluxoSint.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        Call colTemp.Add(objFluxoSint.dtData)
        Call colTemp.Add(objFluxoSint.dRecValorSistema)
        Call colTemp.Add(objFluxoSint.dRecValorAjustado)
        dPercRec = PercentParaDbl(GridFCaixa.TextMatrix(iIndice1, GRID_VALOR_PERCENTUAL_RECEBER_COL))
        Call colTemp.Add(dPercRec)
        Call colTemp.Add(objFluxoSint.dPagValorSistema)
        Call colTemp.Add(objFluxoSint.dPagValorAjustado)
        
        dPercPag = PercentParaDbl(GridFCaixa.TextMatrix(iIndice1, GRID_VALOR_PERCENTUAL_PAGAMENTO_COL))
        Call colTemp.Add(dPercPag)
        Call colTemp.Add(objFluxoSint.dTesValorSistema)
        Call colTemp.Add(objFluxoSint.dTesValorAjustado)
        dPercTes = PercentParaDbl(GridFCaixa.TextMatrix(iIndice1, GRID_VALOR_PERCENTUAL_TESOURARIA_COL))
        Call colTemp.Add(dPercTes)
        Call colTemp.Add(objFluxoSint.dSaldoValorSistema)
        Call colTemp.Add(objFluxoSint.dSaldoValorAjustado)
        dPercSaldo = PercentParaDbl(GridFCaixa.TextMatrix(iIndice1, GRID_SALDO_PERCENTUAL_COL))
        Call colTemp.Add(dPercSaldo)
        
        dSaldoAcumuladoSist = dSaldoAcumuladoSist + objFluxoSint.dRecValorSistema + objFluxoSint.dTesValorSistema - objFluxoSint.dPagValorSistema
        Call colTemp.Add(dSaldoAcumuladoSist)
        dSaldoAcumuladoAjust = dSaldoAcumuladoAjust + objFluxoSint.dRecValorAjustado + objFluxoSint.dTesValorAjustado - objFluxoSint.dPagValorAjustado
        Call colTemp.Add(dSaldoAcumuladoAjust)
                
        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 47924
    
    Next
    
    lErro = objRelTela.ExecutarRel("", "TNOMEFLUXO", objFluxo1.sFluxo, "DDATABASE", objFluxo1.dtDataBase)
    If lErro <> SUCESSO Then Error 47925
    
    Exit Sub
    
Erro_Botao_ImprimeFluxo_Click:

    Select Case Err
          
        Case 47922, 47923, 47924, 47925
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160475)
     
    End Select

    Exit Sub

End Sub

Function Grid_FCaixa_Obter(colFluxoSint As Collection) As Long

Dim objFluxoSint As ClassFluxoSint
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_FCaixa_Obter

    For iLinha = 1 To objGrid1.iLinhasExistentes - 1

        Set objFluxoSint = New ClassFluxoSint
        
        'se nao for a linha do saldo inicial
        If iLinha <> 1 Then
            objFluxoSint.dtData = StrParaDate(GridFCaixa.TextMatrix(iLinha, GRID_DATA_COL))
        Else
            objFluxoSint.dtData = DATA_NULA
        End If
        
        objFluxoSint.dRecValorSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_SISTEMA_RECEBER_COL))
        objFluxoSint.dRecValorAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_AJUSTADO_RECEBER_COL))
        
        objFluxoSint.dPagValorSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_SISTEMA_PAGAMENTO_COL))
        objFluxoSint.dPagValorAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_AJUSTADO_PAGAMENTO_COL))
         
        objFluxoSint.dTesValorSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_SISTEMA_TESOURARIA_COL))
        objFluxoSint.dTesValorAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_VALOR_AJUSTADO_TESOURARIA_COL))
                
        objFluxoSint.dSaldoValorSistema = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_SISTEMA_COL))
        objFluxoSint.dSaldoValorAjustado = StrParaDbl(GridFCaixa.TextMatrix(iLinha, GRID_SALDO_AJUSTADO_COL))
    
        colFluxoSint.Add objFluxoSint
 
    Next
    
    Grid_FCaixa_Obter = SUCESSO
    
    Exit Function
    
Erro_Grid_FCaixa_Obter:

    Grid_FCaixa_Obter = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160476)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_SINTETICO_PROJETADO
    Set Form_Load_Ocx = Me
    Caption = "Fluxo Sintético - Projetado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoSintProj"
    
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

