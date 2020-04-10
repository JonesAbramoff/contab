VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoPag2Ocx 
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   ScaleHeight     =   6015
   ScaleWidth      =   9420
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   7080
      TabIndex        =   0
      Top             =   120
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   5070
      TabIndex        =   8
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSMask.MaskEdBox Titulo 
      Height          =   225
      Left            =   3240
      TabIndex        =   4
      Top             =   135
      Width           =   720
      _ExtentX        =   1270
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
      Left            =   2415
      TabIndex        =   2
      Top             =   135
      Width           =   525
      _ExtentX        =   926
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
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSMask.MaskEdBox Desconto 
      Height          =   225
      Left            =   7080
      TabIndex        =   11
      Top             =   2040
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
   Begin MSMask.MaskEdBox Juros 
      Height          =   225
      Left            =   6240
      TabIndex        =   10
      Top             =   1680
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
   Begin MSMask.MaskEdBox Multa 
      Height          =   225
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
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
   Begin MSMask.MaskEdBox ValorTotal 
      Height          =   225
      Left            =   5400
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.TextBox DataVencimento 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1035
   End
   Begin VB.TextBox DataEmissao 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   4290
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   3338
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2745
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "BorderoPag2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   150
         Picture         =   "BorderoPag2Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1117
         Picture         =   "BorderoPag2Ocx.ctx":08DC
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.TextBox Tipo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Height          =   555
      Left            =   7320
      Picture         =   "BorderoPag2Ocx.ctx":106E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionado"
      Height          =   1845
      Left            =   2760
      TabIndex        =   29
      Top             =   3400
      Width           =   4125
      Begin VB.CommandButton BotaoSalvarSelecao 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   435
         Width           =   1455
      End
      Begin VB.CommandButton BotaoTrazerSelecao 
         Caption         =   "Carregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label TotalTitulosSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   20
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Top             =   1185
         Width           =   510
      End
      Begin VB.Label QtdTitulosSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   19
         Top             =   495
         Width           =   780
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
         Left            =   150
         TabIndex        =   31
         Top             =   510
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Height          =   1845
      Left            =   120
      TabIndex        =   28
      Top             =   3400
      Width           =   2220
      Begin VB.Label QtdTitulos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   17
         Top             =   495
         Width           =   780
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
         Left            =   105
         TabIndex        =   32
         Top             =   510
         Width           =   540
      End
      Begin VB.Label TotalTitulos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   18
         Top             =   1185
         Width           =   1335
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
         Left            =   135
         TabIndex        =   33
         Top             =   1215
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todas"
      Height          =   555
      Left            =   7320
      Picture         =   "BorderoPag2Ocx.ctx":3F84
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1800
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   555
      Left            =   7320
      Picture         =   "BorderoPag2Ocx.ctx":4F9E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1800
   End
   Begin VB.CheckBox CheckPago 
      Height          =   255
      Left            =   6300
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   4140
      TabIndex        =   5
      Top             =   90
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      Format          =   "###"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GridBorderoPag2 
      Height          =   3000
      Left            =   120
      TabIndex        =   26
      Top             =   210
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   5292
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
End
Attribute VB_Name = "BorderoPag2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjBorderoPagEmissao As ClassBorderoPagEmissao

Dim objGrid As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Titulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_DataVencimento_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Pago_Col As Integer
Dim iGrid_Juros_Col As Integer
Dim iGrid_Multa_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_ValorTotal_Col As Integer

Const ERRO_DESCONTO_MAIOR_OU_IGUAL_VALOR_TOTAL = 0 'parâmetros: iLinha, sDesconto, sValorTotal
'Para o Item %s o Desconto %s não pode ser maior ou igual ao Valor do Título + Juros + Multa = %s.

Const AVISO_DESCONTO_TITULO_VENCIDO = 0
'Existem um ou mais títulos vencidos recebendo Desconto, deseja prosseguir?

Const AVISO_JUROSMULTA_TITULO_NAO_VENCIDO = 0
'Existem um ou mais títulos não vencidos recebendo Multa ou Juros, deseja prosseguir?

Private Sub BotaoFechar_Click()
    
    'Fecha a tela
    Unload Me
    
End Sub

Private Sub BotaoDesmarcar_Click()
'Desmarca todas as parcelas marcadas no Grid

Dim iLinha As Integer
Dim objInfoParcPag As New ClassInfoParcPag
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela a parcela em questão
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_NAO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(iLinha)
        
        'Desmarca no Obj a parcela em questão
        objInfoParcPag.iSeqCheque = 0
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Limpa na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(0)
    TotalTitulosSelecionados.Caption = CStr(Format(0, "Standard"))
    gobjBorderoPagEmissao.iQtdeParcelasSelecionadas = 0
    
End Sub

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objInfoParcPag As New ClassInfoParcPag
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridBorderoPag2.Row = 0 Then Error 60480
        
    'Se foi selecionada uma linha que está preenchida
    If GridBorderoPag2.Row <= objGrid.iLinhasExistentes Then
        
        Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(GridBorderoPag2.Row)
               
        objParcelaPagar.lNumIntDoc = objInfoParcPag.lNumIntParc
        
        'Le o NumInterno do Titulo para passar no objTituloPag
        lErro = CF("ParcelaPagar_Le", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 60479 Then Error 60481
        
        'Se não encontrou a Parcela --> ERRO
        If lErro = 60479 Then Error 60482
        
        objTituloPagar.lNumIntDoc = objParcelaPagar.lNumIntTitulo
        
        'Le os Dados do Titulo
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 60484
        
        If lErro = 18372 Then Error 60485
        
        Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case Err
    
        Case 60480
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
         
        Case 60484, 60481 'Tratado na rotina chamada
        
        Case 60482
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", Err)
        
        Case 60485
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143794)

    End Select

    Exit Sub
  
End Sub

Private Sub BotaoMarcar_Click()
'Marca todas as parcelas no Grid

Dim iLinha As Integer
Dim dTotalTitulosSelecionados As Double
Dim iNumTitulosSelecionados As Integer
Dim objInfoParcPag As ClassInfoParcPag

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela a parcela em questão
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(iLinha)
        
        'Marca no Obj a parcela em questão
        objInfoParcPag.iSeqCheque = 1
        
        dTotalTitulosSelecionados = dTotalTitulosSelecionados + CDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col))
        iNumTitulosSelecionados = iNumTitulosSelecionados + 1
    
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Atualiza na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(iNumTitulosSelecionados)
    TotalTitulosSelecionados.Caption = CStr(Format(dTotalTitulosSelecionados, "Standard"))
    gobjBorderoPagEmissao.iQtdeParcelasSelecionadas = iNumTitulosSelecionados
    
End Sub

Private Sub BotaoSeguir_Click()
 
Dim lErro As Long
Dim dValor As Double, objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer
Dim bDesconto As Boolean
Dim bMulta As Boolean

On Error GoTo Erro_BotaoSeguir_Click

    iLinha = 0
    
    'Ao menos uma parcela tem que estar marcada p/pagto
    If CInt(QtdTitulosSelecionados.Caption) = 0 Then gError 7781
    
    If Len(Trim(TotalTitulosSelecionados.Caption)) <> 0 Then
        dValor = CDbl(TotalTitulosSelecionados.Caption)
    End If
    
    If gobjBorderoPagEmissao.dValorParcelasSelecionadas <> dValor Then
        gobjBorderoPagEmissao.dValorParcelasSelecionadas = dValor
    End If
     
    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
        
        iLinha = iLinha + 1
        
        'Se a parcela foi marcada
        If GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then

            lErro = CF("TituloPagar_Verifica_Adiantamento", objInfoParcPag.lFornecedor, objInfoParcPag.iFilialForn)
            If lErro <> SUCESSO Then gError 59473
        
                
        End If
        
    Next
    
    'Alterado por Leo em 12/11/01 **************************
    'Alterado em função dos novos campos do grid (juros, multa, desconto e valor total)
    iLinha = 0
    
    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
        
        iLinha = iLinha + 1
        
        'Se a parcela foi marcada
        If GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then
                            
            'Trecho de código que serve para verificar se é um título atrasado ou em dia
            'caso atrasado, pergunta se realmente receberá desconto
            'caso em dia, pergunta se realmente será cobrado juros e/ou multa
            
            'Se o documento estiver vencido
            If StrParaDate(GridBorderoPag2.TextMatrix(iLinha, iGrid_DataVencimento_Col)) < gdtDataHoje Then
                
                'Se foi informado algum desconto diferente de zero e nenhuma mensagem a respeiro foi exibida
                If StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Desconto_Col)) > 0 And bDesconto = False Then
                
                    'Avisa que existe desconto p/ um documento vencido
                    lErro = Rotina_Aviso(vbYesNo, "AVISO_DESCONTO_TITULO_VENCIDO")
                    If lErro = vbYes Then
                        bDesconto = True
                    Else
                        Exit Sub
                    End If
                
                End If
            
            'Se o documento NÃO estiver vencido
            Else
                
                'Se foi informado alguma Multa ou Juros diferentes de zero e nenhuma mensagem a respeiro foi exibida

                If (StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Juros_Col)) > 0 _
                Or StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Multa_Col)) > 0) _
                And bMulta = False Then
            
                    'Avisa que existe juros ou multa para um documento NÃO vencido
                    lErro = Rotina_Aviso(vbYesNo, "AVISO_JUROSMULTA_TITULO_NAO_VENCIDO")
                    If lErro = vbYes Then
                        bMulta = True
                    Else
                        Exit Sub
                    End If
                
                End If
            
            End If
    
        End If
        
        'Se houveram avisos em relação a Desconto e a Juros/Multa abandona o Loop
        If bDesconto And bMulta Then Exit For
        
    Next
    
    'Leo até aqui *********************************
    
    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoPag3", gobjBorderoPagEmissao)
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoSeguir_Click:

    Select Case gErr

        Case 59473, 94232
        
        Case 7781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELA_SELECIONADA", Err)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143795)
            
    End Select

    Exit Sub

End Sub

'Maristela(inicio)

Private Sub BotaoSalvarSelecao_Click()
'Salva o status das parcelas do Grid

Dim lErro As Long

On Error GoTo Erro_BotaoSalvarSelecao_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'alterada a chamada da função BorderoPag_SalvarSelecao por leo em 09/11/01
    'Agora a função recebe como parâmetro uma coleção
    'Grava a nova selecao do Grid
     
    lErro = CF("BorderoPag_SalvarSelecao", gobjBorderoPagEmissao.colInfoParcPag)
    If lErro <> SUCESSO Then gError 90417
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoSalvarSelecao_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 90417

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143796)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerSelecao_Click()
'Recupera o status das Parcelas do Grid.

Dim lErro As Long
Dim objInfoParcPag As ClassInfoParcPag
Dim objInfoParcPagBD As ClassInfoParcPag
Dim colInfoParcPagBD As New Collection 'New inserido por Leo em 15/02/01
Dim iLinha As Integer
Dim dTotalTitulosSelecionados As Double
Dim iNumTitulosSelecionados As Integer
Dim dTotalDocGrid As Double

On Error GoTo Erro_BotaoTrazerSelecao_Click
            
    GL_objMDIForm.MousePointer = vbHourglass

    'Carrega na Coleção, o número e o status das Parcelas do Grid
    
    lErro = CF("BorderoPag_RecuperarSelecao", colInfoParcPagBD)
    If lErro <> SUCESSO Then gError 90418
    
    'Coleção do BD
    For Each objInfoParcPagBD In colInfoParcPagBD
        
       iLinha = 1
        
        'Coleção do Grid
        For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
        
            'Se houver parcelas iguais, verifica o status
            If objInfoParcPagBD.lNumIntParc = objInfoParcPag.lNumIntParc Then
                
                objInfoParcPag.iSeqCheque = objInfoParcPagBD.iSeqCheque
                
                'Alterado por Leo em 09/11/01 ************
                objInfoParcPag.dValorDesconto = objInfoParcPagBD.dValorDesconto
                objInfoParcPag.dValorJuros = objInfoParcPagBD.dValorJuros
                objInfoParcPag.dValorMulta = objInfoParcPagBD.dValorMulta
                'Leo até aqui**********
                
                'Marca a Check da parcela, de acordo com o status que está no BD
                GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = objInfoParcPag.iSeqCheque
                
                'Alterado por Leo em 09/11/01 ************
                
                'Preenche o campo Desconto do grid
                GridBorderoPag2.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(objInfoParcPag.dValorDesconto, "STANDARD")
                'Preenche o campo multa do grid
                GridBorderoPag2.TextMatrix(iLinha, iGrid_Multa_Col) = Format(objInfoParcPag.dValorMulta, "STANDARD")
                'Preenche o campo Juros do Grid
                GridBorderoPag2.TextMatrix(iLinha, iGrid_Juros_Col) = Format(objInfoParcPag.dValorJuros, "STANDARD")
                'Calcula a Preenche o campo Valor total da linha do grid
                GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format((objInfoParcPag.dValor + objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta) - objInfoParcPag.dValorDesconto, "STANDARD")
                
                'Leo até aqui **************
            End If
            
            iLinha = iLinha + 1
         
        Next
        
    Next
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(iLinha)
        
        If objInfoParcPag.iSeqCheque = 1 Then
            
            'Alterado por Leo em 09/11/01
            'Agora passará a levar em consideração o valor total e não mais o valor do título
            dTotalTitulosSelecionados = dTotalTitulosSelecionados + StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col))
            iNumTitulosSelecionados = iNumTitulosSelecionados + 1
    
        End If
    
        dTotalDocGrid = dTotalDocGrid + StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col))
        
    Next
        
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    'Atualiza na tela os campos Qtd de Títulos selecionados e Valor total dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(iNumTitulosSelecionados)
    'incluido por Leo em 13/11/01 o preenchimento do TotalTitulos
    TotalTitulos.Caption = Format(dTotalDocGrid, "STANDARD")
    TotalTitulosSelecionados.Caption = Format(dTotalTitulosSelecionados, "Standard")
    gobjBorderoPagEmissao.iQtdeParcelasSelecionadas = iNumTitulosSelecionados
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoTrazerSelecao_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 90418
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143797)
    
    End Select

    Exit Sub

End Sub
'Maristela(fim)

Private Sub BotaoVoltar_Click()

Dim lErro As Long
Dim dValor As Double

On Error GoTo Erro_BotaoVoltar_Click

    If Len(Trim(TotalTitulosSelecionados.Caption)) <> 0 Then
        dValor = CDbl(TotalTitulosSelecionados.Caption)
    End If
    
    If gobjBorderoPagEmissao.dValorParcelasSelecionadas <> dValor Then
        gobjBorderoPagEmissao.dValorParcelasSelecionadas = dValor
    End If
     
    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoPag1", gobjBorderoPagEmissao)
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoVoltar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143798)

    End Select

    Exit Sub

End Sub

Private Sub CheckPago_Click()

Dim iLinha As Integer
Dim dTotalTitulosSelecionados As Double
Dim iNumTitulosSelecionados As Integer
Dim objInfoParcPag As ClassInfoParcPag

    iLinha = GridBorderoPag2.Row
    
    'Passa a linha do Grid para o Obj
    Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(iLinha)
    
    'Passa para o Obj se a parcela em questão foi marcada ou desmarcada
    objInfoParcPag.iSeqCheque = CInt(GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col))
                    
    'Se a parcela foi marcada
    If GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then

        'Acrescenta a nova parcela no somatório de Qtd de Títulos selecionados e Valor total de Títulos selecionados
        gobjBorderoPagEmissao.iQtdeParcelasSelecionadas = gobjBorderoPagEmissao.iQtdeParcelasSelecionadas + 1
        QtdTitulosSelecionados.Caption = CStr(CInt(QtdTitulosSelecionados.Caption) + 1)
        TotalTitulosSelecionados.Caption = Format(StrParaDbl(TotalTitulosSelecionados.Caption) + StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col)), "Standard")
        'Linha acima alterada por Leo em 08/11/01 de: iGrid_Valor_Col, para: iGrid_ValorTotal_Col
        'A partir de hoje o total deverá ser o somatório dos valores totais
        
    Else
    
        'Subtrai a parcela do somatório de Qtd de Títulos selecionados e Valor total de Títulos selecionados
        gobjBorderoPagEmissao.iQtdeParcelasSelecionadas = gobjBorderoPagEmissao.iQtdeParcelasSelecionadas - 1
        QtdTitulosSelecionados.Caption = CStr(CInt(QtdTitulosSelecionados.Caption) - 1)
        TotalTitulosSelecionados.Caption = Format(StrParaDbl(TotalTitulosSelecionados.Caption) - StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col)), "Standard")
        'Linha acima alterada por Leo em 08/11/01 de: iGrid_Valor_Col, para: iGrid_ValorTotal_Col
        'A partir de hoje o total deverá ser o somatório dos valores totais

    End If

End Sub

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

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143799)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente.
'Alterada por Leo em 07/11/01
'Inclusão dos campos Juros, Multa e Desconto no Grid

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa variáveis para saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col
        
            'Caso a coluna seja Juros
            Case iGrid_Multa_Col
                lErro = Saida_Celula_Multa(objGridInt)
                If lErro <> SUCESSO Then gError 94219

            'Caso a coluna seja Multa
            Case iGrid_Desconto_Col
                lErro = Saida_Celula_Desconto(objGridInt)
                If lErro <> SUCESSO Then gError 94220

            'Caso a coluna seja Desconto
            Case iGrid_Juros_Col
                lErro = Saida_Celula_Juros(objGridInt)
                If lErro <> SUCESSO Then gError 94221

        End Select
        
        'Finaliza variáveis para saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 15790

    End If
       
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
                
        Case 15789, 15790, 94219, 94220, 94221
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143800)
        
    End Select

    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid = Nothing

    Set gobjBorderoPagEmissao = Nothing
    
End Sub

Private Sub Desconto_Change()
'leo em 07/11/01 *******

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_GotFocus()
'leo em 07/11/01 *******

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
'leo em 07/11/01 *******

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
'leo em 07/11/01 *******

Dim lErro As Long
    
    Set objGrid.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub GridBorderoPag2_Click()
        
Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection

    Call Grid_Click(objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If
    
    colcolColecao.Add gobjBorderoPagEmissao.colInfoParcPag
    
    Call Ordenacao_ClickGrid(objGrid, , colcolColecao)
        
End Sub

Private Sub GridBorderoPag2_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridBorderoPag2_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid, iAlterado)
    
End Sub

Private Sub GridBorderoPag2_LeaveCell()
 
    Call Saida_Celula(objGrid)
        
End Sub

Private Sub GridBorderoPag2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)
    
End Sub

Private Sub GridBorderoPag2_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridBorderoPag2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGrid.objTela = Me
        
        'Chama o Menu PopUp
        PopupMenu PopUpMenuGrid.mnuGrid, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuGrid.objTela = Nothing
        
    End If

End Sub

Private Sub GridBorderoPag2_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridBorderoPag2_RowColChange()

    Call Grid_RowColChange(objGrid)
       
End Sub

Private Sub GridBorderoPag2_Scroll()

    Call Grid_Scroll(objGrid)
    
End Sub

Private Function Inicializa_Grid_BorderoPag2(objGridInt As AdmGrid, iRegistros As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_BorderoPag2
    
    'Tela em questão
    Set objGrid.objForm = Me
    
    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Pagar")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Parc")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Vencimento")
    'leo em 07/11/01 *******
    objGridInt.colColuna.Add ("Valor Título")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Valor Total")
    objGridInt.colColuna.Add ("Filial Empresa")
    'leo até aqui *******
    
   'campos de edição do grid
    objGridInt.colCampo.Add (CheckPago.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Titulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    'leo em 07/11/01 *******
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Multa.Name)
    objGridInt.colCampo.Add (Juros.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    'leo até aqui *******

    
    iGrid_Pago_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_Filial_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_Titulo_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_DataEmissao_Col = 7
    iGrid_DataVencimento_Col = 8
    'leo em 07/11/01 *******
    iGrid_Valor_Col = 9
    iGrid_Multa_Col = 10
    iGrid_Juros_Col = 11
    iGrid_Desconto_Col = 12
    iGrid_ValorTotal_Col = 13
    iGrid_FilialEmpresa_Col = 14
    'leo até aqui *******
        
    objGridInt.objGrid = GridBorderoPag2
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9
        
    'Todas as linhas do grid
    If objGridInt.iLinhasVisiveis >= iRegistros + 1 Then
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    Else
        objGridInt.objGrid.Rows = iRegistros + 1
    End If
    
    GridBorderoPag2.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_BorderoPag2 = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_BorderoPag2:

    Inicializa_Grid_BorderoPag2 = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143801)
        
    End Select

    Exit Function
        
End Function

Function Trata_Parametros(Optional objBorderoPagEmissao As ClassBorderoPagEmissao) As Long
'Traz os dados das Parcelas a pagar para a Tela

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer, lErro As Long
Dim dTotal As Double
Dim dValorSelecionado As Double
Dim lQuantSelecionado As Long
Dim objFilialEmpresa As New AdmFiliais
Dim dValorTotal As Double

On Error GoTo Erro_Trata_Parametros

    Set gobjBorderoPagEmissao = objBorderoPagEmissao
    
    'Passa para a tela os dados dos Títulos selecionados
    QtdTitulosSelecionados.Caption = CStr(gobjBorderoPagEmissao.iQtdeParcelasSelecionadas)
    TotalTitulosSelecionados.Caption = CStr(Format(gobjBorderoPagEmissao.dValorParcelasSelecionadas, "Standard"))
    
    Set objGrid = New AdmGrid
            
    lErro = Inicializa_Grid_BorderoPag2(objGrid, gobjBorderoPagEmissao.colInfoParcPag.Count)
    If lErro <> SUCESSO Then gError 14250
    
    iLinha = 0

    'Percorre todas as parcelas da Coleção passada por parâmetro
    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag

        iLinha = iLinha + 1
        
        objFilialEmpresa.iCodFilial = objInfoParcPag.iFilialEmpresa
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 82781
        
        'Passa para a tela os dados da parcela em questão
        GridBorderoPag2.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objInfoParcPag.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objInfoParcPag.sRazaoSocialForn
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcPag.iFilialForn
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcPag.sSiglaDocumento
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Titulo_Col) = objInfoParcPag.lNumTitulo
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        
        'Alterado por Leo em 09/11/01 *************
        'Incluidos campos referentes ao Juros, Multa, Desconto e Valor Total.
        
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Juros_Col) = Format(objInfoParcPag.dValorJuros, "STANDARD")
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Multa_Col) = Format(objInfoParcPag.dValorMulta, "STANDARD")
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(objInfoParcPag.dValorDesconto, "STANDARD")
        dValorTotal = (objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta + objInfoParcPag.dValor) - objInfoParcPag.dValorDesconto
        GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(dValorTotal, "STANDARD")
        
        'Leo até aqui *************
        
        If objInfoParcPag.dtDataEmissao <> DATA_NULA Then GridBorderoPag2.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objInfoParcPag.dtDataEmissao, "dd/mm/yyyy")
        
        GridBorderoPag2.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/mm/yyyy")
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objInfoParcPag.dValor, "Standard")
        GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = objInfoParcPag.iSeqCheque
        
        If objInfoParcPag.iSeqCheque = 1 Then
            lQuantSelecionado = lQuantSelecionado + 1
            dValorSelecionado = dValorSelecionado + objInfoParcPag.dValor
        End If
        
        'Soma ao total o valor da parcela em questão
        dTotal = dTotal + dValorTotal

    Next

    'Passa para o Obj o número de parcelas passadas pela Coleção
    objGrid.iLinhasExistentes = iLinha
    
    'Passa para a tela o somatório da Qtd de Títulos e do Número total de títulos
    QtdTitulos.Caption = CStr(objGrid.iLinhasExistentes)
    TotalTitulos.Caption = CStr(Format(dTotal, "Standard"))
    TotalTitulosSelecionados = Format(dValorSelecionado, "Standard")
    QtdTitulosSelecionados = CStr(lQuantSelecionado)
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)
    
    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
     
    Select Case gErr
          
        Case 14250, 82781
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143802)
     
    End Select
     
    Exit Function

End Function

Public Sub mnuGridConsultaDocOriginal_Click()
'Chama a tela de consulta de Títulos a Receber quando essa opção for selecionada no grid
    Call BotaoDocOriginal_Click
End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_PAGT_P2
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Pagamento - Passo 2"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoPag2"
    
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


Private Sub Juros_Change()
'leo em 07/11/01 *******

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Juros_GotFocus()
'leo em 07/11/01 *******

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Juros_KeyPress(KeyAscii As Integer)
'leo em 07/11/01 *******

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Juros_Validate(Cancel As Boolean)
'leo em 07/11/01 *******

Dim lErro As Long
    
    Set objGrid.objControle = Juros
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Multa_Change()
'leo em 07/11/01 *******

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Multa_GotFocus()
'leo em 07/11/01 *******

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Multa_KeyPress(KeyAscii As Integer)
'leo em 07/11/01 *******

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Multa_Validate(Cancel As Boolean)
'leo em 07/11/01 *******

Dim lErro As Long
    
    Set objGrid.objControle = Multa
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

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

'***** fim do trecho a ser copiado ******

Private Sub TotalTitulosSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulosSelecionados, Source, X, Y)
End Sub

Private Sub TotalTitulosSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulosSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub QtdTitulosSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdTitulosSelecionados, Source, X, Y)
End Sub

Private Sub QtdTitulosSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdTitulosSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

'Leo daqui p/ baixo em 07/11/01
'Alterações referentes a inclusão dos campos Juros, multa e desconto no Grid.

Private Function Saida_Celula_Multa(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Multa que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Multa

    Set objGridInt.objControle = Multa

    If Len(Trim(Multa.Text)) = 0 Then Multa.Text = 0

    'verifica se o valor da multa informada não é negativo
    lErro = Valor_NaoNegativo_Critica(Multa.Text)
    If lErro <> SUCESSO Then gError 94227

    Multa.Text = Format(Multa.Text, "Standard")

    'verifica se o valor da multa foi alterado
    If Multa.Text <> GridBorderoPag2.TextMatrix(GridBorderoPag2.Row, iGrid_Multa_Col) Then
    
        'recalcula o valor total da linha do grid
        lErro = ValorTotal_Calcula(GridBorderoPag2.Row, StrParaDbl(Multa.Text))
        If lErro <> SUCESSO Then gError 94231

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 94228
    
    Saida_Celula_Multa = SUCESSO

    Exit Function

Erro_Saida_Celula_Multa:

    Saida_Celula_Multa = gErr


    Select Case gErr

        Case 94227, 94228, 94231
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143803)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Juros(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Juros que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Juros

    Set objGridInt.objControle = Juros

    If Len(Trim(Juros.Text)) = 0 Then Juros.Text = 0

    'Verifica se o Valor do Juros não é negativo
    lErro = Valor_NaoNegativo_Critica(Juros.Text)
    If lErro <> SUCESSO Then gError 94225

    Juros.Text = Format(Juros.Text, "Standard")
               
    'Verifica se o valor do Juros foi alterado
    If Juros.Text <> GridBorderoPag2.TextMatrix(GridBorderoPag2.Row, iGrid_Juros_Col) Then

        'recalcula o valor total da linha do grid
        lErro = ValorTotal_Calcula(GridBorderoPag2.Row, StrParaDbl(Juros.Text))
        If lErro <> SUCESSO Then gError 94230
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 94226
    
    Saida_Celula_Juros = SUCESSO

    Exit Function

Erro_Saida_Celula_Juros:

    Saida_Celula_Juros = gErr


    Select Case gErr

        Case 94225, 94226, 94230
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143804)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto
    
    If Len(Trim(Desconto.Text)) = 0 Then Desconto.Text = 0
    
    'Veririfica se o desconto não é um valor negativo
    lErro = Valor_NaoNegativo_Critica(Desconto.Text)
    If lErro <> SUCESSO Then gError 94223
                   
    Desconto.Text = Format(Desconto.Text, "STANDARD")
                           
    'verifica se o valor do desconto foi alterado
    If Desconto.Text <> GridBorderoPag2.TextMatrix(GridBorderoPag2.Row, iGrid_Desconto_Col) Then
    
        'recalcula o preço total da linha do grid
        lErro = ValorTotal_Calcula(GridBorderoPag2.Row, StrParaDbl(Desconto.Text))
        If lErro <> SUCESSO Then gError 94229
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 94224
        
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 94223, 94224, 94229
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143805)

    End Select

    Exit Function

End Function

Public Function ValorTotal_Calcula(iLinha As Integer, dCampo As Double) As Long
'Função que serve para atualizar o valor total dos documento do grid

Dim dJuros As Double
Dim dMulta As Double
Dim dDesconto As Double
Dim dValor As Double
Dim dValorTotalAux As Double
Dim dValorTotal As Double
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_ValorTotal_Calcula

    'lê os campos que armazenam valores do grid
    dValorTotal = StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col))
    dJuros = StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Juros_Col))
    dMulta = StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Multa_Col))
    dDesconto = StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Desconto_Col))
    dValor = StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col))
                                  
    'Seleciona o campo que esta com foco no grid
    Select Case GridBorderoPag2.Col

        'se o foco estiver no Juros
        Case iGrid_Juros_Col
            'Carrega a variável Juros com o valor recebido como parâmetro
            dJuros = dCampo
        
        'se o foco estiver na Multa
        Case iGrid_Multa_Col
            'Carrega a variável Multa com o valor recebido como parâmetro
            dMulta = dCampo
            
        'se o foco estiver no Desconto
        Case iGrid_Desconto_Col
            'Carrega a variável Desconto com o valor recebido como parâmetro
            dDesconto = dCampo
    
    End Select
                                  
    dValorTotalAux = (dJuros + dMulta + dValor)
    
    'Caso o Desconto maior ou igual ao Valor Total, Erro.
    If dDesconto >= dValorTotalAux Then gError 94222

    'Preenche o campo Valor Total com o novo valor total
    GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(dValorTotalAux - dDesconto, "STANDARD")
    
    'Se a linha do grid estiver selecionada na ChecBox
    If GridBorderoPag2.TextMatrix(iLinha, iGrid_Pago_Col) = PAGO_CHECADO Then
        
        'Atualiza o Valor Total dos camnpos selecionados da tela
        'Diminui o Valor Total anterior do  acumulador e em seguida soma o novo Valor Total
        TotalTitulosSelecionados.Caption = Format(StrParaDbl(TotalTitulosSelecionados.Caption) - dValorTotal, "STANDARD")
        TotalTitulosSelecionados.Caption = Format(StrParaDbl(TotalTitulosSelecionados.Caption) + StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col)), "Standard")
            
        'Atualiza o Valor Total no gobjBorderoPagEmissao
        gobjBorderoPagEmissao.dValorParcelasSelecionadas = StrParaDbl(TotalTitulosSelecionados.Caption)
        
    End If
        
    'Atualiza o Valor Total dos dos documentos do Grid
    'Diminui o anterior do  acumulador e em seguida soma o novo Valor Total
    TotalTitulos.Caption = Format((TotalTitulos.Caption) - dValorTotal, "STANDARD")
    TotalTitulos.Caption = Format(StrParaDbl(TotalTitulos.Caption) + StrParaDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_ValorTotal_Col)), "Standard")
            
    Set objInfoParcPag = gobjBorderoPagEmissao.colInfoParcPag.Item(iLinha)
    
    'Seleciona o controle com o foco
    Select Case GridBorderoPag2.Col

        'se o foco estiver no Juros
        Case iGrid_Juros_Col
            'Carrega a Coleção Global com o Juros Calculado
            objInfoParcPag.dValorJuros = dJuros
        
        'se o foco estiver na Multa
        Case iGrid_Multa_Col
            'Carrega a Coleção Global com a Multa Calculada
            objInfoParcPag.dValorMulta = dMulta
            
        'se o foco estiver no Desconto
        Case iGrid_Desconto_Col
            'Carrega a Coleção Global com o Desconto Calculado
            objInfoParcPag.dValorDesconto = dDesconto
    
    End Select
    
    Exit Function
    
    ValorTotal_Calcula = SUCESSO

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr
    
    Select Case gErr
              
       Case 94222
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_VALOR_TOTAL", gErr, GridBorderoPag2.Row, Format(dDesconto, "STANDARD"), Format(dValorTotalAux, "STANDARD"))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143806)
            
    End Select
            
    Exit Function

End Function

'Leo até aqui ***************************
