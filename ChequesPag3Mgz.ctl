VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequesPag3 
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   ScaleHeight     =   7110
   ScaleWidth      =   8955
   Begin VB.Frame FrameVerso 
      Caption         =   "Verso do Cheque"
      Height          =   1740
      Left            =   105
      TabIndex        =   20
      Top             =   2370
      Width           =   8715
      Begin VB.TextBox TextoVerso 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   975
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "ChequesPag3Mgz.ctx":0000
         Top             =   645
         Width           =   7140
      End
      Begin VB.Label LabelValorCheque 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label6 
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
         Height          =   240
         Left            =   1770
         TabIndex        =   26
         Top             =   285
         Width           =   540
      End
      Begin VB.Label LabelBenefCheque 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   270
         Width           =   3630
      End
      Begin VB.Label Label2 
         Caption         =   "Beneficiário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3765
         TabIndex        =   24
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label LabelNumCheque 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   22
         Top             =   285
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   3240
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6465
      Width           =   2670
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   75
         Picture         =   "ChequesPag3Mgz.ctx":0006
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2115
         Picture         =   "ChequesPag3Mgz.ctx":0764
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1110
         Picture         =   "ChequesPag3Mgz.ctx":08E2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Impressão de Cheques"
      Height          =   1860
      Left            =   1140
      TabIndex        =   15
      Top             =   4545
      Width           =   7065
      Begin VB.OptionButton OptionVerso 
         Caption         =   "Verso"
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
         Left            =   3645
         TabIndex        =   29
         Top             =   195
         Width           =   1005
      End
      Begin VB.OptionButton OptionFrente 
         Caption         =   "Frente"
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
         Left            =   2355
         TabIndex        =   28
         Top             =   195
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton BotaoImprimirAPartir 
         Caption         =   "Imprimir a Partir do Cheque Selecionado"
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
         Left            =   405
         TabIndex        =   10
         Top             =   1395
         Width           =   6240
      End
      Begin VB.CommandButton BotaoConfigurarImpressao 
         Caption         =   "Configurar Impressão..."
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
         Left            =   420
         TabIndex        =   6
         Top             =   540
         Width           =   3015
      End
      Begin VB.CommandButton BotaoImprimirSelecao 
         Caption         =   "Imprimir os  Cheques Selecionados"
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
         Left            =   420
         TabIndex        =   8
         Top             =   975
         Width           =   3375
      End
      Begin VB.CommandButton BotaoImprimirTeste 
         Caption         =   "Imprimir Teste"
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
         Left            =   3630
         TabIndex        =   7
         Top             =   540
         Width           =   3015
      End
      Begin VB.CommandButton BotaoImprimirTudo 
         Caption         =   "Imprimir Todos os Cheques"
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
         Left            =   3945
         TabIndex        =   9
         Top             =   960
         Width           =   2700
      End
   End
   Begin VB.CommandButton BotaoNumAuto 
      Caption         =   "Gerar numeração automática dos Cheques abaixo do Cheque selecionado"
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
      Left            =   1095
      TabIndex        =   5
      Top             =   4185
      Width           =   7140
   End
   Begin VB.CheckBox Atualizar 
      BackColor       =   &H80000005&
      Height          =   210
      Left            =   7440
      TabIndex        =   3
      Top             =   1125
      Width           =   1245
   End
   Begin MSMask.MaskEdBox Beneficiario 
      Height          =   225
      Left            =   3315
      TabIndex        =   2
      Top             =   1155
      Width           =   4200
      _ExtentX        =   7408
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
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   1890
      TabIndex        =   1
      Top             =   1125
      Width           =   1365
      _ExtentX        =   2408
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
   Begin MSMask.MaskEdBox Cheque 
      Height          =   225
      Left            =   855
      TabIndex        =   0
      Top             =   1125
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
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
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridChequesPag3 
      Height          =   1860
      Left            =   105
      TabIndex        =   4
      Top             =   510
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Qtde de Cheques:"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
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
      Left            =   1185
      TabIndex        =   17
      Top             =   180
      Width           =   1350
   End
   Begin VB.Label LabelConta 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   135
      Width           =   1995
   End
   Begin VB.Label LabelQtdCheques 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7170
      TabIndex        =   19
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "ChequesPag3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iGrid_Cheque_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Beneficiario_Col As Integer
Dim iGrid_Atualizar_Col As Integer
Dim objGridChequesPag3 As AdmGrid
Dim gobjChequesPag As ClassChequesPag

'30/11/01 Marcelo
Dim iChequeImpresso As Integer

Private Sub Atualizar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Private Sub Atualizar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Private Sub Atualizar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Atualizar
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub BotaoConfigurarImpressao_Click()

    Call Sist_ImpressoraDlg(1)

End Sub

Private Sub BotaoFechar_Click()
    
    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoImprimirAPartir_Click()

Dim iIndice As Integer, objInfoChequePag As ClassInfoChequePag
Dim iLinhaInicial As Integer, iLinha As Integer, lErro As Long

On Error GoTo Erro_BotaoImprimirAPartir_Click

    iLinhaInicial = 0
   
   'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridChequesPag3.iLinhasExistentes
        If GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
            iLinhaInicial = iLinha
            Exit For
        End If
    Next
    
    'uma linha tem que estar selecionada
    If iLinhaInicial = 0 Then Error 32124
    
    'desmarcar todos acima e marcar todos c/indice a partir da 1a linha marcada
    
    'percorre a colecao marcando todos os cheques selecionados e desmarcando os outros
    For iIndice = 1 To gobjChequesPag.ColInfoChequePag.Count
    
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iIndice)
        If iLinhaInicial <= iIndice Then
            objInfoChequePag.iImprimir = 1
        Else
            objInfoChequePag.iImprimir = 0
        End If
    Next
    
    'imprimir os cheques marcados na colecao
    Call ImprimeChequesSelecionados
    
    Exit Sub
    
Erro_BotaoImprimirAPartir_Click:

    Select Case Err

        Case 32124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_CHEQUES_MARCADOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub
        
End Sub

Private Sub BotaoImprimirSelecao_Click()
    
Dim iIndice As Integer, objInfoChequePag As ClassInfoChequePag, iQtde As Integer

On Error GoTo Erro_BotaoImprimirSelecao_Click

    iQtde = 0
    
    'percorre a colecao marcando todos os cheques selecionados e desmarcando os outros
    For iIndice = 1 To gobjChequesPag.ColInfoChequePag.Count
    
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iIndice)
        If GridChequesPag3.TextMatrix(iIndice, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
            objInfoChequePag.iImprimir = 1
            iQtde = iQtde + 1
        Else
            objInfoChequePag.iImprimir = 0
        End If
    Next
    
    If iQtde > 0 Then
    
        'imprimir os cheques marcados na colecao
        Call ImprimeChequesSelecionados
        
    Else
    
        Error 56756
    
    End If

    Exit Sub
     
Erro_BotaoImprimirSelecao_Click:

    Select Case Err
          
        Case 56756
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECIONE_CHEQUES_NO_GRID", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoImprimirTeste_Click()

Dim lErro As Long, lNumImpressao As Long

On Error GoTo Erro_ImprimirTeste_Click

    lErro = CF("Cheques_PrepararTesteImpressao", lNumImpressao)
    If lErro <> SUCESSO Then Error 19464

    gobjChequesPag.lNumImpressao = lNumImpressao

    lErro = ImprimirCheques(lNumImpressao, gobjChequesPag.sLayoutCheque, gobjChequesPag.dtEmissao)
    If lErro <> SUCESSO Then Error 19465

    Exit Sub
    
Erro_ImprimirTeste_Click:

    Select Case Err

        Case 19464, 19465
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoImprimirTudo_Click()
    
Dim objInfoChequePag As ClassInfoChequePag

    'percorre a colecao marcando tudo
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag
    
        objInfoChequePag.iImprimir = 1
        
    Next
    
    'imprime os cheques marcados na colecao
    Call ImprimeChequesSelecionados
    
End Sub

Private Sub BotaoNumAuto_Click()
'Gera numeração automática de cheques a partir do cheque selecionado no grid

Dim lErro As Long
Dim iNumParcelasMarcadas As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim iLinhaMarcada As Integer
Dim lChequeMarcado As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_BotaoNumAuto_Click
    
   lChequeMarcado = 0
   iLinhaMarcada = 0
   iNumParcelasMarcadas = 0
   
   'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridChequesPag3.iLinhasExistentes
        
        'Se o Cheque está marcado
        If GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_CHECADO Then
    
            'Passa a linha do Grid para o Obj
            Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iLinha)
            
            'Lê os dados do Cheque marcado
            lChequeMarcado = objInfoChequePag.lNumRealCheque
            iLinhaMarcada = iLinha
            
            'Faz o somatório do número de Cheques marcadas
            iNumParcelasMarcadas = iNumParcelasMarcadas + 1
            
            'Desmarca o Cheque
            GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_NAO_CHECADO
            
        End If
        
    Next
    
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridChequesPag3)
    
    'Verifica se há um número diferente de 1 de Cheques marcados
    If iNumParcelasMarcadas <> 1 Then Error 15876
    
    'Verifica se o Cheque marcado é o último do Grid
    If iLinhaMarcada = objGridChequesPag3.iLinhasExistentes Then Exit Sub
    
    iLinha = iLinhaMarcada
    iIndice = 0
   
    'Percorre todos os Cheques a partir do que foi selecionado
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag

        iIndice = iIndice + 1
        iLinha = iLinha + 1

        If iLinha > objGridChequesPag3.iLinhasExistentes Then Exit For
        
        'Passa a linha do Grid para o Obj
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(iLinha)
            
        'Altera sequencialmente a numeração do Cheque
        objInfoChequePag.lNumRealCheque = lChequeMarcado + CLng(iIndice)
        GridChequesPag3.TextMatrix(iLinha, iGrid_Cheque_Col) = objInfoChequePag.lNumRealCheque
        
    Next
                    
    Exit Sub
    
Erro_BotaoNumAuto_Click:

    Select Case Err
    
        Case 15876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_CHEQUES_MARCADOS", Err)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoSeguir_Click()

'30/11/01 Marcelo inicio
Dim vbMsgRes As VbMsgBoxResult
    
    If iChequeImpresso = 0 Then
    
        'perguntar se prossegue mesmo sem ter impresso
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NAO_IMP_CHEQUE")
        If vbMsgRes <> vbYes Then Exit Sub
        
    End If
    
'30/11/01 Marcelo fim

    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequesPag4", gobjChequesPag)
    
    'Fecha a tela
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
        
        
    '30/11/01 Marcelo
    iChequeImpresso = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
    
    End Select
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do Grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Set gObjChequesPag = objChequesPag
    
    'Chama rotina de inicialização da saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt.objGrid Is GridChequesPag3 Then

            Select Case objGridInt.objGrid.Col
            
                'Se a célula for o campo Cheque
                Case iGrid_Cheque_Col
                    
                    Set objGridInt.objControle = Cheque
                    
                   'Chama função de tratamento de saída da célula Cheque
                    lErro = Saida_Celula_Cheque(objGridInt)
                    If lErro <> SUCESSO Then Error 15874
                    
            End Select

        End If
        
        'Chama função de finalização da saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 15875

    End If

    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
    
        Case 15874, 15875
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
        
    End Select

    Exit Function

End Function

Private Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("ChequesPag2", gobjChequesPag)

    'Fecha a tela
    Unload Me
    
End Sub

Private Sub Cheque_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Private Sub Cheque_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Private Sub Cheque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Cheque
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridChequesPag3 = Nothing

    Set gobjChequesPag = Nothing
    
End Sub

Private Sub TextoVerso_Validate(Cancel As Boolean)
Dim objInfoChequePag As ClassInfoChequePag

    If Not (gobjChequesPag Is Nothing) Then
    
        If GridChequesPag3.Row >= 1 And GridChequesPag3.Row <= objGridChequesPag3.iLinhasExistentes Then
            
            'Passa os dados da linha do Grid para o Obj
            Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(GridChequesPag3.Row)
            
            objInfoChequePag.sVerso = TextoVerso.Text
        
        End If
    
    End If
    
End Sub

Private Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)
      
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub Beneficiario_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridChequesPag3)

End Sub

Private Sub Beneficiario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequesPag3)
    
End Sub

Private Sub Beneficiario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequesPag3.objControle = Beneficiario
    lErro = Grid_Campo_Libera_Foco(objGridChequesPag3)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridChequesPag3_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridChequesPag3, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    End If
    
End Sub

Private Sub GridChequesPag3_GotFocus()
    
    Call Grid_Recebe_Foco(objGridChequesPag3)

End Sub

Private Sub GridChequesPag3_EnterCell()
    
    Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    
End Sub

Private Sub GridChequesPag3_LeaveCell()
    
    Call Saida_Celula(objGridChequesPag3)
    
End Sub

Private Sub GridChequesPag3_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridChequesPag3)
    
End Sub

Private Sub GridChequesPag3_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridChequesPag3, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequesPag3, iAlterado)
    End If

End Sub

Private Sub GridChequesPag3_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridChequesPag3)

End Sub

Private Sub GridChequesPag3_RowColChange()
Dim objInfoChequePag As ClassInfoChequePag

    Call Grid_RowColChange(objGridChequesPag3)
       
    If Not (gobjChequesPag Is Nothing) Then
    
        If GridChequesPag3.Row >= 1 And GridChequesPag3.Row <= objGridChequesPag3.iLinhasExistentes Then
            
            'Passa os dados da linha do Grid para o Obj
            Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(GridChequesPag3.Row)
            
            LabelNumCheque.Caption = CStr(objInfoChequePag.lNumRealCheque)
            LabelValorCheque = Format(objInfoChequePag.dValor, "Standard")
            LabelBenefCheque = objInfoChequePag.sFavorecido
            TextoVerso.Text = objInfoChequePag.sVerso
        
        End If
    
    End If

End Sub

Private Sub GridChequesPag3_Scroll()

    Call Grid_Scroll(objGridChequesPag3)
    
End Sub

Private Function Inicializa_Grid_ChequesPag3(objGridInt As AdmGrid, iNumCheques As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_ChequesPag3
    
    'tela em questão
    Set objGridChequesPag3.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cheque")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Beneficiário")
    objGridInt.colColuna.Add ("Selecionar")
        
   'campos de edição do grid
    objGridInt.colCampo.Add (Cheque.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Beneficiario.Name)
    objGridInt.colCampo.Add (Atualizar.Name)
    
    iGrid_Cheque_Col = 1
    iGrid_Valor_Col = 2
    iGrid_Beneficiario_Col = 3
    iGrid_Atualizar_Col = 4
        
    objGridInt.objGrid = GridChequesPag3
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    'todas as linhas do grid
    If iNumCheques > objGridInt.iLinhasVisiveis Then
        objGridInt.objGrid.Rows = iNumCheques + 1
    Else
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    End If
    
    GridChequesPag3.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ChequesPag3 = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_ChequesPag3:

    Inicializa_Grid_ChequesPag3 = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
        
    End Select

    Exit Function
        
End Function

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long
'Traz para a tela os dados dos Cheques marcados para emissão

Dim objInfoChequePag As ClassInfoChequePag
Dim iLinha As Integer, lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjChequesPag = objChequesPag
    
    'Passa a Conta Corrente para a tela
    LabelConta.Caption = CStr(gobjChequesPag.iCta)
    
    Set objGridChequesPag3 = New AdmGrid
        
    lErro = Inicializa_Grid_ChequesPag3(objGridChequesPag3, gobjChequesPag.ColInfoChequePag.Count)
    If lErro <> SUCESSO Then Error 19376
    
    iLinha = 0
   
    'Percorre todos os Cheques da Coleção passada por parâmetro
    For Each objInfoChequePag In gobjChequesPag.ColInfoChequePag

        iLinha = iLinha + 1

        'Passa para a tela os dados do Cheque em questão
        GridChequesPag3.TextMatrix(iLinha, iGrid_Cheque_Col) = objInfoChequePag.lNumRealCheque
        GridChequesPag3.TextMatrix(iLinha, iGrid_Valor_Col) = CStr(Format(objInfoChequePag.dValor, "Standard"))
        GridChequesPag3.TextMatrix(iLinha, iGrid_Beneficiario_Col) = objInfoChequePag.sFavorecido
        GridChequesPag3.TextMatrix(iLinha, iGrid_Atualizar_Col) = ATUALIZAR_NAO_CHECADO
        
    Next

    'Passa para o Obj o número de Cheques passados pela Coleção
    objGridChequesPag3.iLinhasExistentes = iLinha
    
    'Se o número de Cheques for maior que o número de linhas do Grid
    If iLinha + 1 > GridChequesPag3.Rows Then
        
        'Altera o número de linhas do Grid de acordo com o número de Cheques
        GridChequesPag3.Rows = iLinha + 1
    
    End If
    
    'Passa para a tela a Qtd de Cheques
    LabelQtdCheques.Caption = CStr(objGridChequesPag3.iLinhasExistentes)
        
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridChequesPag3)
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 19376
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Function Saida_Celula_Cheque(objGridInt As AdmGrid) As Long

Dim objInfoChequePag As ClassInfoChequePag
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Cheque

    If Not (gobjChequesPag Is Nothing) Then
    
        'Passa os dados da linha do Grid para o Obj
        Set objInfoChequePag = gobjChequesPag.ColInfoChequePag.Item(GridChequesPag3.Row)
            
        If Len(Trim(Cheque.Text)) <> 0 Then
        
            'Passa para o Obj o valor do Cheque que está na tela
            objInfoChequePag.lNumRealCheque = CLng(Trim(Cheque.Text))
            
        Else
        
            'Inicializa o Cheque no Obj
            objInfoChequePag.lNumRealCheque = 0
            
        End If
        
    End If
    
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15877
    
    Saida_Celula_Cheque = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Cheque:

    Saida_Celula_Cheque = Err
    
    Select Case Err

        Case 15877
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
            
    End Select
    
    Exit Function
    
End Function

Private Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19467
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade, "TIGNORARMARGEM", "S", "TFRENTE", IIf(OptionFrente.Value, "S", "N"))
    If lErro <> SUCESSO Then Error 7431

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = Err

    Select Case Err

        Case 7431, 19467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Sub ImprimeChequesSelecionados()

Dim lErro As Long

On Error GoTo Erro_ImprimirTeste_Click

    lErro = CF("ChequesPag_PrepararImpressao", gobjChequesPag)
    If lErro <> SUCESSO Then Error 19307

    lErro = ImprimirCheques(gobjChequesPag.lNumImpressao, gobjChequesPag.sLayoutCheque, gobjChequesPag.dtEmissao)
    If lErro <> SUCESSO Then Error 19308
    
    '30/11/01 Marcelo
    iChequeImpresso = 1
    
    Exit Sub
    
Erro_ImprimirTeste_Click:

    Select Case Err

        Case 19307, 19308
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_IMPRESSAO_CHEQUES_P3
    Set Form_Load_Ocx = Me
    Caption = "Impressão de Cheques - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequesPag3"
    
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



Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConta, Source, X, Y)
End Sub

Private Sub LabelConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConta, Button, Shift, X, Y)
End Sub

Private Sub LabelQtdCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQtdCheques, Source, X, Y)
End Sub

Private Sub LabelQtdCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQtdCheques, Button, Shift, X, Y)
End Sub

