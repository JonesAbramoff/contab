VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LancamentoEstornoOcx 
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   9495
   Begin VB.CheckBox Gerencial 
      Height          =   210
      Left            =   6705
      TabIndex        =   27
      Tag             =   "1"
      Top             =   2325
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   6210
      ScaleHeight     =   765
      ScaleWidth      =   3045
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   3105
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   2580
         Picture         =   "LancamentoEstornoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   600
         Left            =   60
         Picture         =   "LancamentoEstornoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1200
      End
      Begin VB.CommandButton BotaoExtornar 
         Height          =   600
         Left            =   1350
         Picture         =   "LancamentoEstornoOcx.ctx":1F40
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   90
         Width           =   1155
      End
   End
   Begin VB.ComboBox Origem 
      Height          =   315
      Left            =   735
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   1695
   End
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   5310
      MaxLength       =   150
      TabIndex        =   7
      Top             =   1635
      Width           =   3300
   End
   Begin MSMask.MaskEdBox Debito 
      Height          =   225
      Left            =   4155
      TabIndex        =   6
      Top             =   1650
      Width           =   1080
      _ExtentX        =   1905
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
   Begin MSMask.MaskEdBox Credito 
      Height          =   225
      Left            =   3000
      TabIndex        =   5
      Top             =   1650
      Width           =   1080
      _ExtentX        =   1905
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
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   2280
      TabIndex        =   4
      Top             =   1650
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Conta 
      Height          =   225
      Left            =   975
      TabIndex        =   3
      Top             =   1635
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   510
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Documento 
      Height          =   285
      Left            =   5070
      TabIndex        =   1
      Top             =   135
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid GridLancamentos 
      Height          =   1860
      Left            =   255
      TabIndex        =   8
      Top             =   1335
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Lote 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   135
      Width           =   630
   End
   Begin VB.Label Label5 
      Caption         =   "Período:"
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
      Left            =   4305
      TabIndex        =   15
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Origem:"
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
      Height          =   255
      Left            =   45
      TabIndex        =   16
      Top             =   165
      Width           =   720
   End
   Begin VB.Label LoteLabel 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   17
      Top             =   180
      Width           =   450
   End
   Begin VB.Label DocumentoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Documento:"
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
      Left            =   4020
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   165
      Width           =   1020
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   210
      TabIndex        =   19
      Top             =   570
      Width           =   525
   End
   Begin VB.Label LabelTotais 
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
      Height          =   225
      Left            =   1815
      TabIndex        =   20
      Top             =   3315
      Width           =   705
   End
   Begin VB.Label TotalDebito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3870
      TabIndex        =   21
      Top             =   3315
      Width           =   1155
   End
   Begin VB.Label TotalCredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2625
      TabIndex        =   22
      Top             =   3315
      Width           =   1155
   End
   Begin VB.Label Periodo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5070
      TabIndex        =   23
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3105
      TabIndex        =   24
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "Exercício:"
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
      Left            =   2190
      TabIndex        =   25
      Top             =   570
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Lançamentos"
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
      Left            =   315
      TabIndex        =   26
      Top             =   1110
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -420
      X2              =   9690
      Y1              =   1005
      Y2              =   1005
   End
End
Attribute VB_Name = "LancamentoEstornoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_Gerencial_Col As Integer

Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

Private Sub BotaoExtornar_Click()

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objBrowseConfigura As New AdmBrowseConfigura
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialEmpresaSalva As Integer
Dim iAchou As Integer

On Error GoTo Erro_BotaoExtornar_Click

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then gError 41680

    'Se Documento estiver vazio, Erro
    If Len(Trim(Documento.Text)) = 0 Then gError 41681

    'Se Data estiver vazia, Erro
    If Len(Trim(Data.Text)) = 0 Then gError 41682

    lErro = Move_Tela_Memoria(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then gError 41683

    iFilialEmpresaSalva = objLancamento_Cabecalho.iFilialEmpresa
    
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100


        'Lê algum lançamento contido no documento em questão
        lErro = CF("Lancamento_Le_Doc1", objLancamento_Cabecalho)
        If lErro <> SUCESSO And lErro <> 83863 And lErro <> 83864 And lErro <> 83865 Then gError 41684
    
        If lErro <> 83863 Then iAchou = 1
    
        If lErro = 83864 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ESTORNO_DOCUMENTO_ESTORNADO")
            If vbMsgRes <> vbYes Then gError 83867
        
        
        End If
    
        If lErro = 83865 Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ESTORNO_DOCUMENTO_ESTORNADOR")
            If vbMsgRes <> vbYes Then gError 83868
    
        End If
    
        If giContabGerencial = 0 Then Exit Do
        
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar

    Loop
    
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresaSalva
    
    If iAchou = 0 Then gError 41685
    
    'chama a tela LancamentoEstorno1 Modal
    lErro = Chama_Tela_Modal("LancamentoEstorno1", objLancamento_Cabecalho, objBrowseConfigura)
    If lErro <> SUCESSO Then gError 41686

    If objBrowseConfigura.iTelaOK = CANCELA Then gError 41687

    Call Limpa_Tela_Lancamentos

    Exit Sub

Erro_BotaoExtornar_Click:

    Select Case gErr

        Case 41680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", gErr)
            Origem.SetFocus

        Case 41681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus

        Case 41682
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 41683, 41684, 41686, 83867, 83868

        Case 41685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", gErr, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case 41687
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_ESTORNO_LANCAMENTO_CANCELADO", objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162209)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Move_Tela_Memoria(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Obtém Período e Exercício correspondentes à data
    dtData = CDate(Data.Text)

    'Lê o Período
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 41679

    'Preenche objLote
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.lDoc = CLng(Documento.Text)
    objLancamento_Cabecalho.dtData = dtData

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 41679

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162210)

    End Select

    Exit Function

End Function

Private Sub BotaoConsultar_Click()

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_BotaoConsultar_Click

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 41654

    'Se Documento estiver vazio, Erro
    If Len(Documento.ClipText) = 0 Then Error 41655

    'Se Data estiver vazio ==> Erro
    If Len(Data.ClipText) = 0 Then Error 41656

    lErro = Move_Tela_Memoria(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 41657

    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 41658

    Exit Sub

Erro_BotaoConsultar_Click:

    Select Case Err

        Case 41654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", Err)
            Origem.SetFocus

        Case 41655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
            Documento.SetFocus

        Case 41656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)
            Data.SetFocus

        Case 41657, 41658

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162211)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim objLote As New ClassLote
Dim vbMsgRes As VbMsgBoxResult
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 41660

        'Obtém Período e Exercício correspondentes à data
        dtData = CDate(Data.Text)

        'Lê o Período
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41661

        'Lê o Exercício
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 41662

        'Exercício não cadastrado
        If lErro = 10083 Then Error 41663

        'Preenche campo de período
        Periodo.Caption = objPeriodo.sNomeExterno

        'Preenche campo de exercício
        Exercicio.Caption = objExercicio.sNomeExterno

    Else

        'Limpa os campos período e exercício
        Periodo.Caption = ""
        Exercicio.Caption = ""

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 41660, 41661, 41662

        Case 41663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162212)

    End Select

    Exit Sub

End Sub

Private Sub Documento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Documento)

End Sub

Private Sub DocumentoLabel_Click()

Dim lErro As Long
Dim dtData As Date
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_DocumentoLabel_Click

    If Len(Data.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41664

    Else

        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0

    End If

    If Len(Documento.Text) = 0 Then
        objLancamento_Detalhe.lDoc = 0
    Else
        objLancamento_Detalhe.lDoc = CLng(Documento.ClipText)
    End If

    objLancamento_Detalhe.iFilialEmpresa = giFilialEmpresa
    objLancamento_Detalhe.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLancamento_Detalhe.iExercicio = objPeriodo.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Detalhe.iPeriodoLote = objPeriodo.iPeriodo

    Call Chama_Tela("LancamentoLista", colSelecao, objLancamento_Detalhe, objEventoLancamento)

    Exit Sub

Erro_DocumentoLabel_Click:

    Select Case Err

        Case 41664

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162213)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Cabecalho() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sDescricao As String
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio


On Error GoTo Erro_Inicializa_Cabecalho

    Origem.Clear

    'Inicializar as Origens
    For iIndice = 1 To gobjColOrigem.Count
        Origem.AddItem gobjColOrigem.Item(iIndice).sDescricao
    Next

    sDescricao = gobjColOrigem.Descricao(gsOrigemAtual)

    'Mostra a Origem atual
    For iIndice = 0 To Origem.ListCount - 1
        If Origem.List(iIndice) = sDescricao Then
            Origem.ListIndex = iIndice
            Exit For
        End If
    Next

    'Inicializa Data
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 55720
    
    Periodo.Caption = objPeriodo.sNomeExterno
    
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 55721
    
    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then Error 55722
    
    Exercicio.Caption = objExercicio.sNomeExterno
        
    Inicializa_Cabecalho = SUCESSO

    Exit Function

Erro_Inicializa_Cabecalho:

    Inicializa_Cabecalho = Err

    Select Case Err
    
        Case 55720, 55721
            
        Case 55722
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162214)
            
    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoLancamento = New AdmEvento
    Set objGrid1 = New AdmGrid

    lErro = Inicializa_Grid_Lancamento(objGrid1)
    If lErro <> SUCESSO Then Error 41665

    GridLancamentos.Row = 1
    GridLancamentos.Col = 1

    lErro = Inicializa_Cabecalho()
    If lErro <> SUCESSO Then Error 41666

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41665, 41666

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162215)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Lancamento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Lancamento

    'Tela em questão
    Set objGrid1.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("Histórico")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")

   'Campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Debito.Name)
    objGridInt.colCampo.Add (Credito.Name)
    objGridInt.colCampo.Add (Historico.Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (Gerencial.Name)

    'Indica onde estão situadas as colunas do grid
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

    If giContabGerencial = 1 Then
        iGrid_Gerencial_Col = iGrid_Historico_Col + 1
    Else
        Gerencial.Visible = False
    End If

    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 41667

    objGridInt.objGrid = GridLancamentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
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

    LabelTotais.Top = TotalCredito.Top + (TotalCredito.Height - LabelTotais.Height) / 2
    LabelTotais.Left = TotalDebito.Left - LabelTotais.Width

    Inicializa_Grid_Lancamento = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Lancamento:

    Inicializa_Grid_Lancamento = Err

    Select Case Err

        Case 41667

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162216)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objLancamento_Detalhe As ClassLancamento_Detalhe) As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoLancamento = Nothing

    Set objGrid1 = Nothing
    
End Sub

Private Sub objEventoLancamento_evSelecao(obj1 As Object)
'Traz o lançamento selecionado para a tela

Dim lErro As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1

    objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
    objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
    objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc

    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 41668

    Me.Show

    Exit Sub

Erro_objEventoLancamento_evSelecao:

    Select Case Err

        Case 41668

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162217)

    End Select

    Exit Sub

End Sub

Private Function Traz_Doc_Tela(objDoc As ClassLancamento_Cabecalho) As Long
'Traz os dados do documento para a tela

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim sDescricao As String
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim dColunaSoma As Double
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim iFilialEmpresaSalva As Integer
Dim iAchou As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_Traz_Doc_Tela

    Call Limpa_Tela_Lancamentos

    iFilialEmpresaSalva = objDoc.iFilialEmpresa

    Do While objDoc.iFilialEmpresa > 0 And objDoc.iFilialEmpresa < 100

        'Lê os lançamentos contidos no documento em questão
        lErro = CF("Lancamentos_Le_Doc", objDoc, colLancamentos)
        If lErro <> SUCESSO And lErro <> 28700 Then Error 41669

        If lErro = SUCESSO Then iAchou = 1

        If giContabGerencial = 0 Then Exit Do

        objDoc.iFilialEmpresa = objDoc.iFilialEmpresa - giFilialAuxiliar

    Loop

    objDoc.iFilialEmpresa = iFilialEmpresaSalva
    
    'se não encontrou o documento
    If iAchou = 0 Then gError 41695
    
    For iIndice = colLancamentos.Count To 1 Step -1
        
        For iIndice1 = iIndice - 1 To 1 Step -1
        
            If colLancamentos(iIndice).iSeq = colLancamentos(iIndice1).iSeq Then
                colLancamentos.Remove (iIndice)
                Exit For
            End If
            
        Next
    
    Next

    Documento.Text = CStr(objDoc.lDoc)

    Set objLanc = colLancamentos.Item(1)

    Lote.Caption = CStr(objLanc.iLote)

    'Inicializa Data
    Data.Text = Format(objLanc.dtData, "dd/mm/yy")

    'Coloca o período relativo a data na tela
    lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 41670

    Periodo.Caption = objPeriodo.sNomeExterno

    'Coloca o exercício na tela
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 41671

    'Se o exercício não está cadastrado
    If lErro = 10083 Then Error 41672

    Exercicio.Caption = objExercicio.sNomeExterno

    sDescricao = gobjColOrigem.Descricao(objDoc.sOrigem)
    For iIndice = 0 To Origem.ListCount - 1
        If Origem.List(iIndice) = sDescricao Then
            Origem.ListIndex = iIndice
            Exit For
        End If
    Next

    If colLancamentos.Count > MAX_LANCAMENTOS_POR_DOC_CTB + 1 Then gError 197923
    
    If colLancamentos.Count >= objGrid1.objGrid.Rows Then
        Call Refaz_Grid(objGrid1, colLancamentos.Count)
    End If

    'Move os dados para a tela
    For Each objLanc In colLancamentos

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objLanc.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 41673

        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True

        'Coloca a conta na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = Conta.Text

        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

            If objLanc.sCcl <> "" Then

                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)

                lErro = Mascara_MascararCcl(objLanc.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 41674

            Else
            
                sCclMascarado = ""

            End If

            'Coloca o centro de custo na tela
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = sCclMascarado

        End If

        'Coloca o valor na tela
        If objLanc.dValor > 0 Then
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Credito_Col) = Format(objLanc.dValor, "Standard")
        Else
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Debito_Col) = Format(-objLanc.dValor, "Standard")
        End If

        'Coloca o histórico na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Historico_Col) = objLanc.sHistorico

        If giContabGerencial = 1 Then GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Gerencial_Col) = CStr(objLanc.iGerencial)

        objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1

    Next

    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dColunaSoma, "Standard")

    Call Grid_Refresh_Checkbox(objGrid1)

    Traz_Doc_Tela = SUCESSO

    Exit Function

Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err

        Case 41669, 41670, 41671

        Case 41672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 41673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objLanc.sConta)

        Case 41674
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objLanc.sCcl)
            
        Case 41695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objDoc.sOrigem, objDoc.iExercicio, objDoc.iPeriodoLan, objDoc.lDoc)
            
        Case 197923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_LANC_MAIOR_LIMITE", gErr, colLancamentos.Count, MAX_LANCAMENTOS_POR_DOC_CTB)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162218)

    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double

Dim iLinha As Integer
Dim dAcumulador As Double

    dAcumulador = 0

    For iLinha = 1 To objGrid1.iLinhasExistentes
        If Len(GridLancamentos.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(GridLancamentos.TextMatrix(iLinha, iColuna))
        End If
    Next

    GridColuna_Soma = dAcumulador

End Function

Function Limpa_Tela_Lancamentos() As Long

    Call Grid_Limpa(objGrid1)
    
    Lote.Caption = ""
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    Documento.Text = ""
        
    Limpa_Tela_Lancamentos = SUCESSO

End Function

Private Function Inicializa_Mascaras() As Long
'Inicializa as máscaras de conta e centro de custo

Dim lErro As Long
Dim sMascaraConta As String
Dim sMascaraCcl As String

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)

    'Lê a máscara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 41675

    Conta.Mask = sMascaraConta

    'Se usa centro de custo/lucro ==> inicializa máscara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        sMascaraCcl = String(STRING_CCL, 0)

        'Lê a máscara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 41676

        Ccl.Mask = sMascaraCcl

    End If

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 41675, 41676

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162219)

    End Select

    Exit Function

End Function

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 41677

        Data.Text = sData

    End If

    Call Data_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 41677

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162220)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 41678

        Data.Text = sData

    End If

    Call Data_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 41678

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162221)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "LancConsulta"
        
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then
        objLancamento_Cabecalho.dtData = DATA_NULA
    Else
        objLancamento_Cabecalho.dtData = CDate(Data.Text)
    End If
    
    'Lote
    If Len(Lote.Caption) = 0 Then
        objLancamento_Cabecalho.iLote = 0
    Else
        objLancamento_Cabecalho.iLote = CInt(Lote.Caption)
    End If
    
    'Documento
    If Len(Trim(Documento.ClipText)) = 0 Then
        objLancamento_Cabecalho.lDoc = 0
    Else
        objLancamento_Cabecalho.lDoc = CLng(Documento.ClipText)
    End If
    
    If Len(Trim(Exercicio.Caption)) > 0 Then objExercicio.sNomeExterno = Exercicio.Caption
    
    'Lê o Exercício
    lErro = CF("Exercicio_Le_Codigo", objExercicio)
    If lErro <> SUCESSO And lErro <> 28732 Then Error 59567
    
    If lErro = 28732 Then Error 59568
    
    If objExercicio.iExercicio <> 0 Then objPeriodo.iExercicio = objExercicio.iExercicio
    If Len(Trim(Periodo.Caption)) > 0 Then objPeriodo.sNomeExterno = Periodo.Caption
    
    'Lê o Período
    lErro = CF("Periodo_Le_Codigo", objPeriodo)
    If lErro <> SUCESSO And lErro <> 28736 Then Error 59569
    
    If lErro = 28736 Then Error 59570
    
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLancamento_Cabecalho.iExercicio = objExercicio.iExercicio
    objLancamento_Cabecalho.iPeriodoLote = objPeriodo.iPeriodo
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objLancamento_Cabecalho.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Origem", objLancamento_Cabecalho.sOrigem, STRING_ORIGEM, "Origem"
    colCampoValor.Add "Data", objLancamento_Cabecalho.dtData, 0, "Data"
    colCampoValor.Add "Lote", objLancamento_Cabecalho.iLote, 0, "Lote"
    colCampoValor.Add "Doc", objLancamento_Cabecalho.lDoc, 0, "Doc"
    colCampoValor.Add "Exercicio", objLancamento_Cabecalho.iExercicio, 0, "Exercicio"
    colCampoValor.Add "PeriodoLan", objLancamento_Cabecalho.iPeriodoLan, 0, "PeriodoLan"
    colCampoValor.Add "PeriodoLote", objLancamento_Cabecalho.iPeriodoLote, 0, "PeriodoLote"
   
    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case 59567, 59569
        
        Case 59568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_INEXISTENTE", Err, objExercicio.sNomeExterno)
            
        Case 59570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_EXERCICIO_INEXISTENTE", Err, objPeriodo.iExercicio, objPeriodo.sNomeExterno)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162222)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Tela_Preenche

    objLancamento_Cabecalho.dtData = colCampoValor.Item("Data").vValor

    If objLancamento_Cabecalho.dtData <> 0 Then
    
        objLancamento_Cabecalho.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objLancamento_Cabecalho.sOrigem = colCampoValor.Item("Origem").vValor
        objLancamento_Cabecalho.iLote = colCampoValor.Item("Lote").vValor
        objLancamento_Cabecalho.lDoc = colCampoValor.Item("Doc").vValor
        objLancamento_Cabecalho.iExercicio = colCampoValor.Item("Exercicio").vValor
        objLancamento_Cabecalho.iPeriodoLan = colCampoValor.Item("PeriodoLan").vValor
        objLancamento_Cabecalho.iPeriodoLote = colCampoValor.Item("PeriodoLote").vValor
        
        lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
        If lErro <> SUCESSO And lErro <> 5843 Then Error 59571

        If lErro = 5843 Then Error 59572

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 59571

        Case 59572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162223)

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

    Parent.HelpContextID = IDH_LANCAMENTO_EXTORNO
    Set Form_Load_Ocx = Me
    Caption = "Estorno de Documento Contábil"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LancamentoEstorno"
    
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
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Documento Then
            Call DocumentoLabel_Click
        End If
    
    End If

End Sub





Private Sub Lote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Lote, Source, X, Y)
End Sub

Private Sub Lote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Lote, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LoteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LoteLabel, Source, X, Y)
End Sub

Private Sub LoteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LoteLabel, Button, Shift, X, Y)
End Sub

Private Sub DocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DocumentoLabel, Source, X, Y)
End Sub

Private Sub DocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DocumentoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

Private Sub Periodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Periodo, Source, X, Y)
End Sub

Private Sub Periodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Periodo, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 10

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

