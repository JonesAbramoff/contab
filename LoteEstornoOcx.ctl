VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LoteEstornoOcx 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   9240
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   5700
      ScaleHeight     =   690
      ScaleWidth      =   3345
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   3405
      Begin VB.CommandButton BotaoFechar 
         Height          =   525
         Left            =   2820
         Picture         =   "LoteEstornoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   525
         Left            =   45
         Picture         =   "LoteEstornoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   1230
      End
      Begin VB.CommandButton BotaoExtornar 
         Height          =   525
         Left            =   1425
         Picture         =   "LoteEstornoOcx.ctx":1F40
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.ComboBox Origem 
      Height          =   315
      Left            =   3570
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   675
      Width           =   2010
   End
   Begin VB.ComboBox Periodo 
      Height          =   315
      Left            =   3585
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1980
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1695
   End
   Begin VB.CommandButton LancamentoLote 
      Caption         =   "Lançamentos do Lote"
      Height          =   765
      Left            =   6945
      Picture         =   "LoteEstornoOcx.ctx":3B82
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3735
      Width           =   2100
   End
   Begin VB.Frame SSFrame1 
      Caption         =   "Valores Atuais"
      Height          =   1650
      Left            =   120
      TabIndex        =   9
      Top             =   1905
      Width           =   8910
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número de Documentos:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1260
         Width           =   2100
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Débitos:"
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
         Left            =   360
         TabIndex        =   11
         Top             =   855
         Width           =   2070
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Créditos:"
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
         Left            =   300
         TabIndex        =   12
         Top             =   435
         Width           =   2115
      End
      Begin VB.Label TotCredAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   13
         Top             =   375
         Width           =   1575
      End
      Begin VB.Label TotDebAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   14
         Top             =   795
         Width           =   1575
      End
      Begin VB.Label NumLancAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Left            =   2475
         TabIndex        =   15
         Top             =   1230
         Width           =   615
      End
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   645
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   2865
      TabIndex        =   25
      Top             =   780
      Width           =   660
   End
   Begin VB.Label LabelLote 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   510
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   24
      Top             =   735
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   75
      TabIndex        =   23
      Top             =   210
      Width           =   885
   End
   Begin VB.Label IdLoteExterno 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2730
      TabIndex        =   16
      Top             =   3945
      Width           =   1530
   End
   Begin VB.Label TotLancInf 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   17
      Top             =   1290
      Width           =   1530
   End
   Begin VB.Label NumLancInf 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8025
      TabIndex        =   18
      Top             =   1305
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Identificador de Lote Externo:"
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
      Left            =   150
      TabIndex        =   19
      Top             =   3990
      Width           =   2550
   End
   Begin VB.Label Label6 
      Caption         =   "Número de Documentos:"
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
      Left            =   5820
      TabIndex        =   20
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total dos Documentos:"
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
      TabIndex        =   21
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2775
      TabIndex        =   22
      Top             =   255
      Width           =   750
   End
End
Attribute VB_Name = "LoteEstornoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Traz_Dados_Tela(objLote As ClassLote)

Dim iIndice As Integer

    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = objLote.iPeriodo Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next

    If objLote.dTotInf > 0 Then
        TotLancInf.Caption = Format(objLote.dTotInf, "Fixed")
    Else
        TotLancInf.Caption = ""
    End If
    If objLote.iNumDocInf > 0 Then
        NumLancInf.Caption = CStr(objLote.iNumDocInf)
    Else
        NumLancInf.Caption = ""
    End If
    If objLote.dTotCre > 0 Then
        TotCredAtual.Caption = Format(objLote.dTotCre, "Fixed")
    Else
        TotCredAtual.Caption = "0,00"
    End If
    If objLote.dTotDeb > 0 Then
        TotDebAtual.Caption = Format(objLote.dTotDeb, "Fixed")
    Else
        TotDebAtual.Caption = "0,00"
    End If
    If objLote.iNumDocAtual > 0 Then
        NumLancAtual.Caption = CStr(objLote.iNumDocAtual)
    Else
        NumLancAtual.Caption = "0"
    End If
    If objLote.sIdOriginal <> "" Then
        IdLoteExterno.Caption = objLote.sIdOriginal
    Else
        IdLoteExterno.Caption = ""
    End If

End Sub

Private Sub Limpa_Labels()

    TotLancInf.Caption = ""
    NumLancInf.Caption = ""
    TotCredAtual.Caption = "0,00"
    TotDebAtual.Caption = "0,00"
    NumLancAtual.Caption = "0"
    IdLoteExterno.Caption = ""
    
End Sub

Private Sub BotaoConsultar_Click()

Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_BotaoConsultar_Click

    'Se Exercicio estiver vazio, Erro
    If Len(Trim(Exercicio.Text)) = 0 Then Error 41610

    'Se Periodo estiver vazio, Erro
    If Len(Trim(Periodo.Text)) = 0 Then Error 41611

    'Se Lote estiver vazio, Erro
    If Len(Trim(Lote.Text)) = 0 Then Error 41612

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 41613

    Call Move_Tela_Memoria(objLote)

    'Lê o Lote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 9293 Then Error 41614

    If lErro = 9293 Then Error 41615

    'Mostra os dados na tela
    Call Traz_Dados_Tela(objLote)

    Exit Sub

Erro_BotaoConsultar_Click:

    Select Case Err

        Case 41610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)

        Case 41611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)

        Case 41612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 41613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", Err)

        Case 41614

        Case 41615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            Call Limpa_Labels

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162483)

    End Select

    Exit Sub

End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iExercicio As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_Exercicio_Click

    If Exercicio.ListIndex = -1 Then Exit Sub

    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)

    'inicializar os periodos do exercicio atual
    lErro = CF("Periodo_Le_Todos_Exercicio",giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 41616

    Periodo.Clear

    For iIndice = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iIndice)
        Periodo.AddItem objPeriodo.sNomeExterno
        Periodo.ItemData(Periodo.NewIndex) = objPeriodo.iPeriodo
    Next

    'Seleciona o primeiro periodo da combobox
    Periodo.ListIndex = 0

    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 41616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162484)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Cabecalho() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sDescricao As String
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection


On Error GoTo Erro_Inicializa_Cabecalho

    Origem.Clear
    
    'Inicializar as Origens
    For iIndice = 1 To gobjColOrigem.Count
        Origem.AddItem gobjColOrigem.Item(iIndice).sDescricao
    Next

    sDescricao = gobjColOrigem.Descricao(gsOrigemAtual)

    'Mostra a Origem atual
    For iIndice = 0 To Origem.ListCount - 1
        Origem.ListIndex = iIndice
        If Origem.Text = sDescricao Then Exit For
    Next

    'Ler todos os Exercícios
    lErro = CF("Exercicios_Le_Todos",colExercicios)
    If lErro <> SUCESSO Then Error 41617

    Exercicio.Clear
    
    For iIndice = 1 To colExercicios.Count
        Set objExercicio = colExercicios.Item(iIndice)
        Exercicio.AddItem objExercicio.sNomeExterno
        Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
    Next

    'Mostra o Exercício atual
    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = giExercicioAtual Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    'Mostra o Período atual
    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = giPeriodoAtual Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next

    Inicializa_Cabecalho = SUCESSO

    Exit Function

Erro_Inicializa_Cabecalho:

    Inicializa_Cabecalho = Err

    Select Case Err

        Case 41617

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162485)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoLote = New AdmEvento
    Set objEventoLancamento = New AdmEvento

    lErro = Inicializa_Cabecalho()
    If lErro <> SUCESSO Then Error 41624

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41624

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162486)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoLote = Nothing
    Set objEventoLancamento = Nothing

End Sub

Private Sub LabelLote_Click()

Dim objLote As New ClassLote
Dim colSelecao As New Collection

    If Len(Trim(Lote.ClipText)) <> 0 Then objLote.iLote = CInt(Lote.Text)

    '???? ESTE BROWSER DEVE LISTAR SOMENTE OS LOTES CONTABILIZADOS
    'Falta implementar
    Call Chama_Tela("LoteLista", colSelecao, objLote, objEventoLote)

End Sub

Private Sub LancamentoLote_Click()

Dim lErro As Long
Dim objLote As New ClassLote
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim colSelecao As New Collection

On Error GoTo Erro_LancamentoLote_Click

    'Se Exercicio estiver vazio, Erro
    If Len(Trim(Exercicio.Text)) = 0 Then Error 41618

    'Se Periodo estiver vazio, Erro
    If Len(Trim(Periodo.Text)) = 0 Then Error 41619

    'Se Lote estiver vazio, Erro
    If Len(Trim(Lote.Text)) = 0 Then Error 41620

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 41621

    'Preenche objLote
    Call Move_Tela_Memoria(objLote)

    'Lê o Lote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 9293 Then Error 41622

    If lErro = 9293 Then Error 41623

    'Preenche objLançamento_Detalhe
    objLancamento_Detalhe.sOrigem = objLote.sOrigem
    objLancamento_Detalhe.iExercicio = objLote.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objLote.iPeriodo
    objLancamento_Detalhe.iLote = objLote.iLote

    'Adiciona filtro: iFilialEmpresa, sOrigem, iExercicio, iPeriodo
    colSelecao.Add objLancamento_Detalhe.sOrigem
    colSelecao.Add objLancamento_Detalhe.iExercicio
    colSelecao.Add objLancamento_Detalhe.iPeriodoLan
    colSelecao.Add objLancamento_Detalhe.iLote

    'Chama Tela LancamentosLista
    Call Chama_Tela("LancamentoLista_Lote", colSelecao, objLancamento_Detalhe, objEventoLancamento)

    Exit Sub

Erro_LancamentoLote_Click:

    Select Case Err

        Case 41618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)

        Case 41619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)

        Case 41620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 41621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", Err)

        Case 41622

        Case 41623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162487)

    End Select

    Exit Sub

End Sub

Private Sub Lote_GotFocus()
    Call MaskEdBox_TrataGotFocus(Lote)
End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objLote As ClassLote
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1

    Lote.Text = CStr(objLote.iLote)

    sDescricao = gobjColOrigem.Descricao(objLote.sOrigem)

    For iIndice = 0 To Origem.ListCount - 1
        If Origem.List(iIndice) = sDescricao Then
            Origem.ListIndex = iIndice
            Exit For
        End If
    Next

    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = objLote.iExercicio Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = objLote.iPeriodo Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next

    'Mostra os dados na tela
    Call Traz_Dados_Tela(objLote)

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162488)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objLote As ClassLote) As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoExtornar_Click()

Dim lErro As Long
Dim objLote As New ClassLote
Dim objBrowseConfigura As New AdmBrowseConfigura

On Error GoTo Erro_BotaoExtornar_Click

    'Se Exercicio estiver vazio, Erro
    If Len(Trim(Exercicio.Text)) = 0 Then Error 41636

    'Se Periodo estiver vazio, Erro
    If Len(Trim(Periodo.Text)) = 0 Then Error 41637

    'Se Lote estiver vazio, Erro
    If Len(Trim(Lote.Text)) = 0 Then Error 41638

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 41639

    Call Move_Tela_Memoria(objLote)

    'Lê o Lote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 9293 Then Error 41626

    If lErro = 9293 Then Error 41627

    'chama a tela LoteEstorno1 Modal
    lErro = Chama_Tela_Modal("LoteEstorno1", objLote, objBrowseConfigura)
    If lErro <> SUCESSO Then Error 41628

    'verifica se não foi cancelado
    If objBrowseConfigura.iTelaOK = CANCELA Then Error 41629

    Call LoteEstorno_LimpaTela

    Exit Sub

Erro_BotaoExtornar_Click:

    Select Case Err

        Case 41626, 41628

        Case 41627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case 41629
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_ESTORNO_LOTE_CANCELADO", objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 41636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)

        Case 41637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)

        Case 41638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 41639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162489)

    End Select

    Exit Sub

End Sub

Private Sub LoteEstorno_LimpaTela()

Dim lErro As Long

On Error GoTo Erro_LoteEstorno_LimpaTela

    Call Limpa_Labels

    Lote.Text = ""
    
    lErro = Inicializa_Cabecalho()
    If lErro <> SUCESSO Then Error 41625

    Exit Sub

Erro_LoteEstorno_LimpaTela:

    Select Case Err

        Case 41625

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162490)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria(objLote As ClassLote)
    
    'carrega em memória os dados da tela
    If Exercicio.ListIndex = -1 Then
        objLote.iExercicio = 0
    Else
        objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    End If
    
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    
    If Len(TotLancInf.Caption) = 0 Then
        objLote.dTotInf = 0
    Else
        objLote.dTotInf = StrParaDbl(TotLancInf.Caption)
    End If
        
    If Len(NumLancInf.Caption) = 0 Then
        objLote.iNumDocInf = 0
    Else
        objLote.iNumDocInf = CInt(NumLancInf.Caption)
    End If
    
    objLote.sIdOriginal = IdLoteExterno.Caption
    
    If Len(Trim(Lote.Text)) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(Lote.Text)
    End If
    
End Sub

Private Function Traz_Lote_Tela(objLote As ClassLote) As Long

Dim sDescricao As String
Dim sExercicio As String
Dim lErro As Long
Dim iIndice As Integer
Dim iLoteAtualizado As Integer

On Error GoTo Erro_Traz_Lote_Tela

    Origem.Text = gobjColOrigem.Descricao(objLote.sOrigem)

    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado",objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 59573

    'le o lote contido em objLote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 5435 Then Error 59574
    
    'move os dados para a tela
    sDescricao = gobjColOrigem.Descricao(objLote.sOrigem)
    
    'mostra o Exercicio
    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = objLote.iExercicio Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next
        
    'mostra o periodo
    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = objLote.iPeriodo Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Lote.Text = CStr(objLote.iLote)
    
    'se o lote está cadastrado, coloca o restante das informações na tela
    If lErro = SUCESSO Then
            
        TotLancInf = Format(objLote.dTotInf, "Fixed")
        NumLancInf = CStr(objLote.iNumDocInf)
        TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
        TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
        NumLancAtual.Caption = Format(objLote.iNumDocAtual, "##,##0")
        IdLoteExterno = objLote.sIdOriginal
        
    End If
    
    Traz_Lote_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Lote_Tela:

    Traz_Lote_Tela = Err

    Select Case Err
    
        Case 59573, 59574
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162491)
        
    End Select
    
    Exit Function
        
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objLote As New ClassLote
Dim colLancamento_Detalhe As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Lote"
            
    'Le os dados da Tela de Lotes
    Call Move_Tela_Memoria(objLote)
  
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objLote.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Origem", objLote.sOrigem, STRING_ORIGEM, "Origem"
    colCampoValor.Add "Lote", objLote.iLote, 0, "Lote"
    colCampoValor.Add "Exercicio", objLote.iExercicio, 0, "Exercicio"
    colCampoValor.Add "Periodo", objLote.iPeriodo, 0, "Periodo"
    
    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
'    If sSiglaModulo = MODULO_CONTABILIDADE Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTABILIDADE
'    If sSiglaModulo = MODULO_CONTASAPAGAR Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTASAPAGAR
'    If sSiglaModulo = MODULO_CONTASARECEBER Then colSelecao.Add "Origem", OP_IGUAL, MODULO_CONTASARECEBER
'    If sSiglaModulo = MODULO_TESOURARIA Then colSelecao.Add "Origem", OP_IGUAL, MODULO_TESOURARIA
'    If sSiglaModulo = MODULO_FATURAMENTO Then colSelecao.Add "Origem", OP_IGUAL, MODULO_FATURAMENTO
'    If sSiglaModulo = MODULO_ESTOQUE Then colSelecao.Add "Origem", OP_IGUAL, MODULO_ESTOQUE
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162492)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_Tela_Preenche

    objLote.iLote = colCampoValor.Item("Lote").vValor

    If objLote.iLote <> 0 Then
    
        objLote.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objLote.sOrigem = colCampoValor.Item("Origem").vValor
        objLote.iPeriodo = colCampoValor.Item("Periodo").vValor
        objLote.iExercicio = colCampoValor.Item("Exercicio").vValor
       
        lErro = Traz_Lote_Tela(objLote)
        If lErro <> SUCESSO Then Error 59575
                
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 59575
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162493)

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

    Parent.HelpContextID = IDH_EXTORNO_LOTE_CONTABILIZADO
    Set Form_Load_Ocx = Me
    Caption = "Estorno de Lote Contabilizado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteEstorno"
    
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
        
        If Me.ActiveControl Is Lote Then
            Call LabelLote_Click
        End If
    
    End If

End Sub


Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub TotCredAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotCredAtual, Source, X, Y)
End Sub

Private Sub TotCredAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotCredAtual, Button, Shift, X, Y)
End Sub

Private Sub TotDebAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotDebAtual, Source, X, Y)
End Sub

Private Sub TotDebAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotDebAtual, Button, Shift, X, Y)
End Sub

Private Sub NumLancAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumLancAtual, Source, X, Y)
End Sub

Private Sub NumLancAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumLancAtual, Button, Shift, X, Y)
End Sub

Private Sub IdLoteExterno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IdLoteExterno, Source, X, Y)
End Sub

Private Sub IdLoteExterno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IdLoteExterno, Button, Shift, X, Y)
End Sub

Private Sub TotLancInf_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotLancInf, Source, X, Y)
End Sub

Private Sub TotLancInf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotLancInf, Button, Shift, X, Y)
End Sub

Private Sub NumLancInf_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumLancInf, Source, X, Y)
End Sub

Private Sub NumLancInf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumLancInf, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub LabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLote, Source, X, Y)
End Sub

Private Sub LabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLote, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

