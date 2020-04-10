VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LoteConsultaOcx 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   8025
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   5940
      ScaleHeight     =   765
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   180
      Width           =   1920
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   1380
         Picture         =   "LoteConsultaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   600
         Left            =   75
         Picture         =   "LoteConsultaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.CommandButton LancamentoLote 
      Caption         =   "Lançamentos do Lote"
      Height          =   765
      Left            =   5370
      Picture         =   "LoteConsultaOcx.ctx":1F40
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3750
      Width           =   2100
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   1695
   End
   Begin VB.ComboBox Periodo 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   1980
   End
   Begin VB.ComboBox Origem 
      Height          =   315
      Left            =   3585
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   2010
   End
   Begin VB.Frame SSFrame1 
      Caption         =   "Valores Atuais"
      Height          =   1650
      Left            =   135
      TabIndex        =   8
      Top             =   1920
      Width           =   7515
      Begin VB.Label NumLancAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Left            =   2475
         TabIndex        =   9
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label TotDebAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   10
         Top             =   802
         Width           =   1575
      End
      Begin VB.Label TotCredAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Height          =   300
         Left            =   2475
         TabIndex        =   11
         Top             =   375
         Width           =   1575
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
         Top             =   420
         Width           =   2115
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
         Left            =   345
         TabIndex        =   13
         Top             =   825
         Width           =   2070
      End
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
         Left            =   315
         TabIndex        =   14
         Top             =   1260
         Width           =   2100
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
      Left            =   525
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   735
      Width           =   450
   End
   Begin VB.Label Label2 
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
      Left            =   2850
      TabIndex        =   16
      Top             =   735
      Width           =   660
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
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   255
      Width           =   885
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
      Left            =   2790
      TabIndex        =   18
      Top             =   270
      Width           =   750
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
      Left            =   180
      TabIndex        =   19
      Top             =   1335
      Width           =   2475
   End
   Begin VB.Label Label6 
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
      Left            =   4455
      TabIndex        =   20
      Top             =   1335
      Width           =   2100
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
      Left            =   180
      TabIndex        =   21
      Top             =   4035
      Width           =   2550
   End
   Begin VB.Label NumLancInf 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   22
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label TotLancInf 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   23
      Top             =   1305
      Width           =   1530
   End
   Begin VB.Label IdLoteExterno 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2805
      TabIndex        =   24
      Top             =   4005
      Width           =   1530
   End
End
Attribute VB_Name = "LoteConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Revisado Por: Mario
'Data: 8/10/98

'Pendencias: transferir mensagem. Acertar os browsers (só lotes e lançamentos contabilizados)
'            Esta tela deve ser chamada somente pelo browser. E o browser deve ser chamado através do menu de consultas

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

Private Sub BotaoConsultar_Click()

Dim lErro As Long
Dim objLote As New ClassLote

On Error GoTo Erro_BotaoConsultar_Click

    'Se Exercicio estiver vazio, Erro
    If Len(Trim(Exercicio.Text)) = 0 Then Error 28675

    'Se Periodo estiver vazio, Erro
    If Len(Trim(Periodo.Text)) = 0 Then Error 28676

    'Se Lote estiver vazio, Erro
    If Len(Trim(Lote.Text)) = 0 Then Error 28677

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 28678

    'Preenche objLote
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    objLote.iLote = CInt(Lote.Text)

    'Lê o Lote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 9293 Then Error 28679

    If lErro = 9293 Then Error 28680

    'Mostra os dados na tela
    If objLote.dTotInf > 0 Then
        TotLancInf.Caption = Format(objLote.dTotInf, "Standard")
    Else
        TotLancInf.Caption = ""
    End If
    If objLote.iNumLancInf > 0 Then
        NumLancInf.Caption = CStr(objLote.iNumDocInf)
    Else
        NumLancInf.Caption = ""
    End If
    If objLote.dTotCre > 0 Then
        TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
    Else
        TotCredAtual.Caption = "0,00"
    End If
    If objLote.dTotDeb > 0 Then
        TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
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

    Exit Sub

Erro_BotaoConsultar_Click:

    Select Case Err

        Case 28675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)

        Case 28676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)

        Case 28677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 28678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", Err)

        Case 28679

        Case 28680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162449)

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
    If lErro <> SUCESSO Then Error 28682

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

        Case 28682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162450)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim iLote As Integer
Dim sDescricao As String
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_Form_Load

    Set objEventoLote = New AdmEvento
    Set objEventoLancamento = New AdmEvento

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
    If lErro <> SUCESSO Then Error 28681

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

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28681

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162451)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoLote = Nothing
    Set objEventoLancamento = Nothing

End Sub

Private Sub LabelLote_Click()

Dim objLote As New ClassLote
Dim colSelecao As New Collection

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
    If Len(Trim(Exercicio.Text)) = 0 Then Error 28669

    'Se Periodo estiver vazio, Erro
    If Len(Trim(Periodo.Text)) = 0 Then Error 28670

    'Se Lote estiver vazio, Erro
    If Len(Trim(Lote.Text)) = 0 Then Error 28671

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 28672

    'Preenche objLote
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLote.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    objLote.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
    objLote.iLote = CInt(Lote.Text)

    'Lê o Lote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 9293 Then Error 28673

    If lErro = 9293 Then Error 28674

    'Preenche objLançamento_Detalhe
    objLancamento_Detalhe.sOrigem = objLote.sOrigem
    objLancamento_Detalhe.iExercicio = objLote.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objLote.iPeriodo
    objLancamento_Detalhe.iLote = objLote.iLote

    'Adiciona filtro: iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote
    colSelecao.Add objLancamento_Detalhe.sOrigem
    colSelecao.Add objLancamento_Detalhe.iExercicio
    colSelecao.Add objLancamento_Detalhe.iPeriodoLan
    colSelecao.Add objLancamento_Detalhe.iLote
    
    'Chama Tela LancamentosLista
    Call Chama_Tela("LancamentoLista_Lote", colSelecao, objLancamento_Detalhe, objEventoLancamento)

    Exit Sub

Erro_LancamentoLote_Click:

    Select Case Err

        Case 28669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)

        Case 28670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)

        Case 28671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)

        Case 28672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", Err)

        Case 28673

        Case 28674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162452)

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

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1
    
    For iIndice = 0 To Exercicio.ListCount - 1
        If Exercicio.ItemData(iIndice) = giExercicioAtual Then
            Exercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    For iIndice = 0 To Periodo.ListCount - 1
        If Periodo.ItemData(iIndice) = giPeriodoAtual Then
            Periodo.ListIndex = iIndice
            Exit For
        End If
    Next

    Lote.Text = CStr(objLote.iLote)

    'Mostra os dados na tela
    If objLote.dTotInf > 0 Then
        TotLancInf.Caption = Format(objLote.dTotInf, "Standard")
    Else
        TotLancInf.Caption = ""
    End If
    If objLote.iNumDocInf > 0 Then
        NumLancInf.Caption = CStr(objLote.iNumDocInf)
    Else
        NumLancInf.Caption = ""
    End If
    If objLote.dTotCre > 0 Then
        TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
    Else
        TotCredAtual.Caption = "0,00"
    End If
    If objLote.dTotDeb > 0 Then
        TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
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
    
    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162453)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objLote As ClassLote) As Long
    
Dim lErro As Long, iIndice As Integer

    If Not objLote Is Nothing Then
    
        Lote.Text = CStr(objLote.iLote)
         
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
         
         
        'mostra a origem
        For iIndice = 0 To Origem.ListCount - 1
            If gobjColOrigem.Origem(Origem.List(iIndice)) = objLote.sOrigem Then
                Origem.ListIndex = iIndice
                Exit For
            End If
        Next
                 
         
        'Mostra os dados na tela
        If objLote.dTotInf > 0 Then
            TotLancInf.Caption = Format(objLote.dTotInf, "Standard")
        Else
            TotLancInf.Caption = ""
        End If
        If objLote.iNumLancInf > 0 Then
            NumLancInf.Caption = CStr(objLote.iNumLancInf)
        Else
            NumLancInf.Caption = ""
        End If
        If objLote.dTotCre > 0 Then
            TotCredAtual.Caption = Format(objLote.dTotCre, "Standard")
        Else
            TotCredAtual.Caption = "0,00"
        End If
        If objLote.dTotDeb > 0 Then
            TotDebAtual.Caption = Format(objLote.dTotDeb, "Standard")
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
        
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:
    
    Trata_Parametros = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162454)

    End Select
    
    Exit Function
   
End Function

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
    If lErro <> SUCESSO Then Error 59576

    'le o lote contido em objLote
    lErro = CF("Lote_Le",objLote)
    If lErro <> SUCESSO And lErro <> 5435 Then Error 59577
    
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
    
        Case 59576, 59577
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162455)
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162456)

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
        If lErro <> SUCESSO Then Error 59578

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 59578

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162457)

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
    
    Parent.HelpContextID = IDH_LOTE_CONSULTA
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Lote Contabilizado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteConsulta"
    
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



Private Sub NumLancAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumLancAtual, Source, X, Y)
End Sub

Private Sub NumLancAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumLancAtual, Button, Shift, X, Y)
End Sub

Private Sub TotDebAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotDebAtual, Source, X, Y)
End Sub

Private Sub TotDebAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotDebAtual, Button, Shift, X, Y)
End Sub

Private Sub TotCredAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotCredAtual, Source, X, Y)
End Sub

Private Sub TotCredAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotCredAtual, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
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

Private Sub LabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLote, Source, X, Y)
End Sub

Private Sub LabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLote, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub NumLancInf_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumLancInf, Source, X, Y)
End Sub

Private Sub NumLancInf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumLancInf, Button, Shift, X, Y)
End Sub

Private Sub TotLancInf_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotLancInf, Source, X, Y)
End Sub

Private Sub TotLancInf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotLancInf, Button, Shift, X, Y)
End Sub

Private Sub IdLoteExterno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IdLoteExterno, Source, X, Y)
End Sub

Private Sub IdLoteExterno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IdLoteExterno, Button, Shift, X, Y)
End Sub

