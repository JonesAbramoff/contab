VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LancamentoConsultaOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11835
   Begin VB.CheckBox Gerencial 
      Enabled         =   0   'False
      Height          =   210
      Left            =   9690
      TabIndex        =   31
      Tag             =   "1"
      Top             =   2460
      Width           =   870
   End
   Begin VB.CommandButton BotaoDocOriginal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7200
      Picture         =   "LancamentoConsultaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   225
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   8940
      ScaleHeight     =   765
      ScaleWidth      =   2760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2820
      Begin VB.CommandButton BotaoExcluir 
         Height          =   600
         Left            =   1815
         Picture         =   "LancamentoConsultaOcx.ctx":2F16
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   600
         Left            =   1335
         Picture         =   "LancamentoConsultaOcx.ctx":30A0
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   600
         Left            =   60
         Picture         =   "LancamentoConsultaOcx.ctx":31FA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   600
         Left            =   2310
         Picture         =   "LancamentoConsultaOcx.ctx":4FBC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5280
      MaxLength       =   150
      TabIndex        =   7
      Top             =   2100
      Width           =   4260
   End
   Begin VB.ComboBox Origem 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   1935
   End
   Begin MSMask.MaskEdBox Debito 
      Height          =   225
      Left            =   4140
      TabIndex        =   6
      Top             =   2100
      Width           =   1155
      _ExtentX        =   2037
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
      Left            =   2970
      TabIndex        =   5
      Top             =   2100
      Width           =   1155
      _ExtentX        =   2037
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
      Left            =   2265
      TabIndex        =   4
      Top             =   2100
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
      Left            =   960
      TabIndex        =   3
      Top             =   2085
      Width           =   1756
      _ExtentX        =   3096
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
      Height          =   315
      Left            =   2190
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   585
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Documento 
      Height          =   285
      Left            =   5865
      TabIndex        =   1
      Top             =   210
      Width           =   1155
      _ExtentX        =   2037
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
      Left            =   1035
      TabIndex        =   2
      Top             =   585
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
      Left            =   450
      TabIndex        =   8
      Top             =   1785
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1050
      TabIndex        =   27
      Top             =   990
      Width           =   1980
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   26
      Top             =   1005
      Width           =   615
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
      Left            =   495
      TabIndex        =   13
      Top             =   1560
      Width           =   1140
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
      Left            =   2850
      TabIndex        =   14
      Top             =   630
      Width           =   870
   End
   Begin VB.Label Exercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3780
      TabIndex        =   15
      Top             =   585
      Width           =   1185
   End
   Begin VB.Label Periodo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5850
      TabIndex        =   16
      Top             =   585
      Width           =   1185
   End
   Begin VB.Label TotalCredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2610
      TabIndex        =   17
      Top             =   3750
      Width           =   1155
   End
   Begin VB.Label TotalDebito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3870
      TabIndex        =   18
      Top             =   3750
      Width           =   1155
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
      Left            =   1800
      TabIndex        =   19
      Top             =   3765
      Width           =   705
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   495
      TabIndex        =   20
      Top             =   630
      Width           =   480
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
      Left            =   4815
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   21
      Top             =   225
      Width           =   1020
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
      Left            =   3300
      TabIndex        =   22
      Top             =   225
      Width           =   450
   End
   Begin VB.Label Label1 
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
      Left            =   315
      TabIndex        =   23
      Top             =   240
      Width           =   660
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
      Left            =   5085
      TabIndex        =   24
      Top             =   630
      Width           =   735
   End
   Begin VB.Label Lote 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3780
      TabIndex        =   25
      Top             =   195
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -570
      X2              =   11805
      Y1              =   1350
      Y2              =   1350
   End
End
Attribute VB_Name = "LancamentoConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Revisado por: Mário
'Data: 8/10/98

'Pendencias: Colocar o browser correto, transferir a função que acessa BD
'            Esta tela deve ser chamada somente pelo browser. E o browser deve ser chamado através do menu de consultas

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
Dim iLote_Lost_Focus As Integer
Dim iData_Lost_Focus As Integer
Dim iAlterado As Integer


Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim vbMsgRes As VbMsgBoxResult
Dim lDoc As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
     
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then gError 183106
    
    'Documento
    If Len(Documento.ClipText) = 0 Then gError 183107
     
    'Origem só pode ser CTB
'    If gobjColOrigem.Origem(Origem.Caption) <> "CTB" Then gError 59511
 
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DOCUMENTO")
    
    If vbMsgRes = vbYes Then
    
        'Obtém Período e Exercício correspondentes à data
        dtData = CDate(Data.Text)
    
        'Lê o Período
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then gError 183108
    
        objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
        objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Text)
        objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
        objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
        objLancamento_Cabecalho.lDoc = StrParaLong(Documento.ClipText)
        objLancamento_Cabecalho.dtData = StrParaDate(Data.Text)
     
        lErro = CF("Lancamento_Exclui_4", objLancamento_Cabecalho)
        If lErro <> SUCESSO Then gError 183109
    
        Call Limpa_Tela_Lancamentos
    
        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 183106
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
        
        Case 183107
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case 183108, 183109
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183110)
        
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoConsultar_Click()

Dim lErro As Long
Dim dtData As Date
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim objPeriodo As New ClassPeriodo
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLanc As ClassLancamento_Detalhe
Dim colLancamentos As New Collection

On Error GoTo Erro_BotaoConsultar_Click

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then Error 28720

    'Se Documento estiver vazio, Erro
    If Len(Documento.ClipText) = 0 Then Error 28721

    'Se Data estiver vazio ==> Erro
    If Len(Data.ClipText) = 0 Then Error 28722
    
    'Obtém Período e Exercício correspondentes à data
    dtData = CDate(Data.Text)

    'Lê o Período
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 28693

    'Preenche objLote
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.lDoc = CLng(Documento.Text)
    
    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 28714 Then Error 28749

    'Documento não cadastrado
    If lErro = 28714 Then Error 28750
    
    Exit Sub

Erro_BotaoConsultar_Click:

    Select Case Err

        Case 28720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", Err)

        Case 28721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
            Documento.SetFocus

        Case 28722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)
            Data.SetFocus

        Case 28693, 28749

        Case 28750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162188)

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

    If iLote_Lost_Focus = 0 Then

        If Len(Data.ClipText) > 0 Then

            lErro = Data_Critica(Data.Text)
            If lErro <> SUCESSO Then Error 28689

            'Obtém Período e Exercício correspondentes à data
            dtData = CDate(Data.Text)

            'Lê o Período
            lErro = CF("Periodo_Le", dtData, objPeriodo)
            If lErro <> SUCESSO Then Error 28690

            'Lê o Exercício
            lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
            If lErro <> SUCESSO And lErro <> 10083 Then Error 28691

            'Exercício não cadastrado
            If lErro = 10083 Then Error 28692

            'Preenche campo de período
            Periodo.Caption = objPeriodo.sNomeExterno

            'Preenche campo de exercício
            Exercicio.Caption = objExercicio.sNomeExterno

        Else
        
            'Limpa os campos período e exercício
            Periodo.Caption = ""
            Exercicio.Caption = ""

        End If

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 28689
            iData_Lost_Focus = 1

        Case 28690, 28691, 28695

        Case 28692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162189)

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
        If lErro <> SUCESSO Then Error 28709

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

        Case 28709

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162190)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_Form_Load

    iData_Lost_Focus = 0
    iLote_Lost_Focus = 0

    Set objEventoLancamento = New AdmEvento

    Set objGrid1 = New AdmGrid

    lErro = Inicializa_Grid_Lancamento(objGrid1)
    If lErro <> SUCESSO Then Error 28711

    GridLancamentos.Row = 1
    GridLancamentos.Col = 1

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

    lErro = Traz_Cabecalho_Tela()
    If lErro <> SUCESSO Then Error 55719

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28711, 55719

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162191)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
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
    If lErro <> SUCESSO Then Error 28712

    objGridInt.objGrid = GridLancamentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = MAX_LANCAMENTOS_POR_DOC_CTB + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 18 '7

    GridLancamentos.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

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

        Case 28712

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162192)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objLancamento_Detalhe As ClassLancamento_Detalhe) As Long
     
Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Trata_Parametros
    
    If Not (objLancamento_Detalhe Is Nothing) Then
    
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
        objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
        objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
        objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
        objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc
         
        'Traz os dados para tela -----> o que veio no parametro
        lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
        If lErro <> SUCESSO Then Error 40842
    
        iAlterado = 0
        
    End If
        
    Trata_Parametros = SUCESSO

    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 40842

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162193)

    End Select

    iAlterado = 0
    
    Exit Function
    
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridLancamentos.Col

            Case iGrid_Historico_Col
            
                lErro = Saida_Celula_Historico(objGridInt)
                If lErro <> SUCESSO Then gError 92116
               
               

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92117
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 92116
        
        Case 92117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162194)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Historico

    Set objHistPadrao = New ClassHistPadrao
    
    Set objGridInt.objControle = Historico
                
    If Left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
    
        sValor = Trim(Mid(Historico.Text, 2))
        
        lErro = Valor_Inteiro_Critica(sValor)
        If lErro <> SUCESSO Then gError 92118
        
        objHistPadrao.iHistPadrao = CInt(sValor)
                
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then gError 92119
        
        If lErro = 5446 Then gError 92120

        Historico.Text = objHistPadrao.sDescHistPadrao
        
    End If
    
    If Len(Historico.Text) > 0 Then
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92121

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = gErr
    
    Select Case gErr
    
        Case 92118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", gErr, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 92119, 92121
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92120
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162195)
        
    End Select

    Exit Function

End Function

Private Sub GridLancamentos_Click()

Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Private Sub GridLancamentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridLancamentos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Private Sub GridLancamentos_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Private Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGrid1)
    If lErro <> SUCESSO Then gError 92129
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case gErr
    
        Case 92129
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162196)
    
    End Select

    Exit Sub

End Sub

Private Sub GridLancamentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridLancamentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridLancamentos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Private Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Historico_GotFocus()
    
Dim iPos As Integer
    
    Call Grid_Campo_Recebe_Foco(objGrid1)
    
    If Len(Historico.Text) > 0 Then
        iPos = InStr(Historico.Text, CARACTER_HISTORICO_PARAM)
        If iPos > 0 Then
            Historico.SelStart = iPos - 1
            Historico.SelLength = 1
        End If
    End If
    
End Sub

Private Sub Historico_KeyPress(KeyAscii As Integer)

Dim iInicio As Integer
Dim iPos As Integer
Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Historico_KeyPress

    'se digitou ENTER
    If KeyAscii = vbKeyReturn Then
        
        If Len(Historico.Text) > 0 Then
        
            If Left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(Historico.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44073
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44074
                
                If lErro = 5446 Then Error 44075
        
                Historico.Text = objHistPadrao.sDescHistPadrao
                Historico.SelStart = 0
                
            End If
    
            If Historico.SelText = CARACTER_HISTORICO_PARAM Then
                iInicio = Historico.SelStart + 2
            Else
                iInicio = Historico.SelStart
            End If
        
            If iInicio = 0 Then iInicio = 1
        
            iPos = InStr(iInicio, Historico.Text, CARACTER_HISTORICO_PARAM)
            If iPos > 0 Then
                Historico.SelStart = iPos - 1
                Historico.SelLength = 1
                Exit Sub
            End If
        End If
    End If

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
    Exit Sub
    
Erro_Historico_KeyPress:

    Select Case Err
    
        Case 44073
            objGrid1.iExecutaSaidaCelula = 0
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            objGrid1.iExecutaSaidaCelula = 1
        
        Case 44074

        Case 44075
            objGrid1.iExecutaSaidaCelula = 0
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Chama_Tela("HistoricoPadrao", objHistPadrao)
            
            Else
                Historico.SetFocus
            End If
            
            objGrid1.iExecutaSaidaCelula = 1
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162197)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Historico
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoLancamento = Nothing

    Set objGrid1 = Nothing
    
End Sub

Private Sub objEventoLancamento_evSelecao(obj1 As Object)
'Traz o lançamento selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice As Integer
Dim sDescricao As String
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1

    objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
    objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
    objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc
       
    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 28710

    Me.Show

    Exit Sub

Erro_objEventoLancamento_evSelecao:

    Select Case Err

        Case 28710

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162198)

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
    
     'Mostra a origem passada como parametro
    For iIndice = 0 To Origem.ListCount - 1
        If Origem.List(iIndice) = gobjColOrigem.Descricao(objDoc.sOrigem) Then
            Origem.ListIndex = iIndice
            Exit For
        End If
    Next

    Documento.Text = CStr(objDoc.lDoc)

    iFilialEmpresaSalva = objDoc.iFilialEmpresa

    Do While objDoc.iFilialEmpresa > 0 And objDoc.iFilialEmpresa < 100

        'Lê os lançamentos contidos no documento em questão
        lErro = CF("Lancamentos_Le_Doc", objDoc, colLancamentos)
        If lErro <> SUCESSO And lErro <> 28700 Then Error 28713

        If lErro = SUCESSO Then iAchou = 1

        If giContabGerencial = 0 Then Exit Do

        objDoc.iFilialEmpresa = objDoc.iFilialEmpresa - giFilialAuxiliar

    Loop

    objDoc.iFilialEmpresa = iFilialEmpresaSalva
    
    
    'se não encontrou o documento
    If iAchou = 0 Then gError 28714

    For iIndice = colLancamentos.Count To 1 Step -1
        
        For iIndice1 = iIndice - 1 To 1 Step -1
        
            If colLancamentos(iIndice).iSeq = colLancamentos(iIndice1).iSeq Then
                colLancamentos.Remove (iIndice)
                Exit For
            End If
            
        Next
    
    Next


    Set objLanc = colLancamentos.Item(1)

    Lote.Caption = CStr(objLanc.iLote)

    'Inicializa Data
    Data.Text = Format(objLanc.dtData, "dd/mm/yy")

    'Coloca o período relativo a data na tela
    lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 28715

    Periodo.Caption = objPeriodo.sNomeExterno

    'Coloca o exercício na tela
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 28716

    'Se o exercício não está cadastrado
    If lErro = 10083 Then Error 28717

    Exercicio.Caption = objExercicio.sNomeExterno

    'Move os dados para a tela
    For Each objLanc In colLancamentos

        Select Case objLanc.iStatus
            
            Case VOUCHER_ESTORNADO
                Status.Caption = STRINT_VOUCHER_ESTORNADO
                
            Case VOUCHER_NORMAL
                Status.Caption = STRINT_VOUCHER_NORMAL
                
            Case VOUCHER_ESTORNADOR
                Status.Caption = STRINT_VOUCHER_ESTORNADOR
                
        End Select

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objLanc.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 28718

        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True

        'Coloca a conta na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = Conta.Text

        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

            'mascara o centro de custo
            sCclMascarado = String(STRING_CCL, 0)

            If objLanc.sCcl <> "" Then

                lErro = Mascara_MascararCcl(objLanc.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 28719

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

        If giContabGerencial = 1 Then
            If objLanc.iSeq > objGrid1.iLinhasExistentes Then objGrid1.iLinhasExistentes = objLanc.iSeq
        Else
            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1
        End If
            
    Next
    
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
        
    Call Grid_Refresh_Checkbox(objGrid1)
        
    iAlterado = 0
    
    Traz_Doc_Tela = SUCESSO

    Exit Function

Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err

        Case 28713, 28714, 28715, 28716

        Case 28717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 28718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objLanc.sConta)

        Case 28719
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objLanc.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162199)

    End Select

    iAlterado = 0
    
    Exit Function

End Function

Private Sub BotaoGravar_Click()

    Call Gravar_Registro
    
    iAlterado = 0
    
End Sub

Public Function Gravar_Registro() As Long
    
Dim lErro As Long
Dim lDoc As Long
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim iIndice1 As Integer
Dim dSoma As Double
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then gError 92125
        
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then gError 92122
    
    'Verifica a existencia de pelo menos um lançamento
    If objGrid1.iLinhasExistentes = 0 Then gError 92123
    
    'Documento
    If Len(Documento.ClipText) = 0 Then gError 92124
        
    lErro = Move_Tela_Memoria(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then gError 92127
        
    'Preenche Objeto Lançamento_Detalhe
    lErro = Grid_Lancamento_Detalhe(colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 92130
        
    lErro = CF("Lancamento_Altera_Historico", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 92131
        
    Call Limpa_Tela_Lancamentos

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 92122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus
            
        Case 92123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_LANCAMENTOS_GRAVAR", gErr)
        
        Case 92124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus
            
        Case 92125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", gErr)
            Origem.SetFocus
            
        Case 92127, 92130, 92131
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162200)
            
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Obtém Período e Exercício correspondentes à data
    dtData = CDate(Data.Text)

    'Lê o Período
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 92126

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

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 92126

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162201)

    End Select

    Exit Function

End Function

Function Grid_Lancamento_Detalhe(colLancamento_Detalhe As Collection) As Long

Dim iIndice1 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim lErro As Long

On Error GoTo Erro_Grid_Lancamento_Detalhe

    For iIndice1 = 1 To objGrid1.iLinhasExistentes
        
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.iSeq = iIndice1
    
        'Armazena Histórico e Ccl
        objLancamento_Detalhe.sHistorico = GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col)
        If giContabGerencial = 1 Then objLancamento_Detalhe.iGerencial = GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col)
            
        'verifica se o historico tem parametros que deveriam ter sido substituidos
        If InStr(objLancamento_Detalhe.sHistorico, CARACTER_HISTORICO_PARAM) <> 0 Then gError 92128
            
        'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
        colLancamento_Detalhe.Add objLancamento_Detalhe
        
    Next
    
    Grid_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Grid_Lancamento_Detalhe:

    Grid_Lancamento_Detalhe = gErr

    Select Case gErr
    
        Case 92128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICO_PARAM", gErr)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Historico_Col
            GridLancamentos.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162202)
            
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
    If lErro <> SUCESSO Then Error 28726

    Conta.Mask = sMascaraConta

    'Se usa centro de custo/lucro ==> inicializa máscara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        sMascaraCcl = String(STRING_CCL, 0)

        'Lê a máscara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 28727

        Ccl.Mask = sMascaraCcl

    End If

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 28726, 28727

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162203)

    End Select

    Exit Function


End Function

Private Sub Panel3D1_Click()

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 28742
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_DownClick:
    
    Select Case Err
    
        Case 28742
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162204)
        
    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 28743
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_UpClick:
    
    Select Case Err
    
        Case 28743
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162205)
        
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Cabecalho_Tela() As Long

Dim sDescricao As String
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim iIndice As Integer
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Traz_Cabecalho_Tela

    'Inicializa Data
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 55716
    
    Periodo.Caption = objPeriodo.sNomeExterno
    
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 55717
    
    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then Error 55718
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Traz_Cabecalho_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Cabecalho_Tela:

    Traz_Cabecalho_Tela = Err

    Select Case Err
    
        Case 55716, 55717
            
        Case 55718
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162206)
    
    End Select
    
    Exit Function

End Function

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
    If lErro <> SUCESSO And lErro <> 28732 Then Error 59561
    
    If lErro = 28732 Then Error 59562
    
    If objExercicio.iExercicio <> 0 Then objPeriodo.iExercicio = objExercicio.iExercicio
    If Len(Trim(Periodo.Caption)) > 0 Then objPeriodo.sNomeExterno = Periodo.Caption
    
    'Lê o Período
    lErro = CF("Periodo_Le_Codigo", objPeriodo)
    If lErro <> SUCESSO And lErro <> 28736 Then Error 59563
    
    If lErro = 28736 Then Error 59564
    
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
    
        Case 59561, 59563
        
        Case 59562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_INEXISTENTE", Err, objExercicio.sNomeExterno)
            
        Case 59564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_EXERCICIO_INEXISTENTE", Err, objPeriodo.iExercicio, objPeriodo.sNomeExterno)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162207)

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
        If lErro <> SUCESSO And lErro <> 5843 Then Error 59565

        If lErro = 5843 Then Error 59566

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 59565

        Case 59566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162208)

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
    
    Parent.HelpContextID = IDH_LANCAMENTO_CONSULTA
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Lançamentos Contabilizados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LancamentoConsulta"
    
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

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Periodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Periodo, Source, X, Y)
End Sub

Private Sub Periodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Periodo, Button, Shift, X, Y)
End Sub

Private Sub TotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCredito, Source, X, Y)
End Sub

Private Sub TotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCredito, Button, Shift, X, Y)
End Sub

Private Sub TotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalDebito, Source, X, Y)
End Sub

Private Sub TotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalDebito, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub DocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DocumentoLabel, Source, X, Y)
End Sub

Private Sub DocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DocumentoLabel, Button, Shift, X, Y)
End Sub

Private Sub LoteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LoteLabel, Source, X, Y)
End Sub

Private Sub LoteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LoteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Lote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Lote, Source, X, Y)
End Sub

Private Sub Lote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Lote, Button, Shift, X, Y)
End Sub

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim dtData As Date
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim objPeriodo As New ClassPeriodo, sNomeTela As String
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLanc As ClassLancamento_Detalhe, objNFiscal As New ClassNFiscal
Dim colLancamentos As New Collection, iOrigemLcto As Integer
Dim objTituloPagar As New ClassTituloPagar, objTituloReceber As New ClassTituloReceber
Dim alComando(0 To 2) As Long, lNumIntDoc As Long, iIndice As Integer

On Error GoTo Erro_BotaoDocOriginal_Click

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 130796
    Next
    
    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Text)) = 0 Then gError 28720

    'Se Documento estiver vazio, Erro
    If Len(Documento.ClipText) = 0 Then gError 28721

    'Se Data estiver vazio ==> Erro
    If Len(Data.ClipText) = 0 Then gError 28722
    
    'Obtém Período e Exercício correspondentes à data
    dtData = CDate(Data.Text)

    'Lê o Período
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 28693

    'Preenche objLote
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Text)
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.lDoc = CLng(Documento.Text)
    
    'Lê os lançamentos contidos no documento em questão
    lErro = CF("Lancamentos_Le_Doc", objLancamento_Cabecalho, colLancamentos)
    If lErro <> SUCESSO And lErro <> 28700 Then gError 28713

    'se o documento não estiver cadastrado
    If lErro = 28700 Then gError 28714

    'Se encontrou o documento
    If lErro = SUCESSO Then

        Set objLanc = colLancamentos.Item(1)
        
        If objLanc.iTransacao <> 0 Then
        
            lErro = Comando_Executar(alComando(2), "SELECT OrigemLcto FROM TransacaoCTB WHERE Codigo = ? ORDER BY subtipo", iOrigemLcto, objLanc.iTransacao)
            If lErro <> AD_SQL_SUCESSO Then gError 130793
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 130793
            
            Select Case iOrigemLcto
            
                Case 2
                    'se a baixa tiver sido de mais de um titulo
                    'deveria abrir browse de titulos que tiveram baixa associada à baixapag
                    lErro = Comando_Executar(alComando(0), "SELECT titulospagtodos.numintdoc FROM baixaspag, baixasparcpag, parcelaspagtodas, titulospagtodos WHERE baixaspag.NumIntBaixa = baixasparcpag.NumIntBaixa AND baixasparcpag.numintparcela = parcelaspagtodas.numintdoc AND titulospagtodos.numintdoc = parcelaspagtodas.numinttitulo AND baixaspag.numintbaixa = ?", lNumIntDoc, objLanc.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 130797
                    
                    lErro = Comando_BuscarProximo(alComando(0))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130798
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objTituloPagar.lNumIntDoc = lNumIntDoc
                        
                        lErro = CF("TituloPagar_Le_Todos", objTituloPagar)
                        If lErro <> SUCESSO Then gError 130791
                    
                        lErro = Chama_Tela("TituloPagar_Consulta", objTituloPagar)
                        If lErro <> SUCESSO Then gError 130794
                        
                    End If
                    
                Case 5
                    'se a baixa tiver sido de mais de um titulo
                    'deveria abrir browse de titulos que tiveram baixa associada à baixarec
                    lErro = Comando_Executar(alComando(0), "SELECT titulosrectodos.numintdoc FROM baixasrec, baixasparcrec, parcelasrectodas, titulosrectodos WHERE baixasrec.NumIntBaixa = baixasparcrec.NumIntBaixa AND baixasparcrec.numintparcela = parcelasrectodas.numintdoc AND titulosrectodos.numintdoc = parcelasrectodas.numinttitulo AND baixasrec.numintbaixa = ?", lNumIntDoc, objLanc.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 130797
                    
                    lErro = Comando_BuscarProximo(alComando(0))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130798
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objTituloReceber.lNumIntDoc = lNumIntDoc
                        lErro = Chama_Tela("TituloReceber_Consulta", objTituloReceber)
                        If lErro <> SUCESSO Then gError 130794
                    End If
                    
                Case 15
                    objTituloPagar.lNumIntDoc = objLanc.lNumIntDoc
                    
                    lErro = CF("TituloPagar_Le_Todos", objTituloPagar)
                    If lErro <> SUCESSO Then gError 130791
                    
                    lErro = Chama_Tela("TituloPagar_Consulta", objTituloPagar)
                    If lErro <> SUCESSO Then gError 130791
                    
                Case 16
                    objTituloReceber.lNumIntDoc = objLanc.lNumIntDoc
                    lErro = Chama_Tela("TituloReceber_Consulta", objTituloReceber)
                    If lErro <> SUCESSO Then gError 130792
                    
                Case 10
                    sNomeTela = String(STRING_NOME_TELA, 0)
                    lErro = Comando_Executar(alComando(1), "SELECT TiposDocInfo.NomeTelaNFiscal FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumIntDoc = ?", sNomeTela, objLanc.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 130800
                    
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130801
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objNFiscal.lNumIntDoc = objLanc.lNumIntDoc
                        lErro = Chama_Tela(sNomeTela, objNFiscal)
                        If lErro <> SUCESSO Then gError 130802
                    End If
                                        
            End Select
        
        End If
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 130796
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 28720
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", gErr)

        Case 28721
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus

        Case 28722
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 28693, 28749, 130791, 130972, 130973, 130794, 130975, 130797, 130798, 130800, 130801, 130802

        Case 28714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", gErr, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162188)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Sub

End Sub


