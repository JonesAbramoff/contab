VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form BorderoPagGeraArqRem 
   Caption         =   "Geração de Arquivo de Pagtos - CNAB"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   525
      Left            =   2535
      Picture         =   "BorderoPagGeraArqRem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2145
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   525
      Left            =   1260
      Picture         =   "BorderoPagGeraArqRem.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2145
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   765
      TabIndex        =   2
      Top             =   1635
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Total de Títulos: "
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
      Left            =   1155
      TabIndex        =   7
      Top             =   690
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Títulos Processados:"
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
      Left            =   765
      TabIndex        =   6
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2670
      TabIndex        =   5
      Top             =   1110
      Width           =   1365
   End
   Begin VB.Label TotalTitulos 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2655
      TabIndex        =   4
      Top             =   615
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Geração de Arquivo Remessa de Pagtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   4710
   End
End
Attribute VB_Name = "BorderoPagGeraArqRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sNomeArqParam As String

Public objGeracaoArqCNABPag As New ClassGeracaoArqCNABPag

Dim giCancelaBatch As Integer
Dim giExecutando As Integer ' 0: nao está executando, 1: em andamento

Public Function Trata_Parametros(sDiretorio As String, objBorderoPagEmissao As ClassBorderoPagEmissao) As Long

Dim lErro As Long
Dim objBorderoPagto As New ClassBorderoPagto

On Error GoTo Erro_Trata_Parametros

    giExecutando = ESTADO_PARADO
   
    Set objGeracaoArqCNABPag.objTelaAtualizacao = Me
   
    'Passa para a tela os dados dos Títulos selecionados
    TotalTitulos.Caption = CStr(objBorderoPagEmissao.iQtdeParcelasSelecionadas)
    TitulosProcessados.Caption = "0"

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    With objBorderoPagEmissao
        objBorderoPagto.dtDataEmissao = .dtEmissao
        objBorderoPagto.dtDataEnvio = DATA_NULA
        objBorderoPagto.dtDataVencimento = .dtVencto
        objBorderoPagto.iCodConta = .iCta
        objBorderoPagto.iExcluido = 0
        objBorderoPagto.iNumArqRemessa = 0
        objBorderoPagto.iTipoDeCobranca = .iTipoCobranca
        objBorderoPagto.iTitOutroBanco = .iLiqTitOutroBco
        objBorderoPagto.lNumero = .lNumero
        objBorderoPagto.lNumIntBordero = .lNumeroInt
        objBorderoPagto.sNomeArq = ""
        objGeracaoArqCNABPag.lQuantTitulos = .iQtdeParcelasSelecionadas
        objGeracaoArqCNABPag.sDiretorio = sDiretorio

    End With
    
    Set objGeracaoArqCNABPag.objBorderoPagto = objBorderoPagto

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143826)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    End If
    
    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim objCNABPagRem As New ClassCNABPagRem
Dim lErro As Long, sErro As String

On Error GoTo Erro_BotaoProcessar_Click
   
    If giCancelaBatch <> CANCELA_BATCH Then

        giExecutando = ESTADO_ANDAMENTO
                    
        lErro = objCNABPagRem.BorderosPagto_Criar_ArquivoCNAB(objGeracaoArqCNABPag)
                
        giExecutando = ESTADO_PARADO

        If lErro <> SUCESSO And lErro <> 59190 Then Error 51680
        If lErro = 59190 Then Error 51679 'interrompeu

        'Fecha a tela
        Unload Me
        
    End If

    Exit Sub

Erro_BotaoProcessar_Click:

    Select Case Err

        Case 51679
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")

        Case 51680

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143827)

    End Select

''    If giCancelaBatch <> CANCELA_BATCH Then Call Rotina_ErrosBatch2("Geração de Arquivo de Pagamentos")
    
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

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143828)

    End Select

    giCancelaBatch = CANCELA_BATCH

    Exit Function

End Function

Private Sub Form_Load()
    giCancelaBatch = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objGeracaoArqCNABPag.objTelaAtualizacao = Nothing
    Set objGeracaoArqCNABPag = Nothing
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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
    
    giCancelaBatch = 0
    
    BotaoOK.Enabled = True
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case gErr

        Case 189875 To 189879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189880)

    End Select

''    If giCancelaBatch <> CANCELA_BATCH Then
''        Call Rotina_ErrosBatch2("Geração de Arquivo de Pagamentos")
''    End If

    giCancelaBatch = CANCELA_BATCH

    Exit Sub

End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

