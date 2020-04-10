VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GeracaoArqRemCobr2 
   Caption         =   "Geração de Arquivo de Remessa - CNAB"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton BotaoInterromper 
      Caption         =   "Interromper Processamento"
      Enabled         =   0   'False
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
      Left            =   990
      TabIndex        =   3
      Top             =   2085
      Width           =   3105
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   1575
      ScaleHeight     =   495
      ScaleWidth      =   1830
      TabIndex        =   0
      Top             =   2625
      Width           =   1890
      Begin VB.CommandButton BotaoProcessar 
         Caption         =   "Processar"
         Enabled         =   0   'False
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   1050
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1305
         Picture         =   "GeracaoArqRemCobr2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   1185
      TabIndex        =   4
      Top             =   1515
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Borderos:"
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
      Left            =   1140
      TabIndex        =   9
      Top             =   660
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Borderos Processados:"
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
      Left            =   765
      TabIndex        =   8
      Top             =   1140
      Width           =   1965
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Processamento dos Borderos"
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
      Left            =   1050
      TabIndex        =   7
      Top             =   165
      Width           =   3105
   End
   Begin VB.Label BorderosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2775
      TabIndex        =   6
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label TotalBorderos 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2775
      TabIndex        =   5
      Top             =   630
      Width           =   1350
   End
End
Attribute VB_Name = "GeracaoArqRemCobr2"
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO
        
    BorderosProcessados.Caption = "0"

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160773)
    
    End Select
    
    giCancelaBatch = CANCELA_BATCH
    
    Exit Sub

End Sub

''Function Trata_Parametros(objCobrancaEletronica As ClassCobrancaEletronica) As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_Trata_Parametros
''
''    giCancelaBatch = 0
''    giExecutando = ESTADO_PARADO
''
''    Set gobjCobrancaEletronica = objCobrancaEletronica
''
''    'Passa para a tela os dados dos Borderos selecionados
''    TotalBorderos.Caption = CStr(objCobrancaEletronica.colBorderos.Count)
''    BorderosProcessados.Caption = "0"
''
''    ProgressBar1.Min = 0
''    ProgressBar1.Max = 100
''
''    Trata_Parametros = SUCESSO
''
''    Exit Function
''
''Erro_Trata_Parametros:
''
''    Trata_Parametros = Err
''
''    Select Case Err
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160774)
''
''    End Select
''
''    giCancelaBatch = CANCELA_BATCH
''
''    Exit Function
''
''End Function



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
    
    giExecutando = ESTADO_ANDAMENTO
    
    Set gobjCobrancaEletronica.objTelaAtualizacao = Me
    
    lErro = CF("CobrancaEletronica_Criar_ArquivoRemessa", gobjCobrancaEletronica)
            
    giExecutando = ESTADO_PARADO

    BotaoInterromper.Enabled = False

    'Se o processamento foi cancelado
    If giCancelaBatch = CANCELA_BATCH Then
        
        'Exibe aviso
        Call Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
        'Fecha a tela
        Unload Me
        'Sai da função
        Exit Sub
    
    End If
    
    
    If lErro <> SUCESSO And lErro <> 59190 Then Error 51680
    If lErro = 59190 Then Error 51679 'interrompeu

    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoProcessar_Click:

    Select Case Err

        Case 51679
            Call Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")

        Case 51680

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160775)

    End Select

''    If giCancelaBatch <> CANCELA_BATCH Then Call Rotina_ErrosBatch2("Geração de Arquivo de Cobrança")
    
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

        iProcessados = CInt(BorderosProcessados.Caption)
        iTotal = CInt(TotalBorderos.Caption)

        iProcessados = iProcessados + iNumProc
        BorderosProcessados.Caption = CStr(iProcessados)

        ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)

        giExecutando = ESTADO_ANDAMENTO

    End If

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160776)

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
    
End Sub

Private Sub BotaoInterromper_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    
    End If
    
    'Fecha a tela
    Unload Me

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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub BorderosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(BorderosProcessados, Source, X, Y)
End Sub

Private Sub BorderosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(BorderosProcessados, Button, Shift, X, Y)
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
''    giCancelaBatch = 0
''
''    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
''    If lErro <> SUCESSO Then gError 189876
''
''    lErro = CF("Retorna_ColFiliais")
''    If lErro <> SUCESSO Then gError 189877
''
''    GL_lUltimoErro = SUCESSO
    
    giCancelaBatch = 0
    
    'Passa para a tela os dados dos Borderos selecionados
    TotalBorderos.Caption = CStr(gobjCobrancaEletronica.colBorderos.Count)
    
    BotaoProcessar.Enabled = True
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case gErr

        Case 189875 To 189879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189880)

    End Select

''    If giCancelaBatch <> CANCELA_BATCH Then
''        Call Rotina_ErrosBatch2("Geração de Arquivo de Cobrança")
''    End If

    giCancelaBatch = CANCELA_BATCH

    Exit Sub

End Sub

Private Sub TotalBorderos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalBorderos, Source, X, Y)
End Sub

Private Sub TotalBorderos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalBorderos, Button, Shift, X, Y)
End Sub

