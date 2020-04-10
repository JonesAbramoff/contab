VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PagtosPeriodicosOcx 
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ScaleHeight     =   3960
   ScaleWidth      =   6870
   Begin VB.TextBox Descricao 
      Height          =   315
      Left            =   1305
      TabIndex        =   21
      Top             =   1575
      Width           =   5340
   End
   Begin VB.ComboBox Periodicidade 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2205
      Width           =   1860
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2355
      Picture         =   "PagtosPeriodicosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4620
      ScaleHeight     =   495
      ScaleWidth      =   1995
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Width           =   2055
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   1500
         Picture         =   "PagtosPeriodicosOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   1020
         Picture         =   "PagtosPeriodicosOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   390
         Left            =   555
         Picture         =   "PagtosPeriodicosOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   60
         Picture         =   "PagtosPeriodicosOcx.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   210
      TabIndex        =   4
      Top             =   2160
      Width           =   2820
      Begin MSComCtl2.UpDown UpDownInicio 
         Height          =   300
         Left            =   2190
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicio 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownTermino 
         Height          =   300
         Left            =   2190
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   690
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataTermino 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   690
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataProximo 
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   1125
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownProximo 
         Height          =   300
         Left            =   2190
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1125
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Próximo:"
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
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1155
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Término:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "Início:"
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
         Left            =   405
         TabIndex        =   11
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4890
      TabIndex        =   0
      Top             =   900
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1305
      TabIndex        =   1
      Top             =   900
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   1305
      TabIndex        =   22
      Top             =   240
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.Label CodigoLabel 
      Caption         =   "Código:"
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
      Height          =   210
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   25
      Top             =   270
      Width           =   795
   End
   Begin VB.Label Label8 
      Caption         =   "Descrição:"
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
      Height          =   210
      Left            =   315
      TabIndex        =   24
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Periodicidade:"
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
      Left            =   3495
      TabIndex        =   23
      Top             =   2250
      Width           =   1230
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4350
      TabIndex        =   3
      Top             =   945
      Width           =   465
   End
   Begin VB.Label FornecedorLabel 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   975
      Width           =   1035
   End
End
Attribute VB_Name = "PagtosPeriodicosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private iFornecedorAlterado As Integer

'Browsers
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoPagtosPeriodicos As AdmEvento
Attribute objEventoPagtosPeriodicos.VB_VarHelpID = -1

Const STRING_PAGTOSPERIODICOS_DESCRICAO = 250
Const STRING_PERIODICIDADESCPR_DESCRICAO = 50

Type typePagtosPeriodicos

     dtInicio As Date
     dtTermino As Date
     dtProximo As Date
     iFilial As Integer
     iFilialEmpresa As Integer
     iPeriodicidade As Integer
     lCodigo As Long
     sDescricao As String
     lFornecedor As Long
    
End Type

Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFornecedorAlterado = 0

    'browse
    Set objEventoFornecedor = New AdmEvento
    Set objEventoPagtosPeriodicos = New AdmEvento
    
    'Preenche Lista da Combobox com os dados da tabela PeriodicidadeCPR
    lErro = Carrega_ComboPeriodicidade()
    If lErro <> SUCESSO Then gError 122562
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 122562
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164258)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub


Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objPagtosPeriodicos As New ClassPagtosPeriodicos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "PagtosPeriodicos"

    lErro = Move_Tela_Memoria(objPagtosPeriodicos)
    If lErro <> SUCESSO Then gError 122563

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Inicio", objPagtosPeriodicos.dtInicio, 0, "Inicio"
    colCampoValor.Add "Termino", objPagtosPeriodicos.dtTermino, 0, "Termino"
    colCampoValor.Add "Proximo", objPagtosPeriodicos.dtProximo, 0, "Proximo"
    colCampoValor.Add "FilialEmpresa", objPagtosPeriodicos.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Filial", objPagtosPeriodicos.iFilial, 0, "Filial"
    colCampoValor.Add "Periodicidade", objPagtosPeriodicos.iPeriodicidade, 0, "Periodicidade"
    colCampoValor.Add "Fornecedor", objPagtosPeriodicos.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Codigo", objPagtosPeriodicos.lCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objPagtosPeriodicos.sDescricao, STRING_PAGTOSPERIODICOS_DESCRICAO, "Descricao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 122563

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164259)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPagtosPeriodicos As ClassPagtosPeriodicos) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de Pagtos Periódicos

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objPagtosPeriodicos Is Nothing) Then

        lErro = PagtosPeriodicos_Le(objPagtosPeriodicos)
        If lErro <> SUCESSO And lErro <> 122566 Then gError 122564

        If lErro = SUCESSO Then

            lErro = Traz_PagtosPeriodicos_Tela(objPagtosPeriodicos)
            If lErro <> SUCESSO Then gError 122565
                        
        Else
        
            Codigo.Text = CStr(objPagtosPeriodicos.lCodigo)
            
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 122564, 122565

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164260)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPagtosPeriodicos As New ClassPagtosPeriodicos

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objPagtosPeriodicos.dtInicio = colCampoValor.Item("Inicio").vValor
    objPagtosPeriodicos.dtTermino = colCampoValor.Item("Termino").vValor
    objPagtosPeriodicos.dtProximo = colCampoValor.Item("Proximo").vValor
    objPagtosPeriodicos.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objPagtosPeriodicos.iFilial = colCampoValor.Item("Filial").vValor
    objPagtosPeriodicos.iPeriodicidade = colCampoValor.Item("Periodicidade").vValor
    objPagtosPeriodicos.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objPagtosPeriodicos.lCodigo = colCampoValor.Item("Codigo").vValor
    objPagtosPeriodicos.sDescricao = colCampoValor.Item("Descricao").vValor
    
    lErro = Traz_PagtosPeriodicos_Tela(objPagtosPeriodicos)
    If lErro <> SUCESSO Then gError 122567

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 122567

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164261)

    End Select

    Exit Sub

End Sub

Function Carrega_ComboPeriodicidade() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As New AdmCodigoNome

   
On Error GoTo Erro_Carrega_ComboPeriodicidade
    
    'Preenche ColCodigoDescricao com dados da tabela PeriodicidadesCPR
    lErro = CF("Cod_Nomes_Le", "PeriodicidadesCPR", "Codigo", "Descricao", STRING_PERIODICIDADESCPR_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 122568
     
    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na List da Combo Periodicidade
        Periodicidade.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        Periodicidade.ItemData(Periodicidade.NewIndex) = objCodDescricao.iCodigo

    Next
     
    Carrega_ComboPeriodicidade = SUCESSO
    
    Exit Function
    
Erro_Carrega_ComboPeriodicidade:

    Carrega_ComboPeriodicidade = gErr
    
    Select Case gErr
    
        Case 122568
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164262)

    End Select

    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPagtosPeriodicos As New ClassPagtosPeriodicos

On Error GoTo Erro_BotaoExcluir_Click
    
    'Verifica se o Código do Recebimento está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 122569
    
    objPagtosPeriodicos.lCodigo = CLng(Codigo.Text)
    
    objPagtosPeriodicos.iFilialEmpresa = giFilialEmpresa
    
    'Envia mensagem pedindo confirmação de exclusão
    If Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PAGTOPERIODICO", objPagtosPeriodicos.lCodigo) = vbYes Then
        
        'Exclui o Recebimento da tabela
        lErro = PagtosPeriodicos_Exclui(objPagtosPeriodicos)
        If lErro <> SUCESSO Then gError 122570
                
        'Limpa a tela
        Call Limpa_Tela_PagtosPeriodicos
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr
                   
        Case 122569
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
    
        Case 122570
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164263)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 122571

    'Limpa a tela
    Call Limpa_Tela_PagtosPeriodicos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 122571

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164264)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Verifica se dados de PagtosPeriodicos necessários foram preenchidos
'Grava RecebPeriodico no BD

Dim lErro As Long
Dim objPagtosPeriodicos As New ClassPagtosPeriodicos

On Error GoTo Erro_Gravar_Registro
    
    'Verifica se o Codigo do Recebimento foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 122572

    'Verifica se o Código do Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 122573
    
    'Verifica se a Filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 122574
    
    'Verifica se a Descricao foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 122575
    
    'Verifica se a Periodicidade foi preenchida
    If Len(Trim(Periodicidade.Text)) = 0 Then gError 122576
    
    'Verifica se a Data de Inicio foi preenchida
    If StrParaDate(DataInicio.Text) = DATA_NULA Then gError 122577
        
    'Verifica se a Data do Proximo Recebimento foi preenchida
    If StrParaDate(DataProximo.Text) = DATA_NULA Then gError 122578
               
    'Verifica se a Data do Proximo Recebimento está entre DtInicio e DtFIm
    If (StrParaDate(DataProximo.Text) < StrParaDate(DataInicio.Text)) Or (StrParaDate(DataProximo.Text) > StrParaDate(DataTermino.Text) And StrParaDate(DataTermino.Text) <> DATA_NULA) Then gError 122579
    
    'Verifica se DataTermino é menor que DataInicio
    If StrParaDate(DataTermino.Text) < StrParaDate(DataInicio.Text) And StrParaDate(DataTermino.Text) <> DATA_NULA Then gError 122580
        
    lErro = Move_Tela_Memoria(objPagtosPeriodicos)
    If lErro <> SUCESSO Then gError 122581
    
    'Realiza Inclusão/Alteração no BD
    lErro = PagtosPeriodicos_Grava(objPagtosPeriodicos)
    If lErro <> SUCESSO Then gError 122582

    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case 122572
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 122573
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 122574
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 122575
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 122576
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADE_NAO_PREENCHIDO", gErr)
        
        Case 122580
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_MAIOR_DATAINICIO", gErr)

        
        Case 122577
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
        
        Case 122578
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPROXIMO_NAO_PREENCHIDA", gErr)
        
        Case 122579
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPROXIMO_FORA_INTERVALO", gErr)
        
        Case 122581, 122582
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164265)

    End Select
    
    Exit Function

End Function

Function Traz_PagtosPeriodicos_Tela(objPagtosPeriodicos As ClassPagtosPeriodicos) As Long

On Error GoTo Erro_Traz_PagtosPeriodicos_Tela

    'Mostra os dados na tela
    Call DateParaMasked(DataInicio, objPagtosPeriodicos.dtInicio)
    Call DateParaMasked(DataTermino, objPagtosPeriodicos.dtTermino)
    Call DateParaMasked(DataProximo, objPagtosPeriodicos.dtProximo)
    Call Combo_Seleciona_ItemData(Periodicidade, CLng(objPagtosPeriodicos.iPeriodicidade))
   
    Fornecedor.Text = objPagtosPeriodicos.lFornecedor
    Call Fornecedor_Validate(bSGECancelDummy)
    
    Filial.Text = objPagtosPeriodicos.iFilial
    Call Filial_Validate(bSGECancelDummy)
    
    Codigo.Text = objPagtosPeriodicos.lCodigo
    Descricao.Text = objPagtosPeriodicos.sDescricao
        
    iAlterado = 0

    Exit Function

Erro_Traz_PagtosPeriodicos_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164266)

    End Select

End Function

Private Function Move_Tela_Memoria(objPagtosPeriodicos As ClassPagtosPeriodicos) As Long
'Lê os dados que estão na tela PagtosPeriodicos e os coloca em objPagtosPeriodicos

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria
    
    If Len(Trim(CStr(Codigo.Text))) > 0 Then
    
        objPagtosPeriodicos.dtInicio = MaskedParaDate(DataInicio)
        objPagtosPeriodicos.dtTermino = MaskedParaDate(DataTermino)
        objPagtosPeriodicos.dtProximo = MaskedParaDate(DataProximo)
        objPagtosPeriodicos.iFilialEmpresa = giFilialEmpresa
                
        If Len(Trim(CStr(Fornecedor.Text))) > 0 Then
                
            'O código do Fornecedor
            objFornecedor.sNomeReduzido = Fornecedor.Text
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO Then gError 122583
    
            objPagtosPeriodicos.lFornecedor = objFornecedor.lCodigo
            
        End If

        'A filial do Fornecedor
        objPagtosPeriodicos.iFilial = Codigo_Extrai(Filial.Text)
    
        objPagtosPeriodicos.iPeriodicidade = Codigo_Extrai(Periodicidade.Text)
        objPagtosPeriodicos.lCodigo = StrParaLong(Codigo.Text)
        objPagtosPeriodicos.sDescricao = Descricao.Text
        
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 122583

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164267)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_PagtosPeriodicos()
'Limpa todos os campos da tela PagtosPeriodicos
        
    Call Limpa_Tela(Me)
    
    Periodicidade.ListIndex = -1
    Filial.ListIndex = -1
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código do proximo RecebPeriodico
    lErro = PagtosPeriodicos_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 122584
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 122584
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164268)
    
    End Select

    Exit Sub
    
End Sub

Function PagtosPeriodicos_Automatico(lNumIntAuto As Long)
'Funçao que gera automaticamente os numeros internos para Notas Fiscais a Pagar
'AVISO: Esta função deve ser chamada dentro de uma transação

    PagtosPeriodicos_Automatico = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_PAGTOPERIODICO", "PagtosPeriodicos", "Codigo", lNumIntAuto)

End Function

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CodigoLabel_Click()

Dim objPagtosPeriodicos As New ClassPagtosPeriodicos
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objPagtosPeriodicos.lCodigo = Codigo.Text
    'Chama Tela PagtosPeriodicosLista
    Call Chama_Tela("PagtosPeriodicosLista", colSelecao, objPagtosPeriodicos, objEventoPagtosPeriodicos)
    
End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate
        
    'Critica a data digitada
    lErro = Data_Critica(DataInicio.Text)
    If lErro <> SUCESSO Then gError 122585
    
    Exit Sub
    
Erro_DataInicio_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 122585
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164269)
            
    End Select

Exit Sub
    
End Sub

Private Sub DataTermino_Validate(Cancel As Boolean)

    Dim lErro As Long

On Error GoTo Erro_DataTermino_Validate
        
    'Critica a data digitada
    lErro = Data_Critica(DataTermino.Text)
    If lErro <> SUCESSO Then gError 122586
    
    Exit Sub
    
Erro_DataTermino_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122586
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164270)
            
    End Select

Exit Sub

End Sub

Private Sub DataProximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataProximo_Validate
        
    'Critica a data digitada
    lErro = Data_Critica(DataProximo.Text)
    If lErro <> SUCESSO Then gError 122587
    
    Exit Sub
    
Erro_DataProximo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 122587
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164271)
            
    End Select

Exit Sub
    
End Sub


Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
  
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 122588

    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 122589

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 122593
        
        'Se não encontrou
        If lErro = 17660 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 122590
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
             
             vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

             If vbMsgRes = vbYes Then
                 Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
             Else
                gError 122591
             End If
                
        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 122592

    Exit Sub

Erro_Filial_Validate:

    Cancel = True
    
    Select Case gErr

       Case 122588, 122593, 122590

       Case 122589
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

       Case 122591
      
       Case 122592
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164272)

    End Select

    Exit Sub
    
End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 122594

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 122595

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
     
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Limpa Combo de Filial
            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 122594, 122595

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164273)

    End Select

    Exit Sub

End Sub

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = 1
    
End Sub

Function PagtosPeriodicos_Le(ByVal objPagtosPeriodicos As ClassPagtosPeriodicos) As Long
'Carrega em objPagtosPeriodicos o RecebPeriodico que está na
'tabela PagtosPeriodicos e que possui FilialEmpresa e Codigo contido em
'objPagtosPeriodicos

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String

Dim tPagtosPeriodicos As typePagtosPeriodicos

On Error GoTo Erro_PagtosPeriodicos_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 122596

    tPagtosPeriodicos.sDescricao = String(STRING_PAGTOSPERIODICOS_DESCRICAO, 0)

    With tPagtosPeriodicos
        lErro = Comando_Executar(lComando, "SELECT Descricao, Fornecedor, Filial, Periodicidade, Inicio, Termino, Proximo FROM PagtosPeriodicos WHERE Codigo=? AND FilialEmpresa =?", .sDescricao, .lFornecedor, .iFilial, .iPeriodicidade, .dtInicio, .dtTermino, .dtProximo, objPagtosPeriodicos.lCodigo, objPagtosPeriodicos.iFilialEmpresa)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 122597

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122598

    If lErro = AD_SQL_SEM_DADOS Then gError 122566

    With tPagtosPeriodicos
    
        objPagtosPeriodicos.dtInicio = .dtInicio
        objPagtosPeriodicos.dtTermino = .dtTermino
        objPagtosPeriodicos.dtProximo = .dtProximo
        objPagtosPeriodicos.iFilial = .iFilial
        objPagtosPeriodicos.iPeriodicidade = .iPeriodicidade
        objPagtosPeriodicos.lFornecedor = .lFornecedor
        objPagtosPeriodicos.sDescricao = .sDescricao
        
    End With
    
    Call Comando_Fechar(lComando)

    PagtosPeriodicos_Le = SUCESSO

    Exit Function

Erro_PagtosPeriodicos_Le:

    PagtosPeriodicos_Le = gErr

    Select Case gErr

        Case 122596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 122597, 122598
           Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTOSPERIODICOS", gErr)

        Case 122566

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164274)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PagtosPeriodicos_Grava(ByVal objPagtosPeriodicos As ClassPagtosPeriodicos) As Long
'Inclui/Altera registro no BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tPagtosPeriodicos As typePagtosPeriodicos

On Error GoTo Erro_PagtosPeriodicos_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 122599
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 122600
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM PagtosPeriodicos WHERE FilialEmpresa=? AND Codigo=?", 0, tPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122601
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122602
    
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 122603
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PagtosPeriodicos SET Descricao = ?, Fornecedor = ?, Filial = ?, Periodicidade = ?, Inicio = ?, Termino = ?, Proximo = ?", alComando(1), objPagtosPeriodicos.sDescricao, objPagtosPeriodicos.lFornecedor, objPagtosPeriodicos.iFilial, objPagtosPeriodicos.iPeriodicidade, objPagtosPeriodicos.dtInicio, objPagtosPeriodicos.dtTermino, objPagtosPeriodicos.dtProximo)
        If lErro <> AD_SQL_SUCESSO Then gError 122604
    
    Else
            
        'Verifica se FilialEmpresa, Descricao, Fornecedor e Filial já constam em algum registro
        lErro = Comando_Executar(alComando(2), "SELECT FilialEmpresa FROM PagtosPeriodicos WHERE FilialEmpresa=? AND Descricao=? AND Fornecedor=? AND Filial=?", tPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.sDescricao, objPagtosPeriodicos.lFornecedor, objPagtosPeriodicos.iFilial)
        'se consta em algum registro --> Erro
        If lErro <> AD_SQL_SUCESSO Then gError 122605
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122606
    
        If lErro = AD_SQL_SUCESSO Then gError 122607
    
        lErro = Comando_Executar(alComando(3), "INSERT INTO PagtosPeriodicos (FilialEmpresa,Codigo,Descricao,Fornecedor,Filial,Periodicidade,Inicio,Termino,Proximo) VALUES (?,?,?,?,?,?,?,?,?)", objPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.lCodigo, objPagtosPeriodicos.sDescricao, objPagtosPeriodicos.lFornecedor, objPagtosPeriodicos.iFilial, objPagtosPeriodicos.iPeriodicidade, objPagtosPeriodicos.dtInicio, objPagtosPeriodicos.dtTermino, objPagtosPeriodicos.dtProximo)
        If lErro <> AD_SQL_SUCESSO Then gError 122608
    
    End If
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 122609

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PagtosPeriodicos_Grava = SUCESSO

    Exit Function

Erro_PagtosPeriodicos_Grava:

    PagtosPeriodicos_Grava = gErr

    Select Case gErr

        Case 122599
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 122600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 122601, 122602, 122605, 122606
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTOSPERIODICOS", gErr)

        Case 122603
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PAGTOSPERIODICOS", gErr, objPagtosPeriodicos.lCodigo)
            
        Case 122604
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PAGTOSPERIODICOS", gErr, objPagtosPeriodicos.lCodigo)

        Case 122607
            Call Rotina_Erro(vbOKOnly, "ERRO_PAGTOPERIODICO_JA_EXISTENTE", gErr, objPagtosPeriodicos.sDescricao, objPagtosPeriodicos.lFornecedor, objPagtosPeriodicos.iFilial)

        Case 122608
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PAGTOSPERIODICOS", gErr, objPagtosPeriodicos.lCodigo)

        Case 122609
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164275)

    End Select
    
    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function PagtosPeriodicos_Exclui(ByVal objPagtosPeriodicos As ClassPagtosPeriodicos) As Long
'Exclui o registro de PagtosPeriodicos que possui Código e FilialEmpresa igual ao
'que está contido em objPagtosPeriodicos

Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim tPagtosPeriodicos As typePagtosPeriodicos

On Error GoTo Erro_PagtosPeriodicos_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 122610
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 122611
    Next
        
    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM PagtosPeriodicos WHERE Codigo=? AND FilialEmpresa= ?", 0, tPagtosPeriodicos.iFilialEmpresa, objPagtosPeriodicos.lCodigo, objPagtosPeriodicos.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 122612

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122613

    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 122614
        
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM PagtosPeriodicos", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 122615

    End If

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 122616
            
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    PagtosPeriodicos_Exclui = SUCESSO
    
    Exit Function

Erro_PagtosPeriodicos_Exclui:
    
    PagtosPeriodicos_Exclui = gErr
    
    Select Case gErr
    
        Case 122610
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 122611
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
  
        Case 122612, 122613
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTOSPERIODICOS", gErr)
        
        Case 122614
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PAGTOSPERIODICOS", gErr, objPagtosPeriodicos.lCodigo)
        
        Case 122615
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PAGTOSPERIODICOS", gErr, objPagtosPeriodicos.lCodigo)

        Case 122616
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164276)
    
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoFornecedor = Nothing
    Set objEventoPagtosPeriodicos = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pagamentos Periódicos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagtosPeriodicos"
    
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

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma se deseja salvar alterações
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 122617

    'Limpa a tela
    Call Limpa_Tela_PagtosPeriodicos
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 122617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164277)
        
    End Select
    
    Exit Sub


End Sub

Private Sub FilialFornecedor_Change()

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama Tela FornecedoresLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoPagtosPeriodicos_evSelecao(obj1 As Object)

Dim objPagtosPeriodicos As ClassPagtosPeriodicos

    Set objPagtosPeriodicos = obj1
    
    Call Traz_PagtosPeriodicos_Tela(objPagtosPeriodicos)

    Me.Show

    Exit Sub

End Sub


Private Sub Periodicidade_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDownInicio_DownClick()

Dim lErro As Long
'Dim sData As String

On Error GoTo Erro_UpDownInicio_DownClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataInicio, DIMINUI_DATA)
    If lErro Then gError 122618

    Exit Sub

Erro_UpDownInicio_DownClick:

    Select Case gErr

        Case 122618

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownInicio_UpClick()

Dim lErro As Long
'Dim sData As String

On Error GoTo Erro_UpDownInicio_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataInicio, AUMENTA_DATA)
    If lErro Then gError 122619

    Exit Sub

Erro_UpDownInicio_UpClick:

    Select Case gErr

        Case 122619

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164279)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownTermino_DownClick()

Dim lErro As Long
'Dim sData As String

On Error GoTo Erro_UpDownTermino_DownClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataTermino, DIMINUI_DATA)
    If lErro Then gError 122620

    Exit Sub

Erro_UpDownTermino_DownClick:

    Select Case gErr

        Case 122620

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164280)

    End Select

    Exit Sub

End Sub

Private Sub UpDownTermino_UpClick()

Dim lErro As Long
'Dim sData As String

On Error GoTo Erro_UpDownTermino_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataTermino, AUMENTA_DATA)
    If lErro Then gError 122621

    Exit Sub

Erro_UpDownTermino_UpClick:

    Select Case gErr

        Case 122621

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164281)

    End Select

    Exit Sub

End Sub

Private Sub UpDownProximo_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownProximo_DownClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataProximo, DIMINUI_DATA)
    If lErro Then gError 122622

    Exit Sub

Erro_UpDownProximo_DownClick:

    Select Case gErr

        Case 122622

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164282)

    End Select

    Exit Sub

End Sub

Private Sub UpDownProximo_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownProximo_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataProximo, AUMENTA_DATA)
    If lErro Then gError 122623

    Exit Sub

Erro_UpDownProximo_UpClick:

    Select Case gErr

        Case 122623

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164283)

    End Select

    Exit Sub
    
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



