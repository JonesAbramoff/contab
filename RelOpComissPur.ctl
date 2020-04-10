VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpComissPur 
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LockControls    =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   7020
   Begin VB.Frame Frame6 
      Caption         =   "Layout"
      Height          =   495
      Left            =   135
      TabIndex        =   41
      Top             =   5250
      Width           =   6780
      Begin VB.OptionButton OptConsolidado 
         Caption         =   "Consolidado"
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
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Optnf 
         Caption         =   "Detalhado nota a nota"
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
         Left            =   1755
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Títulos"
      Height          =   690
      Left            =   4725
      TabIndex        =   40
      Top             =   1410
      Width           =   2190
      Begin VB.OptionButton OptVcto 
         Caption         =   "Pelo Vencimento"
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
         TabIndex        =   6
         Top             =   450
         Width           =   1950
      End
      Begin VB.OptionButton OptBx 
         Caption         =   "Pela Baixa"
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
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   1950
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data da Emissão"
      Height          =   690
      Left            =   120
      TabIndex        =   37
      Top             =   1410
      Width           =   4575
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   330
         Left            =   1590
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmiInicial 
         Height          =   312
         Left            =   624
         TabIndex        =   3
         Top             =   252
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   330
         Left            =   3405
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmiFinal 
         Height          =   330
         Left            =   2460
         TabIndex        =   4
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "De:"
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
         Height          =   252
         Left            =   240
         TabIndex        =   39
         Top             =   300
         Width           =   396
      End
      Begin VB.Label Label2 
         Caption         =   "Até:"
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
         Height          =   252
         Left            =   2040
         TabIndex        =   38
         Top             =   300
         Width           =   456
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Região de Venda"
      Height          =   1755
      Left            =   135
      TabIndex        =   36
      Top             =   3495
      Width           =   6765
      Begin VB.ListBox ListRegioes 
         Height          =   1410
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   4980
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   525
         Left            =   5145
         Picture         =   "RelOpComissPur.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   255
         Width           =   1530
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   525
         Left            =   5145
         Picture         =   "RelOpComissPur.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   900
         Width           =   1530
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vendedores"
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   2865
      Width           =   6780
      Begin VB.OptionButton OptVendIndir 
         Caption         =   "Vendas Indiretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   10
         Top             =   180
         Width           =   1800
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   4545
         TabIndex        =   11
         Top             =   210
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.OptionButton OptVendDir 
         Caption         =   "Vendas Diretas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   3630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   690
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   6795
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   7
         Top             =   255
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   4020
         TabIndex        =   8
         Top             =   255
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   3600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   315
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpComissPur.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpComissPur.ctx":2356
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpComissPur.ctx":24E0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpComissPur.ctx":2A12
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data da Baixa/Vencimento"
      Height          =   690
      Left            =   120
      TabIndex        =   27
      Top             =   690
      Width           =   4575
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1590
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   312
         Left            =   624
         TabIndex        =   1
         Top             =   252
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   3405
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   330
         Left            =   2460
         TabIndex        =   2
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFim 
         Caption         =   "Até:"
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
         Height          =   252
         Left            =   2040
         TabIndex        =   29
         Top             =   300
         Width           =   456
      End
      Begin VB.Label dIni 
         Caption         =   "De:"
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
         Height          =   252
         Left            =   240
         TabIndex        =   28
         Top             =   300
         Width           =   396
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4740
      Picture         =   "RelOpComissPur.ctx":2B90
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   765
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpComissPur.ctx":2C92
      Left            =   870
      List            =   "RelOpComissPur.ctx":2C94
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   2916
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   195
      TabIndex        =   30
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "RelOpComissPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Dim giClienteInicial As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoVendedor = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 47955
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 47955
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170179)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpLancData(sCliente_I As String, sCliente_F As String, iTipoVend As Integer, iDet As Integer, iBx As Integer) As Long
'a data inicial não pode ser maior que a data final

Dim lErro As Long
Dim iIndice As Integer, iAchou As Integer

On Error GoTo Erro_Critica_Datas_RelOpLancData
    
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataInicial.ClipText) <> 0 And Len(DataFinal.ClipText) <> 0 Then
        'data inicial não pode ser maior que a data final
        If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 47956
    End If
            
    'Critica data se  a Inicial e a Final estiverem Preenchida
    If Len(DataEmiInicial.ClipText) <> 0 And Len(DataEmiFinal.ClipText) <> 0 Then
        'data inicial não pode ser maior que a data final
        If CDate(DataEmiInicial.Text) > CDate(DataEmiFinal.Text) Then gError 47956
    End If
    
    'critica Cliente Inicial e Final
    If ClienteInicial.Text <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_I = ""
    End If
    
    If ClienteFinal.Text <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_F = ""
    End If
            
    If sCliente_I <> "" And sCliente_F <> "" Then
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 71355
    End If
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
    If Optnf.Value Then
        iDet = MARCADO
    Else
        iDet = DESMARCADO
    End If
    
    If OptBx.Value Then
        iBx = MARCADO
    Else
        iBx = DESMARCADO
    End If
    
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
        
    Next
       
    If iAchou = 0 Then gError 207095
            
    Critica_Datas_RelOpLancData = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpLancData:

    Critica_Datas_RelOpLancData = gErr

    Select Case gErr
    
        Case 47956
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 71355
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus

        Case 207095
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_ROTA_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170180)

    End Select
    
    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim iTipoVend As Integer, iIndice As Integer
Dim lNumIntRel As Long, iNRegiao As Integer
Dim sRegiao As String, sListCount As String
Dim iBx As Integer, iDet As Integer
Dim colReg As New Collection, bTodasRegioes As Boolean

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpLancData(sCliente_I, sCliente_F, iTipoVend, iDet, iBx)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataInicial.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", DataEmiInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Trim(DataFinal.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", DataEmiFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                 
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NDET", CStr(iDet))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NBX", CStr(iBx))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    bTodasRegioes = True
    sListCount = "0"
    For iIndice = 0 To ListRegioes.ListCount - 1
        If Not ListRegioes.Selected(iIndice) Then
            bTodasRegioes = False
            Exit For
        End If
    Next
    
    If Not bTodasRegioes Then
        iNRegiao = 1
        'Percorre toda a Lista
        For iIndice = 0 To ListRegioes.ListCount - 1
            If ListRegioes.Selected(iIndice) Then
                sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
                
                colReg.Add Codigo_Extrai(ListRegioes.List(iIndice))
                
                'Inclui todas as Regioes que foram slecionados
                lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
                If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                iNRegiao = iNRegiao + 1
            End If
        Next
        sListCount = iNRegiao - 1
    End If
    
    'Inclui o numero de regiões selecionadas na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("RelComissPur_Prepara", colReg, giFilialEmpresa, StrParaDate(DataEmiInicial.Text), StrParaDate(DataEmiFinal.Text), StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), StrParaLong(sCliente_I), StrParaLong(sCliente_F), iTipoVend, Codigo_Extrai(Vendedor.Text), iBx, iDet, lNumIntRel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170181)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String, iIndice As Integer
Dim sListCount As String, iIndiceRel As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataInicial, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DBXFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("DEMIINIC", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataEmiInicial, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMIFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataEmiFinal, StrParaDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If sParam <> "0" Then
        Vendedor.Text = CInt(sParam)
        Call Vendedor_Validate(bSGECancelDummy)
    End If

    lErro = objRelOpcoes.ObterParametro("NTIPOVEND", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = VENDEDOR_DIRETO Then
        OptVendDir.Value = True
    Else
        OptVendIndir.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("NDET", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = MARCADO Then
        Optnf.Value = True
        OptConsolidado.Value = False
    Else
        Optnf.Value = False
        OptConsolidado.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("NBX", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) = MARCADO Then
        OptBx.Value = True
        OptVcto.Value = False
    Else
        OptBx.Value = False
        OptVcto.Value = True
    End If
    
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next
    
    'Obtem o numero de Regioes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCOUNT", sListCount)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    'Percorre toda a Lista
    
    For iIndice = 0 To ListRegioes.ListCount - 1
        
        If sListCount = "0" Then
            ListRegioes.Selected(iIndice) = True
        Else
            'Percorre todas as Regieos que foram slecionados
            For iIndiceRel = 1 To StrParaInt(sListCount)
                lErro = objRelOpcoes.ObterParametro("NLIST" & SEPARADOR & iIndiceRel, sParam)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Se o cliente não foi excluido
                If sParam = Codigo_Extrai(ListRegioes.List(iIndice)) Then
                    'Marca as Regioes que foram gravados
                    ListRegioes.Selected(iIndice) = True
                End If
            Next
        End If
    Next

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170182)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 47966

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        giClienteInicial = 1
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 47966
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170183)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170184)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 47970

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 47970
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170185)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Call Define_Padrao
        
    giClienteInicial = 1
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170186)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170187)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiFinal)

End Sub

Private Sub DataEmiFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiFinal_Validate

    If Len(DataEmiFinal.ClipText) > 0 Then

        lErro = Data_Critica(DataEmiFinal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataEmiFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170187)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170188)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiInicial)

End Sub

Private Sub DataEmiInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiInicial_Validate

    If Len(DataEmiInicial.ClipText) > 0 Then

        lErro = Data_Critica(DataEmiInicial.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataEmiInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170188)

    End Select

    Exit Sub

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 71356

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 71356
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170189)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 71357

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 71357
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170190)

    End Select

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteInicial.Text = CStr(objCliente.lCodigo)
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = CStr(objCliente.lCodigo)
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    lErro = CarregaList_Regioes
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Define_Padrao

    giClienteInicial = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170191)

    End Select

    Unload Me

    Exit Sub

End Sub

Sub Define_Padrao()

    OptVendDir.Value = True
    OptBx.Value = True
    OptConsolidado.Value = True
    Call Limpa_ListRegioes

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170192)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170193)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170194)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170195)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_DownClick

    lErro = Data_Up_Down_Click(DataEmiInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown3_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170192)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_UpClick

    lErro = Data_Up_Down_Click(DataEmiInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown3_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170193)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_DownClick

    lErro = Data_Up_Down_Click(DataEmiFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown4_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170194)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_UpClick

    lErro = Data_Up_Down_Click(DataEmiFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown4_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataEmiFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170195)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String
Dim lErro As Long, iIndice As Integer
Dim sSub As String, iCount As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170196)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NFISCAL_DEVOLUCAO
    Set Form_Load_Ocx = Me
    Caption = "Comissões por Meta"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpComissPur"
    
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

Public Sub Unload(objme As Object)
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

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Preenche com o Vendedor da tela
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207090

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        ListRegioes.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ListRegioes.ItemData(ListRegioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    CarregaList_Regioes = SUCESSO

    Exit Function

Erro_CarregaList_Regioes:

    CarregaList_Regioes = gErr

    Select Case gErr

        Case 207900

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_ListRegioes()

Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Public Function RetiraNomes_Sel(colRegioes As Collection) As Long
'Retira da combo todos os nomes que não estão selecionados

Dim iIndice As Integer
Dim lCodRegiao As Long

    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            lCodRegiao = LCodigo_Extrai(ListRegioes.List(iIndice))
            colRegioes.Add lCodRegiao
        End If
    Next
    
End Function
