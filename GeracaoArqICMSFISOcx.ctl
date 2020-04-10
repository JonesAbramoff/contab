VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoArqICMSFISOcx 
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   7410
   ScaleWidth      =   5220
   Begin VB.CheckBox Entradas 
      Caption         =   "Entradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1980
      TabIndex        =   40
      Top             =   6945
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CheckBox Saidas 
      Caption         =   "Saídas"
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
      Left            =   375
      TabIndex        =   39
      Top             =   6998
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventário - Registro 74"
      Height          =   780
      Left            =   165
      TabIndex        =   32
      Top             =   6000
      Width           =   4890
      Begin MSComCtl2.UpDown UpDownReg74Ini 
         Height          =   300
         Left            =   1980
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Reg74Ini 
         Height          =   300
         Left            =   810
         TabIndex        =   34
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownReg74Fim 
         Height          =   300
         Left            =   4410
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Reg74Fim 
         Height          =   300
         Left            =   3255
         TabIndex        =   36
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
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
         Left            =   2715
         TabIndex        =   38
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
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
         TabIndex        =   37
         Top             =   345
         Width           =   585
      End
   End
   Begin VB.ComboBox Finalidade 
      Height          =   315
      ItemData        =   "GeracaoArqICMSFISOcx.ctx":0000
      Left            =   1245
      List            =   "GeracaoArqICMSFISOcx.ctx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1305
      Width           =   3825
   End
   Begin VB.OptionButton TipoArq 
      Caption         =   "Interestadual"
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
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   1
      Top             =   855
      Width           =   1665
   End
   Begin VB.OptionButton TipoArq 
      Caption         =   "Integral"
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
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Value           =   -1  'True
      Width           =   1350
   End
   Begin VB.CommandButton BotaoArqCadastrados 
      Caption         =   "Arquivos Cadastrados"
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
      Left            =   2430
      TabIndex        =   16
      Top             =   780
      Width           =   2655
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   2970
      ScaleHeight     =   495
      ScaleWidth      =   2055
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   2115
      Begin VB.CommandButton BotaoSeguir 
         Height          =   360
         Left            =   60
         Picture         =   "GeracaoArqICMSFISOcx.ctx":0062
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Grava"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1050
         Picture         =   "GeracaoArqICMSFISOcx.ctx":01BC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   555
         Picture         =   "GeracaoArqICMSFISOcx.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1530
         Picture         =   "GeracaoArqICMSFISOcx.ctx":0878
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.TextBox NomeArquivo 
      Height          =   285
      Left            =   1050
      MaxLength       =   20
      TabIndex        =   15
      Top             =   5535
      Width           =   3090
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa"
      Height          =   1545
      Left            =   195
      TabIndex        =   26
      Top             =   2235
      Width           =   4905
      Begin MSMask.MaskEdBox TelContato 
         Height          =   315
         Left            =   1770
         TabIndex        =   11
         Top             =   1050
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   " "
      End
      Begin VB.TextBox NomeEmpresa 
         Height          =   315
         Left            =   1770
         MaxLength       =   35
         TabIndex        =   9
         Top             =   270
         Width           =   2895
      End
      Begin VB.TextBox Contato 
         Height          =   315
         Left            =   1770
         MaxLength       =   28
         TabIndex        =   10
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Empresa:"
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
         TabIndex        =   29
         Top             =   315
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Left            =   1005
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Telef. de Contato:"
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
         Left            =   180
         TabIndex        =   27
         Top             =   1095
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Endereço da Empresa"
      Height          =   1485
      Left            =   195
      TabIndex        =   22
      Top             =   3885
      Width           =   4905
      Begin VB.TextBox Complemento 
         Height          =   285
         Left            =   1800
         MaxLength       =   22
         TabIndex        =   14
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Endereco 
         Height          =   285
         Left            =   1800
         MaxLength       =   34
         TabIndex        =   12
         Top             =   300
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   690
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000012&
         Height          =   195
         Left            =   570
         TabIndex        =   25
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   1050
         TabIndex        =   24
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   885
         TabIndex        =   23
         Top             =   345
         Width           =   885
      End
   End
   Begin MSComCtl2.UpDown UpDownDataInicial 
      Height          =   300
      Left            =   1845
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1785
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   300
      Left            =   675
      TabIndex        =   4
      Top             =   1785
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataFinal 
      Height          =   300
      Left            =   4020
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1785
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   300
      Left            =   2865
      TabIndex        =   7
      Top             =   1785
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Finalidade:"
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
      Left            =   240
      TabIndex        =   31
      Top             =   1350
      Width           =   945
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   1845
      Width           =   360
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   1845
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
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
      Left            =   270
      TabIndex        =   30
      Top             =   5565
      Width           =   720
   End
End
Attribute VB_Name = "GeracaoArqICMSFISOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit

'??? "ERRO_LEITURA_CF_ARQICMS"

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoBotaoArq As AdmEvento
Attribute objEventoBotaoArq.VB_VarHelpID = -1

Private Sub BotaoArqCadastrados_Click()

Dim lErro As Long
Dim objInfoArqICMS As New ClassInfoArqICMS
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoArqCadastrados_Click
    
    'Chama ArquivoICMSLista
    Call Chama_Tela("ArquivoICMSLista", colSelecao, objInfoArqICMS, objEventoBotaoArq)
    
    Exit Sub

Erro_BotaoArqCadastrados_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160718)

    End Select

    Exit Sub


End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objInfoArqICMS As New ClassInfoArqICMS

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 69986
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 69987
    
    'Guarda dados da Apuração
    Call Move_Tela_Memoria(objInfoArqICMS)
            
    'Lê a Apuração ICMS a partir da FilialEmpresa, DataInicial e DataFinal
    lErro = CF("InfoArqICMS_Le", objInfoArqICMS)
    If lErro <> SUCESSO And lErro <> 69976 Then gError 69988

    'Se não encontrou, erro
    If lErro = 69976 Then gError 69989

    'Pede a confirmação da exclusão da apuração de ICMS
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_INFOARQICMS", objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If vbMsgRes = vbYes Then
    
        'Exclui a apuração de ICMS
        lErro = CF("InfoArqICMS_Exclui", objInfoArqICMS)
        If lErro <> SUCESSO Then gError 69990
    
        'Limpa a tela
        Call Limpa_Tela(Me)
    
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
    
        iAlterado = 0

    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 69986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)

        Case 69987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
        
        Case 69988, 69990
        
        Case 69989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INFOARQICMS_NAO_CADASTRADO", gErr, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160719)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    'Inicializa as datas
    Call Inicializa_Datas

    Finalidade.ListIndex = 0

    'Inicializa os campos da Filial Empresa
    lErro = Inicializa_Campos_FilialEmpresa()
    If lErro <> SUCESSO Then Error 69998
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case 69998
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160720)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoBotaoArq = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa as datas
    Call Inicializa_Datas
    
    Finalidade.ListIndex = 0

    Set objEventoBotaoArq = New AdmEvento
        
    'Inicializa os campos da Filial Empresa
    lErro = Inicializa_Campos_FilialEmpresa()
    If lErro <> SUCESSO Then Error 69997
    
    If InStr(UCase(gsNomeEmpresa), "INPAL") = 0 Then
        Saidas.Visible = False
        Entradas.Visible = False
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 69997
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160721)

    End Select

    Exit Sub

End Sub

Function Inicializa_Campos_FilialEmpresa() As Long

Dim lErro As Long
Dim objFilial As New AdmFiliais

On Error GoTo Erro_Inicializa_Campos_FilialEmpresa

    objFilial.iCodFilial = giFilialEmpresa

    'Le os dados da Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilial)
    If lErro <> SUCESSO Then gError 69770

    'default p/telefone e contato a partir de objFilial.objEndereco.sContato e objFilial.objEndereco.sTelefone1
    Contato.Text = objFilial.objEndereco.sContato
    NomeEmpresa.Text = gsNomeEmpresa
    NomeArquivo.Text = "ICMS.TXT"
    Endereco.Text = objFilial.objEndereco.sEndereco
    
    Exit Function
    
Erro_Inicializa_Campos_FilialEmpresa:

    Select Case gErr

        Case 69770 'Erro já Tratado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160722)

    End Select

    Exit Function

End Function

Private Sub BotaoSeguir_Click()

Dim lErro As String
Dim objInfoArqICMS As New ClassInfoArqICMS

On Error GoTo Erro_BotaoSeguir_Click

    'verificar se os campos obrigatorios estao preenchidos
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 69771
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 69772
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 69773
    If Len(Trim(NomeEmpresa.Text)) = 0 Then gError 69774
    If Len(Trim(Contato.Text)) = 0 Then gError 69775
    If Len(Trim(TelContato.Text)) = 0 Then gError 69776
    If Len(Trim(Endereco.Text)) = 0 Then gError 69777
    If Len(Trim(Numero.Text)) = 0 Then gError 69778
    
    objInfoArqICMS.iGeraReg54 = 1
    
    
    Call Move_Tela_Memoria(objInfoArqICMS)
    
    If objInfoArqICMS.dtDataInicial > objInfoArqICMS.dtDataFinal Then gError 69780
    
    If objInfoArqICMS.dtReg74DataInicial <> DATA_NULA And objInfoArqICMS.dtReg74DataFinal <> DATA_NULA Then
        If objInfoArqICMS.dtReg74DataInicial > objInfoArqICMS.dtReg74DataFinal Then gError 69780
        '??? verificar se existem inventarios para as datas informadas
    End If
    
    objInfoArqICMS.bIntegral = TipoArq(1).Value
    
    'Chama a Função de Gravação do Arquivo de Icms
    lErro = CF("Gerar_Arquivo_ICMS_FIS", objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69781
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ARQUIVO_GERADO", NomeArquivo.Text)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Unload Me

    Exit Sub

Erro_BotaoSeguir_Click:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 69771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)

        Case 69772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)

        Case 69773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
        
        Case 69774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", gErr)

        Case 69775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_PREENCHIDO", gErr)

        Case 69776
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELCONTATO_NAO_PREENCHIDO", gErr)

        Case 69777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_PREENCHIDO", gErr)

        Case 69778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

'        Case 69779
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_NAO_PREENCHIDO", gErr)

        Case 69780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 69781 'Erro já Tratado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160723)

    End Select

    Exit Sub

End Sub

Private Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFinal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtUltimoDiaMes As Date

On Error GoTo Erro_DataFinal_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataFinal.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 69782

        If Month(CDate(DataFinal.Text)) = 12 Then
            dtUltimoDiaMes = CDate("01/" & Month(CDate(DataFinal.Text) + 1) & "/" & (1 + Year(CDate(DataFinal.Text)))) - 1
        Else
            dtUltimoDiaMes = CDate("01/" & Month(CDate(DataFinal.Text) + 1) & "/" & Year(CDate(DataFinal.Text))) - 1
        End If
        If CDate(DataFinal.Text) <> dtUltimoDiaMes Then gError 69783

        If Len(Trim(DataInicial.ClipText)) > 0 Then
            If DataInicial.Text > DataFinal.Text Then gError 69784
        End If

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 69782 'Erro já Tratado

        Case 69783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_DO_MES", gErr, dtUltimoDiaMes)

        Case 69784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160724)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtPrimeiroDiaMes As Date

On Error GoTo Erro_DataInicial_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataInicial.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 69785

        dtPrimeiroDiaMes = CDate("01/" & Month(CDate(DataInicial.Text)) & "/" & Year(CDate(DataInicial.Text)))

        If CDate(DataInicial.Text) <> dtPrimeiroDiaMes Then gError 69786

        If Len(Trim(DataFinal.ClipText)) > 0 Then
            If DataInicial.Text > DataFinal.Text Then gError 69787
        End If

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 69785

        Case 69786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_DO_MES", gErr, dtPrimeiroDiaMes)

        Case 69787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160725)

    End Select

    Exit Sub

End Sub

Private Sub NomeArquivo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeEmpresa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero)

End Sub

Private Sub objEventoBotaoArq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objInfoArqICMS As ClassInfoArqICMS

On Error GoTo Erro_objEventoBotaoArq_evSelecao

    Set objInfoArqICMS = obj1
    
    Finalidade.ListIndex = 0
    
    Call DateParaMasked(DataInicial, objInfoArqICMS.dtDataInicial)
    Call DateParaMasked(DataFinal, objInfoArqICMS.dtDataFinal)
    
    NomeArquivo.Text = objInfoArqICMS.sNomeArquivo
    NomeEmpresa.Text = objInfoArqICMS.sNome
    Contato.Text = objInfoArqICMS.sContato
    Endereco.Text = objInfoArqICMS.sLogradouro
    Complemento.Text = objInfoArqICMS.sComplemento
    
    TelContato.PromptInclude = False
    TelContato.Text = objInfoArqICMS.sTelContato
    TelContato.PromptInclude = True
    
    Numero.PromptInclude = False
    Numero.Text = objInfoArqICMS.lNumero
    Numero.PromptInclude = True
            
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoArq_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160726)

    End Select

    Exit Sub
    
End Sub

Private Sub TelContato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TelContato_GotFocus()

    Call MaskEdBox_TrataGotFocus(TelContato)

End Sub

Private Sub UpDownDataFinal_DownClick()
'diminui a data final

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 69788

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 69788

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160727)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()
'aumenta a data final

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 69789

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 69789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160728)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_DownClick()
'diminui a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_DownClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 69790

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 69790

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160729)

    End Select

    Exit Sub


End Sub

Private Sub UpDownDataInicial_UpClick()
'aumenta a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_UpClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 69791

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 69791

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160730)

    End Select

    Exit Sub

End Sub

Private Sub Inicializa_Datas()
'inicializa as datas inicial e final e coloca nos respectivos campos

    Dim dtDataInicial As Date
    Dim dtDataFinal As Date
    Dim iMesAtual As Integer
    Dim iAnoAtual As Integer

    'coloca o mes e o ano correntes nas variaveis iMes e iAno
    iMesAtual = Month(gdtDataAtual)
    iAnoAtual = Year(gdtDataAtual)

    'obter data inicial
    If iMesAtual < 4 Then

        dtDataInicial = CDate("01/" & CStr(iMesAtual + 9) & "/" & CStr(iAnoAtual - 1))

    Else

        dtDataInicial = CDate("01/" & CStr(iMesAtual - 3) & "/" & CStr(iAnoAtual))

    End If

    'obter data final
    dtDataFinal = CDate("01/" & CStr(iMesAtual) & "/" & CStr(iAnoAtual)) - 1

    Call DateParaMasked(DataInicial, dtDataInicial)
    Call DateParaMasked(DataFinal, dtDataFinal)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Move_Tela_Memoria(objInfoArqICMS As ClassInfoArqICMS)
    
Dim iPos As Integer
Dim sTel As String

    objInfoArqICMS.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(DataFinal.ClipText)) > 0 Then objInfoArqICMS.dtDataFinal = CDate(DataFinal.Text)
    If Len(Trim(DataInicial.ClipText)) > 0 Then objInfoArqICMS.dtDataInicial = CDate(DataInicial.Text)
    If Len(Trim(Numero.Text)) > 0 Then objInfoArqICMS.lNumero = CLng(Numero.Text)
    objInfoArqICMS.sComplemento = Complemento.Text
    objInfoArqICMS.sContato = Contato.Text
    objInfoArqICMS.sLogradouro = Endereco.Text
    objInfoArqICMS.sNomeArquivo = NomeArquivo.Text
    iPos = InStr(1, TelContato.ClipText, "-")
    If iPos <> 0 Then
        sTel = left(Trim(TelContato.Text), iPos - 1) & right(Trim(TelContato.Text), Len(Trim(TelContato.Text)) - iPos)
    Else
        sTel = Trim(TelContato.Text)
    End If
    objInfoArqICMS.sTelContato = sTel
    objInfoArqICMS.sNome = NomeEmpresa.Text
    objInfoArqICMS.dtReg74DataInicial = MaskedParaDate(Reg74Ini)
    objInfoArqICMS.dtReg74DataFinal = MaskedParaDate(Reg74Fim)
    objInfoArqICMS.iGeraSaidas = Saidas.Value
    objInfoArqICMS.iGeraEntradas = Entradas.Value
    objInfoArqICMS.sFinalidade = Finalidade.Text
    
End Sub

Sub Traz_InfoArqICMS_Tela(objInfoArqICMS As ClassInfoArqICMS)

    If objInfoArqICMS.bIntegral <> 0 Then
        objInfoArqICMS.bIntegral = True
        TipoArq(1).Value = True
        TipoArq(2).Value = False
    Else
        TipoArq(1).Value = False
        TipoArq(2).Value = True
    End If
    
    Call DateParaMasked(DataFinal, objInfoArqICMS.dtDataFinal)
    Call DateParaMasked(DataInicial, objInfoArqICMS.dtDataInicial)
        
    Numero.PromptInclude = False 'Incluido por Daniel
    Numero.Text = CStr(objInfoArqICMS.lNumero) 'Alterado por Leo em 11/12/01
    Numero.PromptInclude = True 'Incluido por Daniel
    
    Complemento.Text = objInfoArqICMS.sComplemento
    Contato.Text = objInfoArqICMS.sContato
    Endereco.Text = objInfoArqICMS.sLogradouro
    NomeArquivo.Text = objInfoArqICMS.sNomeArquivo
    TelContato.PromptInclude = False 'incluido por Leo em 11/12/01
    TelContato.Text = objInfoArqICMS.sTelContato
    TelContato.PromptInclude = True 'incluido por Leo em 11/12/01
    NomeEmpresa.Text = objInfoArqICMS.sNome

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objInfoArqICMS As New ClassInfoArqICMS

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "InfoArqICMS"

    'Move os dados da tela para memória
    Call Move_Tela_Memoria(objInfoArqICMS)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objInfoArqICMS.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "DataInicial", objInfoArqICMS.dtDataInicial, 0, "DataInicial"
    colCampoValor.Add "DataFinal", objInfoArqICMS.dtDataFinal, 0, "DataFinal"
    colCampoValor.Add "Nome", objInfoArqICMS.sNome, STRING_FILIALEMPRESA_NOME, "Nome"
    colCampoValor.Add "Logradouro", objInfoArqICMS.sLogradouro, STRING_LOGRADOURO, "Logradouro"
    colCampoValor.Add "Numero", objInfoArqICMS.lNumero, 0, "Numero"
    colCampoValor.Add "Complemento", objInfoArqICMS.sComplemento, STRING_COMPLEMENTO, "Complemento"
    colCampoValor.Add "Contato", objInfoArqICMS.sContato, STRING_CONTATO_REGAPURACAO, "Contato"
    colCampoValor.Add "TelContato", objInfoArqICMS.sTelContato, STRING_TELCONTATO, "TelContato"
    colCampoValor.Add "NomeArquivo", objInfoArqICMS.sNomeArquivo, STRING_NOME_ARQ_COMPLETO, "NomeArquivo"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160731)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objInfoArqICMS As New ClassInfoArqICMS
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objInfoArqICMS com os dados passados em colCampoValor
    objInfoArqICMS.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objInfoArqICMS.dtDataInicial = colCampoValor.Item("DataInicial").vValor
    objInfoArqICMS.dtDataFinal = colCampoValor.Item("DataFinal").vValor
    objInfoArqICMS.sNome = colCampoValor.Item("Nome").vValor
    objInfoArqICMS.sLogradouro = colCampoValor.Item("Logradouro").vValor
    objInfoArqICMS.lNumero = colCampoValor.Item("Numero").vValor
    objInfoArqICMS.sComplemento = colCampoValor.Item("Complemento").vValor
    objInfoArqICMS.sContato = colCampoValor.Item("Contato").vValor
    objInfoArqICMS.sTelContato = colCampoValor.Item("TelContato").vValor
    objInfoArqICMS.sNomeArquivo = colCampoValor.Item("NomeArquivo").vValor

    'Se o NumIntDoc estiver preenchido
    If objInfoArqICMS.lNumIntDoc <> 0 Then

        'Traz os dados do Arquivo de ICMS para a Tela
        Call Traz_InfoArqICMS_Tela(objInfoArqICMS)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160732)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_ARQICMS
    Set Form_Load_Ocx = Me
    Caption = "Geração de Arquivo para ICMS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "GeracaoArqICMSFIS"

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

Function Trata_Parametros(Optional obj1 As Object) As Long
'Parâmetro incluido por Leo em 11/12/01

Dim lErro As Long
Dim objInfoArqICMS As ClassInfoArqICMS

On Error GoTo Erro_Trata_Parametros
    
    '**** trecho de código incluido por Leo em 11/12/01 ****
    If Not (obj1 Is Nothing) Then
    
        Set objInfoArqICMS = obj1
                
        Call Traz_InfoArqICMS_Tela(objInfoArqICMS)
            
    End If
    '**** Leo até aqui ****
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160733)

    End Select

    Exit Function

End Function


Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

