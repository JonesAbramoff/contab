VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl Despachante 
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   9255
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   2100
      Picture         =   "DespachanteGR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   195
      Width           =   300
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contatos"
      Height          =   1560
      Left            =   225
      TabIndex        =   27
      Top             =   3525
      Width           =   8940
      Begin MSMask.MaskEdBox TextTelefone 
         Height          =   240
         Left            =   5385
         TabIndex        =   32
         Top             =   690
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TextFax 
         Height          =   240
         Left            =   4050
         TabIndex        =   31
         Top             =   780
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin VB.TextBox TextSetor 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   12
         Top             =   435
         Width           =   1605
      End
      Begin VB.TextBox TextContato 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   555
         MaxLength       =   50
         TabIndex        =   11
         Top             =   540
         Width           =   2025
      End
      Begin VB.TextBox TextEmail 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   13
         Top             =   630
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid GridContatos 
         Height          =   1215
         Left            =   150
         TabIndex        =   28
         Top             =   240
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   2143
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7035
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DespachanteGR.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "DespachanteGR.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DespachanteGR.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "DespachanteGR.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboEstado 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1545
      TabIndex        =   9
      Top             =   3090
      Width           =   630
   End
   Begin VB.ComboBox ComboPais 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4305
      TabIndex        =   10
      Top             =   3075
      Width           =   1995
   End
   Begin VB.TextBox TextEndereco 
      Height          =   315
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2040
      Width           =   7560
   End
   Begin MSMask.MaskEdBox MaskCidade 
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Top             =   2550
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskBairro 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   2550
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   12
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskCEP 
      Height          =   315
      Left            =   6960
      TabIndex        =   8
      Top             =   2550
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#####-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskNome 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CGC 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   645
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   14
      Mask            =   "##############"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   180
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelNomeReduzido 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   30
      Top             =   1605
      Width           =   1410
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   825
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   210
      Width           =   660
   End
   Begin VB.Label LabelCGC 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ/CPF:"
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
      TabIndex        =   26
      Top             =   720
      Width           =   975
   End
   Begin VB.Label LabelNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   945
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   24
      Top             =   1155
      Width           =   540
   End
   Begin VB.Label Label1 
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
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   23
      Top             =   2100
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cidade:"
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
      Index           =   4
      Left            =   3600
      TabIndex        =   22
      Top             =   2625
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
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
      Index           =   6
      Left            =   810
      TabIndex        =   21
      Top             =   3135
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bairro:"
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
      Index           =   3
      Left            =   900
      TabIndex        =   20
      Top             =   2595
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CEP:"
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
      Index           =   5
      Left            =   6435
      TabIndex        =   19
      Top             =   2625
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "País:"
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
      Index           =   11
      Left            =   3780
      TabIndex        =   18
      Top             =   3150
      Width           =   495
   End
End
Attribute VB_Name = "Despachante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Definicoes do Grid de Contatos
Dim objGridContatos As New AdmGrid

Dim iGrid_Contato_Col As Integer
Dim iGrid_Setor_Col As Integer
Dim iGrid_Telefone_Col As Integer
Dim iGrid_Fax_Col As Integer
Dim iGrid_Email_Col As Integer

Public iAlterado As Integer

Private WithEvents objEventoDespachante As AdmEvento
Attribute objEventoDespachante.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub LabelCodigo_Click()

Dim colDespachante As Collection
Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_LabelCodigo_Click

    'Carrega todos os dados da minha tela para o objDespachante
    Call Move_Tela_Memoria(objDespachante)

    'Chama o browser Despachante
    Call Chama_Tela("DespachanteLista", colDespachante, objDespachante, objEventoDespachante)
    
    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelCGC_Click()

Dim colDespachante As Collection
Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_LabelCGC_Click

    'Carrega todos os dados da minha tela para o objDespachante
    Call Move_Tela_Memoria(objDespachante)

    'Chama o browser Despachante
    Call Chama_Tela("DespachanteLista", colDespachante, objDespachante, objEventoDespachante)
    
    Exit Sub

Erro_LabelCGC_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelNome_Click()

Dim colDespachante As Collection
Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_LabelNome_Click

    'Carrega todos os dados da minha tela para o objDespachante
    Call Move_Tela_Memoria(objDespachante)

    'Chama o browser Despachante
    Call Chama_Tela("DespachanteLista", colDespachante, objDespachante, objEventoDespachante)
    
    Exit Sub

Erro_LabelNome_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeReduzido_Click()

Dim colDespachante As Collection
Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_LabelNomeReduzido_Click

    'Carrega todos os dados da minha tela para o objDespachante
    Call Move_Tela_Memoria(objDespachante)

    'Chama o browser Despachante
    Call Chama_Tela("DespachanteLista", colDespachante, objDespachante, objEventoDespachante)
    
    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoDespachante_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDespachante As ClassDespachante

On Error GoTo Erro_objEventoDespachante_evSelecao

    Set objDespachante = obj1

    'Move os dados para a tela
    lErro = Traz_Despachante_Tela(objDespachante)
    If lErro <> SUCESSO And lErro <> 96668 Then gError 96686

    'Se não existe o Código passado --> Erro.
    If lErro = 96668 Then gError 96687

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

Erro_objEventoDespachante_evSelecao:

    Select Case gErr

        Case 96686
        
        Case 96687
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objDespachante.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objDespachante As ClassDespachante) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Código selecionado, exibir seus dados
    If Not (objDespachante Is Nothing) Then
    
        'Verifica se o Código existe
        lErro = Traz_Despachante_Tela(objDespachante)
        If lErro <> SUCESSO And lErro <> 96668 Then gError 96666
        
        'Se não existe o Código passado
        If lErro = 96668 Then

            'Limpa a Tela
            Call Limpa_Despachante

            'Se Código não está cadastrado
            If objDespachante.iCodigo <> 0 Then MaskCodigo.Text = CStr(objDespachante.iCodigo)
            If Len(Trim(objDespachante.sNomeReduzido)) <> 0 Then NomeReduzido.Text = objDespachante.sNomeReduzido
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 96666

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0
  
    Exit Function

End Function

Sub Limpa_Despachante()

Dim lErro As Long
    
    'Limpa Tela
    Call Limpa_Tela(Me)
    
    'Limpa o grid
    Call Grid_Limpa(objGridContatos)
    
    'Limpa os outros campos
    ComboEstado.Text = ""
    ComboPais.ListIndex = 0
    
End Sub

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicialização do objEventoDespachante
    Set objEventoDespachante = New AdmEvento
    
    'Executa inicializacao do GridContatos
    lErro = Inicializa_Grid_Contatos(objGridContatos)
    If lErro <> SUCESSO Then gError 96663
        
    lErro = Inicializa_Endereco()
    If lErro <> SUCESSO Then gError 96664
        
    ComboPais.ListIndex = 0
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
       
        Case 96663, 99664
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoDespachante = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Function Traz_Despachante_Tela(objDespachante As ClassDespachante) As Long

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objContato As New ClassContato
    
On Error GoTo Erro_Traz_Despachante_Tela
    
    'Limpa a tela
    Call Limpa_Despachante
    
    'Le os dados do despachante com o código passado
    lErro = CF("Despachante_Le", objDespachante)
    If lErro <> SUCESSO And lErro <> 96679 Then gError 96667
    
    'Se não existe Despachante com o Código passado --> Erro
    If lErro = 96679 Then gError 96668
    
    'Se existe um endereço
    If objDespachante.lEndereco > 0 Then
    
        objDespachante.objEndereco.lCodigo = objDespachante.lEndereco

        'Lê o endereço
        lErro = CF("Endereco_Le", objDespachante.objEndereco)
        If lErro <> SUCESSO Then gError 96669
                      
    End If
    
    'Lê contatos do despachante passado
    lErro = CF("DespachanteContatos_Le", objDespachante)
    If lErro <> SUCESSO Then gError 96670
    
    'preenche a tela com os dados recebidos pelo obj
    MaskCodigo.Text = objDespachante.iCodigo
    MaskNome.Text = objDespachante.sNome
    NomeReduzido.Text = objDespachante.sNomeReduzido
    
    'Se CGC está preenchido
    If Len(Trim(objDespachante.sCgc)) > 0 Then
        CGC.Text = objDespachante.sCgc
        Call CGC_Validate(False)
    End If
    
    ComboEstado.Text = objDespachante.objEndereco.sSiglaEstado
    TextEndereco.Text = objDespachante.objEndereco.sEndereco
    MaskBairro.Text = objDespachante.objEndereco.sBairro
    MaskCidade.Text = objDespachante.objEndereco.sCidade
    MaskCEP.Text = objDespachante.objEndereco.sCEP
    
    'Se existe um país
    If objDespachante.objEndereco.iCodigoPais > 0 Then
        ComboPais.Text = objDespachante.objEndereco.iCodigoPais
        Call ComboPais_Validate(False)
    Else
        ComboPais.Text = ""
    End If
               
    'Carrega a tabela de contatos com os dados do despachante
    Call Carrega_GridContatos(objDespachante)
    
    iAlterado = 0
    
    Exit Function

Erro_Traz_Despachante_Tela:

    Traz_Despachante_Tela = gErr

    Select Case gErr
            
        Case 96667, 96668, 96670
                          
        Case 96669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'Coloca o próximo número a ser gerado na tela

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_DESPACHANTE", "Despachante", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 96684

    'Joga o código na tela
    MaskCodigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 96684

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_GridContatos(objDespachante As ClassDespachante)
'Carrega a tabela de contatos com os dados do despachante

Dim iLinha As Integer
Dim objContato  As ClassContato

    'Limpa o Grid de Contato
    Call Grid_Limpa(objGridContatos)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de contato
    'Para cada Contato encontrado
    For Each objContato In objDespachante.colContato

       iLinha = iLinha + 1

       GridContatos.TextMatrix(iLinha, iGrid_Contato_Col) = objContato.sContato
       GridContatos.TextMatrix(iLinha, iGrid_Setor_Col) = objContato.sSetor
       GridContatos.TextMatrix(iLinha, iGrid_Email_Col) = objContato.sEmail
       GridContatos.TextMatrix(iLinha, iGrid_Telefone_Col) = objContato.sTelefone
       GridContatos.TextMatrix(iLinha, iGrid_Fax_Col) = objContato.sFax

    Next
    'Guarda o número de linhas existentes
    objGridContatos.iLinhasExistentes = iLinha

End Sub

Function Move_Tela_Memoria(objDespachante As ClassDespachante) As Long
'Move os dados da tela para o objDespachante

    objDespachante.iCodigo = StrParaInt(MaskCodigo.Text)
    objDespachante.sCgc = CGC.Text
    objDespachante.sNome = MaskNome.Text
    objDespachante.sNomeReduzido = NomeReduzido.Text
    
    'Move os dados do endereço
    objDespachante.objEndereco.sEndereco = TextEndereco.Text
    objDespachante.objEndereco.sBairro = MaskBairro.Text
    objDespachante.objEndereco.sCidade = MaskCidade.Text
    objDespachante.objEndereco.sCEP = MaskCEP.Text
    objDespachante.objEndereco.sSiglaEstado = ComboEstado.Text
    objDespachante.objEndereco.iCodigoPais = Codigo_Extrai(ComboPais.Text)
    
    'Move os dados do grid
    Call Move_GridContato_Memoria(objDespachante)

End Function

Function Move_GridContato_Memoria(objDespachante As ClassDespachante)
'Move itens do Grid para objDespachante

Dim objContato As ClassContato
Dim iLinha As Integer

    'Para cada linha do Grid
    For iLinha = 1 To objGridContatos.iLinhasExistentes
        
        'Inicializa o objContato
        Set objContato = New ClassContato
        
        'Carrega os dados em objContato
        objContato.sContato = GridContatos.TextMatrix(iLinha, iGrid_Contato_Col)
        objContato.sEmail = GridContatos.TextMatrix(iLinha, iGrid_Email_Col)
        objContato.sFax = GridContatos.TextMatrix(iLinha, iGrid_Fax_Col)
        objContato.sSetor = GridContatos.TextMatrix(iLinha, iGrid_Setor_Col)
        objContato.sTelefone = GridContatos.TextMatrix(iLinha, iGrid_Telefone_Col)
        
        'Preenche a coleção
        objDespachante.colContato.Add objContato

    Next

End Function

Private Function Inicializa_Endereco() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Inicializa_Endereco
    
    'Para cada item encontrado
    For iIndice = 1 To gcolUFs.Count
        
        'Adiciona na Combo Estado
        ComboEstado.AddItem gcolUFs.Item(iIndice)
        
    Next

    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 96665

    'Adiciona na Combo País
    For Each objCodigoDescricao In colCodigoDescricao
        ComboPais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ComboPais.ItemData(ComboPais.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Inicializa_Endereco = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Endereco:

    Inicializa_Endereco = gErr
    
    Select Case gErr
        
        Case 96665
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
        
    End Select
        
    Exit Function
    
End Function

Private Function Inicializa_Grid_Contatos(objGridInt As AdmGrid) As Long
'Inicializa o grid de Contatos da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contatos

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Contato")
    objGridInt.colColuna.Add ("Setor")
    objGridInt.colColuna.Add ("Telefone")
    objGridInt.colColuna.Add ("Fax")
    objGridInt.colColuna.Add ("E-Mail")

    'campos de edição do grid
    objGridInt.colCampo.Add (TextContato.Name)
    objGridInt.colCampo.Add (TextSetor.Name)
    objGridInt.colCampo.Add (TextTelefone.Name)
    objGridInt.colCampo.Add (TextFax.Name)
    objGridInt.colCampo.Add (TextEmail.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Contato_Col = 1
    iGrid_Setor_Col = 2
    iGrid_Telefone_Col = 3
    iGrid_Fax_Col = 4
    iGrid_Email_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridContatos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CONTATOS + 1

    'Largura da primeira coluna
    GridContatos.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contatos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contatos:

    Inicializa_Grid_Contatos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 96697

    'Limpa a Tela
    Call Limpa_Despachante

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 96697

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Controla toda a rotina de gravação

Dim lErro As Long
Dim objDespachante As New ClassDespachante

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram informados
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96698
    If Len(Trim(CGC.ClipText)) = 0 Then gError 96699
    If Len(Trim(MaskNome.Text)) = 0 Then gError 96700
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 96701
    
    'Move os campos da tela para o objDespachante
    Call Move_Tela_Memoria(objDespachante)

    'Verifica se o Código já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objDespachante, objDespachante.iCodigo)
    If lErro <> SUCESSO Then gError 96702

    'Grava o Código no banco de dados
    lErro = CF("Despachante_Grava", objDespachante)
    If lErro <> SUCESSO Then gError 96703

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_CADASTRADO", gErr)

        Case 96700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 96701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_NAO_PREENCHIDO", gErr)

        Case 96702, 96703

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 96704

    'Limpa a Tela
    Call Limpa_Despachante

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 96704

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 96705

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 96705

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui o Despachante do código passado

Dim lErro As Long
Dim objDespachante As New ClassDespachante
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código foi informado, senão --> Erro.
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96706

    objDespachante.iCodigo = CInt(MaskCodigo.Text)

    'Verifica se o Despachante existe
    lErro = CF("Despachante_Le", objDespachante)
    If lErro <> SUCESSO And lErro <> 96679 Then gError 96707

    'Se Despachante não está cadastrado --> Erro
    If lErro = 96679 Then gError 96708

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DESPACHANTE", objDespachante.iCodigo)
    
    'Se confirma
    If vbMsgRet = vbYes Then

        'exclui o Despachante
        lErro = CF("Despachante_Exclui", objDespachante)
        If lErro <> SUCESSO Then gError 96709

        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)

        'Limpa a Tela
        Call Limpa_Despachante

        iAlterado = 0

    End If

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96707, 96709

        Case 96708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objDespachante.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Function Contato_Codigo_Automatico(lCodigo As Long) As Long
'Retorna o proximo número disponivel

Dim lErro As Long

On Error GoTo Erro_Contato_Codigo_Automatico

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico_EmTrans", "CRFatConfig", "NUM_PROX_CONTATO", "Contato", "NumIntDoc", lCodigo)
    If lErro <> SUCESSO Then gError 96733

    Contato_Codigo_Automatico = SUCESSO

    Exit Function

Erro_Contato_Codigo_Automatico:

    Contato_Codigo_Automatico = gErr

    Select Case gErr

        Case 96733

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 96688
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 96689
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            gError 96690

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case gErr

        Case 96688, 96689

        Case 96690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select


    Exit Sub

End Sub

Private Sub ComboEstado_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboEstado_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ComboEstado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ComboEstado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(ComboEstado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
    If ComboEstado.Text = ComboEstado.List(ComboEstado.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(ComboEstado)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 96691

    'Não existe o ítem na ComboBox Estado
    If lErro = 58583 Then gError 96692

    Exit Sub

Erro_ComboEstado_Validate:

    Cancel = True

    Select Case gErr

        Case 96691

        Case 96692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, ComboEstado.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub

Private Sub ComboPais_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ComboPais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboPais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_ComboPais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(ComboPais.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Pais
    If ComboPais.Text = ComboPais.List(ComboPais.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboPais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 96693

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then

        objPais.iCodigo = iCodigo

        'Tenta ler Pais com esse codigo no BD
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 96694
        
        'Se não achou
        If lErro <> SUCESSO Then
            
            'pergunta se deseja cadastrar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
            
            'Se confirma
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Paises", objPais)
            End If

        End If
        
        'Joga o país na tela
        ComboPais.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome

    End If

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 96696

    Exit Sub

Erro_ComboPais_Validate:

    Cancel = True

    Select Case gErr

        Case 96693, 96694, 96695

        Case 96696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(ComboPais.Text))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub GridContatos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_GotFocus()
    Call Grid_Recebe_Foco(objGridContatos)
End Sub

Public Sub GridContatos_EnterCell()
    Call Grid_Entrada_Celula(objGridContatos, iAlterado)
End Sub

Public Sub GridContatos_LeaveCell()
    Call Saida_Celula(objGridContatos)
End Sub

Public Sub GridContatos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContatos)
End Sub

Public Sub GridContatos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContatos)
End Sub

Public Sub GridContatos_RowColChange()
    
    Call Grid_RowColChange(objGridContatos)
    
End Sub

Public Sub GridContatos_Scroll()
    Call Grid_Scroll(objGridContatos)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Contato
        Case iGrid_Contato_Col
            lErro = Saida_Celula_TextContato(objGridInt)
            If lErro <> SUCESSO Then gError 96755
            
        'Setor
        Case iGrid_Setor_Col
            lErro = Saida_Celula_TextSetor(objGridInt)
            If lErro <> SUCESSO Then gError 96756

        'Telefone
        Case iGrid_Telefone_Col
            lErro = Saida_Celula_TextTelefone(objGridInt)
            If lErro <> SUCESSO Then gError 96757

        'Fax
        Case iGrid_Fax_Col
            lErro = Saida_Celula_TextFax(objGridInt)
            If lErro <> SUCESSO Then gError 96758

        'Email
        Case iGrid_Email_Col
            lErro = Saida_Celula_TextEmail(objGridInt)
            If lErro <> SUCESSO Then gError 96759

    End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 96760

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 96755, 96756, 96757, 96758, 96759, 96760

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub TextContato_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub textContato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textContato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textContato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextContato
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TextTelefone_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub textTelefone_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textTelefone_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textTelefone_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextTelefone
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TextFax_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub textFax_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textFax_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textFax_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextFax
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TextEmail_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub textEmail_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textEmail_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textEmail_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextEmail
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TextSetor_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub textSetor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridContatos)

End Sub

Private Sub textSetor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)

End Sub

Private Sub textSetor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = TextSetor
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_TextContato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TextContato que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextContato

    Set objGridInt.objControle = TextContato
    
    If GridContatos.Row - GridContatos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96671

    Saida_Celula_TextContato = SUCESSO

    Exit Function

Erro_Saida_Celula_TextContato:

    Saida_Celula_TextContato = gErr

    Select Case gErr

        Case 96671
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextSetor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TextSetor que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextSetor

    Set objGridInt.objControle = TextSetor

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96672

    Saida_Celula_TextSetor = SUCESSO

    Exit Function

Erro_Saida_Celula_TextSetor:

    Saida_Celula_TextSetor = gErr

    Select Case gErr

        Case 96672
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextTelefone(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TextTelefone que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextTelefone

    Set objGridInt.objControle = TextTelefone

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96673

    Saida_Celula_TextTelefone = SUCESSO

    Exit Function

Erro_Saida_Celula_TextTelefone:

    Saida_Celula_TextTelefone = gErr

    Select Case gErr

        Case 96673
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextFax(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TextFax que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextFax

    Set objGridInt.objControle = TextFax

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96674

    Saida_Celula_TextFax = SUCESSO

    Exit Function

Erro_Saida_Celula_TextFax:

    Saida_Celula_TextFax = gErr

    Select Case gErr

        Case 96674
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TextEmail(objGridInt As AdmGrid) As Long
'Faz a crítica da célula TextEmail que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TextEmail

    Set objGridInt.objControle = TextEmail

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96675

    Saida_Celula_TextEmail = SUCESSO

    Exit Function

Erro_Saida_Celula_TextEmail:

    Saida_Celula_TextEmail = gErr

    Select Case gErr

        Case 96675
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub MaskBairro_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCEP_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub MaskCEP_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MaskCEP, iAlterado)

End Sub

Private Sub MaskCidade_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCodigo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)
'Verifica se o código é válido

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate
    
    'Verifica se código foi informado
    If Len(MaskCodigo.Text) > 0 Then
    
        'Verifica se o código é um valor positivo
        lErro = Valor_Positivo_Critica(MaskCodigo.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 98386
    End If

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 98386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub MaskNome_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NomeReduzido_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objDespachante.iCodigo = colCampoValor.Item("Codigo").vValor
    
    'Se o Código está sendo passado
    If objDespachante.iCodigo <> 0 Then

        'Traz dados para a Tela
        lErro = Traz_Despachante_Tela(objDespachante)
        If lErro <> SUCESSO And lErro <> 96668 Then gError 96684
        
        'Se não encontrar --> erro.
        If lErro = 96668 Then gError 96685

        iAlterado = 0

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 96684
        
        Case 96685
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objDespachante.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objDespachante As New ClassDespachante
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Despachante"

    'Le os dados da Tela Despachante
    Call Move_Tela_Memoria(objDespachante)
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objDespachante.iCodigo, 0, "Codigo"
    colCampoValor.Add "CGC", objDespachante.sCgc, STRING_DESPACHANTE_CGC, "CGC"
    colCampoValor.Add "Nome", objDespachante.sNome, STRING_DESPACHANTE_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objDespachante.sNomeReduzido, STRING_DESPACHANTE_NOMEREDUZIDO, "NomeReduzido"
    colCampoValor.Add "Endereco", objDespachante.lEndereco, 0, "Endereco"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Despachante"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "Despachante"

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
'''    m_Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        'Se F2
        Case KEYCODE_PROXIMO_NUMERO
            'Próximo número do despachante
            Call BotaoProxNum_Click
        
        'Se F3
        Case KEYCODE_BROWSER
            
            'Caso Seja Browser --> chama o correnspondente ao ativo no momento
            If Me.ActiveControl Is CGC Then Call LabelCGC_Click
            If Me.ActiveControl Is MaskNome Then Call LabelNome_Click
            If Me.ActiveControl Is NomeReduzido Then Call LabelNomeReduzido_Click
            If Me.ActiveControl Is MaskCodigo Then Call LabelCodigo_Click
        
    End Select

End Sub

'***** fim do trecho a ser copiado ******

