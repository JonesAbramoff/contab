VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoOutros 
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   KeyPreview      =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   4725
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   3060
      Picture         =   "BorderoOutros.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   810
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   1935
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   75
      Width           =   2700
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   615
         Picture         =   "BorderoOutros.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "BorderoOutros.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2085
         Picture         =   "BorderoOutros.ctx":0776
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   120
         Picture         =   "BorderoOutros.ctx":08F4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1095
         Picture         =   "BorderoOutros.ctx":09F6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.ComboBox AdmMeioPagto 
      Height          =   315
      Left            =   1995
      TabIndex        =   4
      Top             =   1755
      Width           =   2625
   End
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      Left            =   1980
      TabIndex        =   5
      ToolTipText     =   "Formas de Parcelamento"
      Top             =   2235
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhado"
      Height          =   1245
      Left            =   105
      TabIndex        =   16
      Top             =   2685
      Width           =   4515
      Begin MSMask.MaskEdBox ValorEnviar 
         Height          =   300
         Left            =   2070
         TabIndex        =   6
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   3
         Left            =   1455
         TabIndex        =   19
         Top             =   840
         Width           =   510
      End
      Begin VB.Label LabelTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2100
         TabIndex        =   18
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Caixa Central:"
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
         Left            =   480
         TabIndex        =   17
         Top             =   375
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Não Detalhado"
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   4110
      Width           =   4515
      Begin MSMask.MaskEdBox ValorEnviarN 
         Height          =   300
         Left            =   2070
         TabIndex        =   7
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Caixa Central:"
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
         Left            =   480
         TabIndex        =   15
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label LabelTotalN 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2085
         TabIndex        =   14
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   6
         Left            =   1470
         TabIndex        =   13
         Top             =   840
         Width           =   510
      End
   End
   Begin MSMask.MaskEdBox DataEnvio 
      Height          =   300
      Left            =   1995
      TabIndex        =   3
      Top             =   1275
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataEnvio 
      Height          =   300
      Left            =   2955
      TabIndex        =   21
      Top             =   1275
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1995
      TabIndex        =   1
      Top             =   795
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data de Envio:"
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
      Index           =   2
      Left            =   615
      TabIndex        =   25
      Top             =   1335
      Width           =   1290
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
      Left            =   1245
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   24
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Meio de Pagamento:"
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
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   1815
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Index           =   9
      Left            =   675
      TabIndex        =   22
      ToolTipText     =   "Formas de Parcelamento"
      Top             =   2280
      Width           =   1230
   End
End
Attribute VB_Name = "BorderoOutros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iAlterado As Integer

Private WithEvents objEventoBorderoOutros As AdmEvento
Attribute objEventoBorderoOutros.VB_VarHelpID = -1

Dim giAdmMeioPagtoVelho As Integer
Dim giParcelamentoVelho As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objBordero As New ClassBorderoOutros

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o código estiver vazio-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 107420
    
    'se a data estiver em branco-> erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 107421
    
    'se a admmeiopagto estiver em branco-> erro
    If AdmMeioPagto.ListIndex = -1 Then gError 107422
    
    Call Move_Tela_Memoria(objBordero)
    'If lErro <> SUCESSO Then gError 120045

    lErro = CF("BorderoOutros_Le", objBordero)
    If lErro <> SUCESSO And lErro <> 108059 Then gError 120041
    
    If lErro = 108059 Then gError 120043
    
    '???? adaptar para bordero outros
    'ver expr. selecao, nome tsk, etc..
    'aguardando tsk ficar pronto....
    'lErro = objRelatorio.ExecutarDireto("Borderô Outros", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 120042

    'Limpa a Tela
    Call Limpa_Tela_BorderoOutros

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 120037
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 120038
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 120039
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_SELECIONADO", gErr)
        
        Case 120040
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO1", gErr)

        Case 120041, 120042, 120044

        Case 120043
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROOUTROS_NAOENCONTRADO", gErr, objBordero.iFilialEmpresa, objBordero.lNumBordero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 143751)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelCodigo_Click
    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Form_Load

    'carrega a combo de admmeiopagto
    lErro = Carrega_Outros()
    If lErro <> SUCESSO Then gError 108053

    'preenche a data com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    'preenche um tmplojafilial para ler o seu saldo
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_OUTROS
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'le o seu saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108054

    'preenche o total não especificado
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")

    'preenche o total especificado
    LabelTotal.Caption = Format(0, "STANDARD")

    'instancia o objeto com eventos
    Set objEventoBorderoOutros = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 108053, 108054
        
        Case 108055
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJAFILIAL_NAOENCONTRADO", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143752)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoOutros As ClassBorderoOutros) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objBorderoOutros Is Nothing Then
    
        'se o número do borderô estiver preenchido
        If objBorderoOutros.lNumBordero <> 0 Then
            
            'lê o bordero
            lErro = CF("BorderoOutros_Le", objBorderoOutros)
            If lErro <> SUCESSO And lErro <> 108059 Then gError 108060
            
            'se não encontrou
            If lErro <> 108059 Then
        
                'busca o BorderoOutros
                lErro = Traz_BorderoOutros_Tela(objBorderoOutros)
                If lErro <> SUCESSO Then gError 108061
            
            End If
        
        Else
            
            'limpa a tela
            Call Limpa_Tela_BorderoOutros
            
            'preenche o codigo com o codigo buscado
            Codigo.Text = objBorderoOutros.lNumBordero
        
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr

        Case 108060, 108061

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143753)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objBorderoOutros As New ClassBorderoOutros
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodigo_Click

    'se o código estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then
    
        'preenche um borderooutros com os dados necessários para chamar o browser
        objBorderoOutros.iFilialEmpresa = giFilialEmpresa
        objBorderoOutros.lNumBordero = StrParaLong(Codigo.Text)
    
    End If
    
    Call Chama_Tela("BorderoOutrosLista", colSelecao, objBorderoOutros, objEventoBorderoOutros)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143754)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoBorderoOutros_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objBorderoOutros As ClassBorderoOutros

On Error GoTo Erro_objEventoBorderoOutros_evSelecao

    'seta o objBorderoOutros com os dados do obj recebido por parâmetro
    Set objBorderoOutros = obj1
    
    'preenche a tela
    lErro = Traz_BorderoOutros_Tela(objBorderoOutros)
    If lErro <> SUCESSO Then gError 108068
    
    Me.Show
    
    Exit Sub

Erro_objEventoBorderoOutros_evSelecao:
    
    Select Case gErr
    
        Case 108068
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROOUTROS_NAOENCONTRADO", gErr, objBorderoOutros.iFilialEmpresa, objBorderoOutros.lNumBordero)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143755)

    End Select
    
    Exit Sub

End Sub

Public Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim objBorderoOutros As New ClassBorderoOutros

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderoOutros"
    
    'preenche o obj com os dados da tela
    Call Move_Tela_Memoria(objBorderoOutros)
    
    'preenche a coleção de campos-valor
    colCampoValor.Add "FilialEmpresa", objBorderoOutros.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumBordero", objBorderoOutros.lNumBordero, 0, "NumBordero"
    colCampoValor.Add "AdmMeioPagto", objBorderoOutros.iAdmMeioPagto, 0, "AdmMeioPagto"
    colCampoValor.Add "Parcelamento", objBorderoOutros.iParcelamento, 0, "Parcelamento"
    colCampoValor.Add "DataEnvio", objBorderoOutros.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataImpressao", objBorderoOutros.dtDataImpressao, 0, "DataImpressao"
    colCampoValor.Add "DataBackoffice", objBorderoOutros.dtDataBackoffice, 0, "DataBackoffice"
    colCampoValor.Add "Valor", objBorderoOutros.dValor, 0, "Valor"
    colCampoValor.Add "NumIntDocCPR", objBorderoOutros.lNumIntDocCPR, 0, "NumIntDocCPR"

    'estabelece os filtros
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO
    
    Exit Function

Erro_Tela_Extrai:
    
    Tela_Extrai = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143756)

    End Select
    
    Exit Function

End Function

Public Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objBorderoOutros As New ClassBorderoOutros

On Error GoTo Erro_Tela_Preenche

    'preenche os dados necessários para um borderoOutros ser encontrado
    objBorderoOutros.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objBorderoOutros.lNumBordero = colCampoValor.Item("NumBordero").vValor
    objBorderoOutros.dtDataBackoffice = colCampoValor.Item("DataBackoffice").vValor
    objBorderoOutros.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objBorderoOutros.dtDataImpressao = colCampoValor.Item("DataImpressao").vValor
    objBorderoOutros.dValor = colCampoValor.Item("Valor").vValor
    objBorderoOutros.iAdmMeioPagto = colCampoValor.Item("AdmMeioPagto").vValor
    objBorderoOutros.lNumIntDocCPR = colCampoValor.Item("NumIntDocCpr").vValor
    objBorderoOutros.iParcelamento = colCampoValor.Item("Parcelamento").vValor
    
    'traz o bordero para a tela
    lErro = Traz_BorderoOutros_Tela(objBorderoOutros)
    If lErro <> SUCESSO Then gError 108069
    
    Tela_Preenche = SUCESSO
    
    Exit Function
    
Erro_Tela_Preenche:

    Tela_Preenche = gErr
    
    Select Case gErr
    
        Case 108069
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143757)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_BotaoProxNum_Click

    'gera o próximo número de borderô
    lErro = BorderoOutros_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 108070
    
    'coloca o código na tela
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:
    
    Select Case gErr
        
        Case 108070

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143758)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Codigo_Validate

    'se o codigo estiver em branco-> sai
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    
    'critica o código digitado
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 108073
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 108073
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143759)

    End Select
    
    Exit Sub

End Sub

Private Sub DataEnvio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)

End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvio_Validate

    'se a data não estiver preenchida-> sai
    If Len(Trim(DataEnvio.ClipText)) = 0 Then Exit Sub
    
    'critica a data
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then gError 108074

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108074
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143760)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'diminui a data
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 108076
    
    Exit Sub
    
Erro_UpDownDataEnvio_DownClick:
    
    Select Case gErr
    
        Case 108076
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143761)

    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'diminui a data
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 108077
    
    Exit Sub
    
Erro_UpDownDataEnvio_UpClick:
    
    Select Case gErr
    
        Case 108077
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143762)

    End Select
    
    Exit Sub

End Sub

Private Sub AdmMeioPagto_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_AdmMeioPagto_Click
    
    If AdmMeioPagto.ListIndex <> -1 Then
    
    'se o codigo atual for diferente do anterior
    If Codigo_Extrai(AdmMeioPagto.Text) <> giAdmMeioPagtoVelho Then
    
        'carrega a combo de parcelamento
        lErro = Carrega_Parcelamento(Codigo_Extrai(AdmMeioPagto.Text))
        If lErro <> SUCESSO Then gError 105798
        
        'guarda o código velho
        giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)
    
        giParcelamentoVelho = 0
    
    End If
    
    
    End If
    
    Exit Sub
    
Erro_AdmMeioPagto_Click:

    Select Case gErr
    
        Case 105798
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143763)

    End Select
    
    Exit Sub
    
End Sub

Private Sub AdmMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AdmMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_AdmMeioPagto_Validate

    'se a combo está preenchida
    If Len(Trim(AdmMeioPagto.Text)) <> 0 Then
    
        'se o item não foi selecionado na lista
        If AdmMeioPagto.ListIndex = -1 Then
        
            'tenta selecionar na combo
            lErro = Combo_Seleciona(AdmMeioPagto, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 108078
            
            'se não encontrou pelo código-> erro
            If lErro = 6730 Then gError 108079
            
            'se não encontrou pelo nomereduzido-> erro
            If lErro = 6731 Then gError 108080
        
        End If
            
        'se o codigo atual for diferente do anterior
        If Codigo_Extrai(AdmMeioPagto.Text) <> giAdmMeioPagtoVelho Then
        
            'carrega a combo de parcelamento
            lErro = Carrega_Parcelamento(Codigo_Extrai(AdmMeioPagto.Text))
            If lErro <> SUCESSO Then gError 108081
            
            'guarda o código velho
            giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)
        
        End If
        
    Else
        
        Parcelamento.Text = ""
        
        'limpa a combo de parcelamentos
        Parcelamento.Clear
    
        'limpa a label de total
        LabelTotal.Caption = Format(0, "STANDARD")
    
    End If
    
    giParcelamentoVelho = 0
    
    Exit Sub

Erro_AdmMeioPagto_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108078, 108081
        
        Case 108079
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, Codigo_Extrai(AdmMeioPagto.Text))
        
        Case 108080
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmMeioPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143764)

    End Select
    
    Exit Sub

End Sub

Private Sub AdmMeioPagto_GotFocus()
    
    'guarda o código velho
    giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)

End Sub

Private Function Testa_Alteracao_Parcelamento()

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objBorderoOutros As New ClassBorderoOutros

On Error GoTo Erro_Testa_Alteracao_Parcelamento

    'se estiver mudando o parcelamento
    If giParcelamentoVelho <> Codigo_Extrai(Parcelamento.Text) Then
        
        'preenche um admmeiopagtocondpagto para buscar seu saldo
        objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagtoCondPagto.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text)
        objAdmMeioPagtoCondPagto.iParcelamento = Codigo_Extrai(Parcelamento.Text)
        
        'tenta buscar na tabela admmeiopagtocondpagto
        lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
        If lErro <> SUCESSO And lErro <> 107297 Then gError 108082
        
        'se não encontrar->erro
        If lErro = 107297 Then gError 108083
        
        'preenche o saldo
        LabelTotal.Caption = Format(objAdmMeioPagtoCondPagto.dSaldo, "STANDARD")
        
        'preenche um objBorderoOutros para a busca
        objBorderoOutros.iFilialEmpresa = giFilialEmpresa
        objBorderoOutros.lNumBordero = StrParaLong(Codigo.Text)
        
        'lê um borderooutros
        lErro = CF("BorderoOutros_Le", objBorderoOutros)
        If lErro <> SUCESSO And lErro <> 108059 Then gError 108084
        
        'se encontrou e o admmeiopagto e o parcelamento batem com os q estão na tela, atualizar o saldo
        If lErro = SUCESSO _
        And objBorderoOutros.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text) _
        And objBorderoOutros.iParcelamento = Codigo_Extrai(Parcelamento.Text) Then
            LabelTotal.Caption = Format(StrParaDbl(LabelTotal.Caption) + objBorderoOutros.dValor, "STANDARD")
        End If
        
        'guarda o parcelamento velho
        giParcelamentoVelho = Codigo_Extrai(Parcelamento.Text)
        
    End If
    
    Testa_Alteracao_Parcelamento = SUCESSO
    
    Exit Function
    
Erro_Testa_Alteracao_Parcelamento:
    
    Testa_Alteracao_Parcelamento = gErr
    
    Select Case gErr
    
        Case 108082
        
        Case 108083
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_ADMMEIOPAGTO_NAOENCONTRADO", gErr, objAdmMeioPagtoCondPagto.iParcelamento, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto)
            
        Case 108084
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143765)

    End Select
    
    Exit Function

End Function

Private Sub Parcelamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcelamento_Click()

Dim lErro As Long

On Error GoTo Erro_Parcelamento_Click

    lErro = Testa_Alteracao_Parcelamento()
    If lErro <> SUCESSO Then gError 108085
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Parcelamento_Click:

    Select Case gErr
        
        Case 108085
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143766)

    End Select
    
    Exit Sub

End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Parcelamento_Validate

    'se o parcelamento estiver prenchido
    If Len(Trim(Parcelamento.Text)) <> 0 Then
    
        'se o parcelamento for diferente do último selecionado
        If Parcelamento.ListIndex = -1 Then
            
            'tenta selecionar
            lErro = Combo_Seleciona(Parcelamento, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 108086
            
            'se não encontrar pelo código-> erro
            If lErro = 6730 Then gError 108087
            
            'se não encontrar pelo nomereduzido-> erro
            If lErro = 6731 Then gError 108088
        
        End If
        
        lErro = Testa_Alteracao_Parcelamento()
        If lErro <> SUCESSO Then gError 108089
    
    'se não estiver preenchido
    Else
    
        'limpa a label total
        LabelTotal.Caption = Format(0, "STANDARD")
        
    End If
    
    Exit Sub

Erro_Parcelamento_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 108089, 108086
        
        Case 108087
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAOENCONTRADO", gErr, iCodigo)
        
        Case 108088
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAOENCONTRADO", gErr, Parcelamento.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143767)

    End Select

    Exit Sub

End Sub

Private Sub ValorEnviar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviar_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorEnviar_Validate

    If Len(Trim(ValorEnviar.Text)) = 0 Then Exit Sub

    'Critica o Valor digitado
    lErro = Valor_NaoNegativo_Critica(ValorEnviar.Text)
    If lErro <> SUCESSO Then gError 108090
    
    Exit Sub

Erro_ValorEnviar_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108090
    
        Case 108091
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORENVIAR_VALORDISPONIVEL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143768)

    End Select

    Exit Sub

End Sub

Private Sub ValorEnviarN_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviarN_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorEnviarN_Validate
    
    If Len(Trim(ValorEnviarN.Text)) = 0 Then Exit Sub

    'Critica o Valor digitado
    lErro = Valor_NaoNegativo_Critica(ValorEnviarN.Text)
    If lErro <> SUCESSO Then gError 108092
    
    Exit Sub

Erro_ValorEnviarN_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108092
        
        Case 108093
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORENVIAR_VALORDISPONIVEL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143769)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 108099
    
    Call Limpa_Tela_BorderoOutros
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
        
        Case 108099
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143770)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBorderoOutros As New ClassBorderoOutros

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass

    'se estiver no bo->erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 108100
    
    'se o código estiver vazio-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 108101
    
    'se a data estiver em branco-> erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 108102
    
    'se a admmeiopagto estiver em branco-> erro
    If AdmMeioPagto.ListIndex = -1 Then gError 108103
    
    'se o parcelamento estiver em branco-> erro
    If Parcelamento.ListIndex = -1 Then gError 108104
    
    'se a soma dos valores especificados e não especificados foir igual a 0-> erro
    If StrParaDbl(ValorEnviar.Text) + StrParaDbl(ValorEnviarN.Text) = 0 Then gError 108105
    
    'preenche o bordero com os dados da tela
    Call Move_Tela_Memoria(objBorderoOutros)
    
    'testa se é uma alteração
    lErro = Trata_Alteracao(objBorderoOutros, objBorderoOutros.iFilialEmpresa, objBorderoOutros.lNumBordero)
    If lErro <> SUCESSO Then gError 108106
    
    'grava o bordero outros
    lErro = CF("BorderoOutros_Grava", objBorderoOutros)
    If lErro <> SUCESSO Then gError 108107

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 108100
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROOUTROS_GRAVACAO_BACKOFFICE", gErr)
        
        Case 108101
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 108102
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 108103
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_SELECIONADO", gErr)
        
        Case 108104
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO1", gErr)
        
        Case 108105
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROOUTROS_ZERADO", gErr)
        
        Case 108107, 108106
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143771)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objBorderoOutros As New ClassBorderoOutros

On Error GoTo Erro_BotaoExcluir_Click

    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 108176
    
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_BORDEROOUTROS", giFilialEmpresa, Codigo.Text)
    
    If vbMsgResp = vbYes Then
    
        'preenche os atributos necessários à exclusão do bordero
        objBorderoOutros.iFilialEmpresa = giFilialEmpresa
        objBorderoOutros.lNumBordero = StrParaLong(Codigo.Text)
        
        'exclui o borderÔ
        lErro = CF("BorderoOutros_Exclui", objBorderoOutros)
        If lErro <> SUCESSO Then gError 108177
        
        Call Limpa_Tela_BorderoOutros
        
        iAlterado = 0
    
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:
    
    Select Case gErr
    
        Case 108176
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 108177

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143772)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro

On Error GoTo Erro_Botaolimpar_Click
    
    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 108096
    
    Call Limpa_Tela_BorderoOutros
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 108097
    
    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:
    
    Select Case gErr
        
        Case 108096, 108097
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143773)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoOutros()

Dim lErro As Long
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Limpa_Tela_BorderoOutros

    Call Limpa_Tela(Me)

    LabelTotal.Caption = Format(0, "STANDARD")
    
    'preenche um tmplojafilial para ler o seu saldo
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_OUTROS
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'le o seu saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108067
    
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    'preenche a data com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataEnvio.PromptInclude = True
    
    AdmMeioPagto.ListIndex = -1
    Parcelamento.Clear
    Parcelamento.Text = ""
    
    giAdmMeioPagtoVelho = 0
    giParcelamentoVelho = 0

    Exit Sub

Erro_Limpa_Tela_BorderoOutros:

    Select Case gErr
    
        Case 108067
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143774)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub form_unload(Cancel As Integer)

    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

    'libera a memória
    Set objEventoBorderoOutros = Nothing
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function Carrega_Parcelamento(iCodigo As Integer) As Long

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_Parcelamento

    'preenche os atributos para buscar a admmeiopagtocondpagto da admmeiopagto
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    objAdmMeioPagto.iCodigo = iCodigo

    'busca no BD e preenche colcondpagtoloja com os parcelamentos
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 108094
    
    'se não encontrar-> erro
    If lErro = 104086 Then gError 108095
    
    Parcelamento.Text = ""
    
    'limpa a combo
    Parcelamento.Clear
    
    'preenche a combo com os novos valores
    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
    
        Parcelamento.AddItem (objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento)
        Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
    
    Next

    LabelTotal.Caption = ""

    Carrega_Parcelamento = SUCESSO
    
    Exit Function

Erro_Carrega_Parcelamento:

    Carrega_Parcelamento = gErr
    
    Select Case gErr
    
        Case 108094
        
        Case 108095
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTOS_ADMMEIOPAGTO_NAOENCONTRADOS", gErr, objAdmMeioPagto.iFilialEmpresa, objAdmMeioPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143775)

    End Select
    
    Exit Function

End Function

Public Function Traz_BorderoOutros_Tela(objBorderoOutros As ClassBorderoOutros) As Long

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Traz_BorderoOutros_Tela

    Call Limpa_Tela_BorderoOutros

    'preenche a tela
    Codigo.Text = objBorderoOutros.lNumBordero
    DataEnvio.Text = Format(objBorderoOutros.dtDataEnvio, "dd/mm/yy")
    
    AdmMeioPagto.Text = objBorderoOutros.iAdmMeioPagto
    Call AdmMeioPagto_Validate(bSGECancelDummy)
    
    Parcelamento.Text = objBorderoOutros.iParcelamento
    Call Parcelamento_Validate(bSGECancelDummy)
        
    ValorEnviar.Text = Format(objBorderoOutros.dValor, "STANDARD")
    
    'preenche os atributos necessários para buscar uma admmeiopagtocondpagto específica na referida tabela
    objAdmMeioPagtoCondPagto.iAdmMeioPagto = objBorderoOutros.iAdmMeioPagto
    objAdmMeioPagtoCondPagto.iParcelamento = objBorderoOutros.iParcelamento
    objAdmMeioPagtoCondPagto.iFilialEmpresa = objBorderoOutros.iFilialEmpresa
    
    'tenta encontrar o admmeiopagtocondpagto que atenda às condições acima
    lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
    If lErro <> SUCESSO And lErro <> 107297 Then gError 108062
    
    'se não encontrar-> erro
    If lErro = 107297 Then gError 108063
    
    'preenche o total
    LabelTotal.Caption = Format((objAdmMeioPagtoCondPagto.dSaldo + objBorderoOutros.dValor), "STANDARD")
    
    'preenche os atributos necessários para buscar um tmplojafilial
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_OUTROS
    
    'lê o saldo na tabela de não especificados
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 108064
    
    'preenche o total não especificado
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    iAlterado = 0

    Traz_BorderoOutros_Tela = SUCESSO

    Exit Function

Erro_Traz_BorderoOutros_Tela:

    Traz_BorderoOutros_Tela = gErr

    Select Case gErr
    
        Case 108062, 108064
        
        Case 108063
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_ADMMEIOPAGTO_NAOENCONTRADO", gErr, objAdmMeioPagtoCondPagto.iParcelamento, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143776)

    End Select

    Exit Function

End Function

Private Function Carrega_Outros()

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Carrega_Outros

    'busca a admmeiopagto com tipo=TIPOMEIOPAGTOLOJA_OUTROS
    lErro = CF("AdmMeioPagto_Le_TipoMeioPagto", TIPOMEIOPAGTOLOJA_OUTROS, colAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 107360 Then gError 108050

    'se não encontrar-> erro
    If lErro = 107360 Then gError 108051

    'preenche a combo de admmeiopagto
    For Each objAdmMeioPagto In colAdmMeioPagto
        
        If objAdmMeioPagto.iCodigo <> MEIO_PAGAMENTO_CONTRAVALE Then

            AdmMeioPagto.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
            AdmMeioPagto.ItemData(AdmMeioPagto.NewIndex) = objAdmMeioPagto.iCodigo
        End If
    Next

    Carrega_Outros = SUCESSO

    Exit Function

Erro_Carrega_Outros:

    Carrega_Outros = gErr

    Select Case gErr

        Case 108050

        Case 108051
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_TIPOMEIOPAGTO_VAZIA", gErr, TIPOMEIOPAGTOLOJA_OUTROS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143777)

    End Select

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objBorderoOutros As ClassBorderoOutros)

On Error GoTo Erro_Move_Tela_Memoria

    'preenche os atributos de um borderoOutros
    objBorderoOutros.lNumBordero = StrParaLong(Codigo.Text)
    objBorderoOutros.dtDataEnvio = StrParaDate(DataEnvio.Text)
    objBorderoOutros.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text)
    objBorderoOutros.iParcelamento = Codigo_Extrai(Parcelamento.Text)
    objBorderoOutros.dValor = StrParaDbl(ValorEnviar.Text)
    objBorderoOutros.dValorN = StrParaDbl(ValorEnviarN.Text)
    objBorderoOutros.dtDataBackoffice = DATA_NULA
    objBorderoOutros.dtDataImpressao = DATA_NULA
    objBorderoOutros.iFilialEmpresa = giFilialEmpresa
    objBorderoOutros.sAdmMeioPagto = Nome_Extrai(AdmMeioPagto.Text)
    objBorderoOutros.sNomeParcelamento = Nome_Extrai(Parcelamento.Text)

    Exit Sub

Erro_Move_Tela_Memoria:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143778)

    End Select

    Exit Sub

End Sub

Private Function BorderoOutros_Codigo_Automatico(lCodigo As Long) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoOutros_Codigo_Automatico

    'busca o próximo número de borderô automático
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_BORDEROOUTROS", "BorderoOutros", "NumBordero", lCodigo)
    If lErro <> SUCESSO Then gError 108071
    
    BorderoOutros_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_BorderoOutros_Codigo_Automatico:
    
    BorderoOutros_Codigo_Automatico = gErr
    
    Select Case gErr
        
        Case 108071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143779)

    End Select

    Exit Function

End Function

Private Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Trim(Mid(sTexto, iPosicao + 1))

    Nome_Extrai = sString

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Borderô Outros"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoOutros"

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

Private Sub BotaoProxNum_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ActiveControl = Me.ActiveControl

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
